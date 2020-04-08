// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public class FieldRetriever
    {
        public static string InstrText(XElement root, int id)
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            var cachedAnnotationInformation = root.Annotation<Dictionary<int, List<XElement>>>();
            if (cachedAnnotationInformation == null)
            {
                throw new OpenXmlPowerToolsException("Internal error");
            }

            // it is possible that a field code contains no instr text
            if (!cachedAnnotationInformation.ContainsKey(id))
            {
                return "";
            }

            var relevantElements = cachedAnnotationInformation[id];

            var groupedSubFields = relevantElements
                .GroupAdjacent(e =>
                {
                    var s = e.Annotation<Stack<FieldElementTypeInfo>>();
                    var stackElement = s.FirstOrDefault(z => z.Id == id);
                    var elementsBefore = s.TakeWhile(z => z != stackElement);
                    return elementsBefore.Any();
                })
                .ToList();

            var instrText = groupedSubFields
                .Select(g =>
                {
                    if (g.Key == false)
                    {
                        return g.Select(e =>
                        {
                            var s = e.Annotation<Stack<FieldElementTypeInfo>>();
                            var stackElement = s.FirstOrDefault(z => z.Id == id);
                            if (stackElement.FieldElementType == FieldElementTypeEnum.InstrText &&
                                e.Name == w + "instrText")
                            {
                                return e.Value;
                            }

                            return "";
                        })
                            .StringConcatenate();
                    }
                    else
                    {
                        var s = g.First().Annotation<Stack<FieldElementTypeInfo>>();
                        var stackElement = s.FirstOrDefault(z => z.Id == id);
                        var elementBefore = s.TakeWhile(z => z != stackElement).Last();
                        var subFieldId = elementBefore.Id;
                        return InstrText(root, subFieldId);
                    }
                })
                .StringConcatenate();

            return "{" + instrText + "}";
        }

        public static void AnnotateWithFieldInfo(OpenXmlPart part)
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            var root = part.GetXDocument().Root;
            var r = root.DescendantsAndSelf()
                .Rollup(
                    new FieldElementTypeStack
                    {
                        Id = 0,
                        FiStack = null,
                    },
                    (e, s) =>
                    {
                        if (e.Name == w + "fldChar")
                        {
                            if (e.Attribute(w + "fldCharType").Value == "begin")
                            {
                                Stack<FieldElementTypeInfo> fis;
                                if (s.FiStack == null)
                                {
                                    fis = new Stack<FieldElementTypeInfo>();
                                }
                                else
                                {
                                    fis = new Stack<FieldElementTypeInfo>(s.FiStack.Reverse());
                                }

                                fis.Push(
                                    new FieldElementTypeInfo
                                    {
                                        Id = s.Id + 1,
                                        FieldElementType = FieldElementTypeEnum.Begin,
                                    });
                                return new FieldElementTypeStack
                                {
                                    Id = s.Id + 1,
                                    FiStack = fis,
                                };
                            };
                            if (e.Attribute(w + "fldCharType").Value == "separate")
                            {
                                var fis = new Stack<FieldElementTypeInfo>(s.FiStack.Reverse());
                                var wfi = fis.Pop();
                                fis.Push(
                                    new FieldElementTypeInfo
                                    {
                                        Id = wfi.Id,
                                        FieldElementType = FieldElementTypeEnum.Separate,
                                    });
                                return new FieldElementTypeStack
                                {
                                    Id = s.Id,
                                    FiStack = fis,
                                };
                            }
                            if (e.Attribute(w + "fldCharType").Value == "end")
                            {
                                var fis = new Stack<FieldElementTypeInfo>(s.FiStack.Reverse());
                                var wfi = fis.Pop();
                                return new FieldElementTypeStack
                                {
                                    Id = s.Id,
                                    FiStack = fis,
                                };
                            }
                        }
                        if (s.FiStack == null || s.FiStack.Count() == 0)
                        {
                            return s;
                        }

                        var wfi3 = s.FiStack.Peek();
                        if (wfi3.FieldElementType == FieldElementTypeEnum.Begin)
                        {
                            var fis = new Stack<FieldElementTypeInfo>(s.FiStack.Reverse());
                            var wfi2 = fis.Pop();
                            fis.Push(
                                new FieldElementTypeInfo
                                {
                                    Id = wfi2.Id,
                                    FieldElementType = FieldElementTypeEnum.InstrText,
                                });
                            return new FieldElementTypeStack
                            {
                                Id = s.Id,
                                FiStack = fis,
                            };
                        }
                        if (wfi3.FieldElementType == FieldElementTypeEnum.Separate)
                        {
                            var fis = new Stack<FieldElementTypeInfo>(s.FiStack.Reverse());
                            var wfi2 = fis.Pop();
                            fis.Push(
                                new FieldElementTypeInfo
                                {
                                    Id = wfi2.Id,
                                    FieldElementType = FieldElementTypeEnum.Result,
                                });
                            return new FieldElementTypeStack
                            {
                                Id = s.Id,
                                FiStack = fis,
                            };
                        }
                        if (wfi3.FieldElementType == FieldElementTypeEnum.End)
                        {
                            var fis = new Stack<FieldElementTypeInfo>(s.FiStack.Reverse());
                            fis.Pop();
                            if (!fis.Any())
                            {
                                fis = null;
                            }

                            return new FieldElementTypeStack
                            {
                                Id = s.Id,
                                FiStack = fis,
                            };
                        }
                        return s;
                    });
            var elementPlusInfo = root.DescendantsAndSelf().PtZip(r, (t1, t2) =>
            {
                return new
                {
                    Element = t1,
                    Id = t2.Id,
                    WmlFieldInfoStack = t2.FiStack,
                };
            });
            foreach (var item in elementPlusInfo)
            {
                if (item.WmlFieldInfoStack != null)
                {
                    item.Element.AddAnnotation(item.WmlFieldInfoStack);
                }
            }

            //This code is useful when you want to take a look at the annotations, making sure that they are made correctly.
            //
            //foreach (var desc in root.DescendantsAndSelf())
            //{
            //    Stack<FieldElementTypeInfo> s = desc.Annotation<Stack<FieldElementTypeInfo>>();
            //    if (s != null)
            //    {
            //        Console.WriteLine(desc.Name.LocalName.PadRight(20));
            //        foreach (var item in s)
            //        {
            //            Console.WriteLine("    {0:0000} {1}", item.Id, item.FieldElementType.ToString());
            //            Console.ReadKey();
            //        }
            //    }
            //}

            var cachedAnnotationInformation = new Dictionary<int, List<XElement>>();
            foreach (var desc in root.DescendantsTrimmed(d => d.Name == W.rPr || d.Name == W.pPr))
            {
                var s = desc.Annotation<Stack<FieldElementTypeInfo>>();

                if (s != null)
                {
                    foreach (var item in s)
                    {
                        if (item.FieldElementType == FieldElementTypeEnum.InstrText)
                        {
                            if (cachedAnnotationInformation.ContainsKey(item.Id))
                            {
                                cachedAnnotationInformation[item.Id].Add(desc);
                            }
                            else
                            {
                                cachedAnnotationInformation.Add(item.Id, new List<XElement>() { desc });
                            }
                        }
                    }
                }
            }
            root.AddAnnotation(cachedAnnotationInformation);
        }

        private enum State
        {
            InToken,
            InWhiteSpace,
            InQuotedToken,
            OnOpeningQuote,
            OnClosingQuote,
            OnBackslash,
        }

        private static string[] GetTokens(string field)
        {
            var state = State.InWhiteSpace;
            var tokenStart = 0;
            var quoteStart = char.MinValue;
            var tokens = new List<string>();
            for (var c = 0; c < field.Length; c++)
            {
                if (char.IsWhiteSpace(field[c]))
                {
                    if (state == State.InToken)
                    {
                        tokens.Add(field.Substring(tokenStart, c - tokenStart));
                        state = State.InWhiteSpace;
                        continue;
                    }
                    if (state == State.OnOpeningQuote)
                    {
                        tokenStart = c;
                        state = State.InQuotedToken;
                    }
                    if (state == State.OnClosingQuote)
                    {
                        state = State.InWhiteSpace;
                    }

                    continue;
                }
                if (field[c] == '\\')
                {
                    if (state == State.InQuotedToken)
                    {
                        state = State.OnBackslash;
                        continue;
                    }
                }
                if (state == State.OnBackslash)
                {
                    state = State.InQuotedToken;
                    continue;
                }
                if (field[c] == '"' || field[c] == '\'' || field[c] == 0x201d)
                {
                    if (state == State.InWhiteSpace)
                    {
                        quoteStart = field[c];
                        state = State.OnOpeningQuote;
                        continue;
                    }
                    if (state == State.InQuotedToken)
                    {
                        if (field[c] == quoteStart)
                        {
                            tokens.Add(field.Substring(tokenStart, c - tokenStart));
                            state = State.OnClosingQuote;
                            continue;
                        }
                        continue;
                    }
                    if (state == State.OnOpeningQuote)
                    {
                        if (field[c] == quoteStart)
                        {
                            state = State.OnClosingQuote;
                            continue;
                        }
                        else
                        {
                            tokenStart = c;
                            state = State.InQuotedToken;
                            continue;
                        }
                    }
                    continue;
                }
                if (state == State.InWhiteSpace)
                {
                    tokenStart = c;
                    state = State.InToken;
                    continue;
                }
                if (state == State.OnOpeningQuote)
                {
                    tokenStart = c;
                    state = State.InQuotedToken;
                    continue;
                }
                if (state == State.OnClosingQuote)
                {
                    tokenStart = c;
                    state = State.InToken;
                    continue;
                }
            }
            if (state == State.InToken)
            {
                tokens.Add(field.Substring(tokenStart, field.Length - tokenStart));
            }

            return tokens.ToArray();
        }

        public static FieldInfo ParseField(string field)
        {
            var emptyField = new FieldInfo
            {
                FieldType = "",
                Arguments = new string[] { },
                Switches = new string[] { },
            };

            if (field.Length == 0)
            {
                return emptyField;
            }

            var fieldType = field.TrimStart().Split(' ').FirstOrDefault();
            if (fieldType == null)
            {
                return emptyField;
            }

            if (fieldType.ToUpper() != "HYPERLINK" &&
                fieldType.ToUpper() != "REF" &&
                fieldType.ToUpper() != "SEQ" &&
                fieldType.ToUpper() != "STYLEREF")
            {
                return emptyField;
            }

            var tokens = GetTokens(field);
            if (tokens.Length == 0)
            {
                return emptyField;
            }

            var fieldInfo = new FieldInfo()
            {
                FieldType = tokens[0],
                Switches = tokens.Where(t => t[0] == '\\').ToArray(),
                Arguments = tokens.Skip(1).Where(t => t[0] != '\\').ToArray(),
            };
            return fieldInfo;
        }

        public class FieldInfo
        {
            public string FieldType { get; set; }
            public string[] Switches { get; set; }
            public string[] Arguments { get; set; }
        }

        public enum FieldElementTypeEnum
        {
            Begin,
            InstrText,
            Separate,
            Result,
            End,
        };

        public class FieldElementTypeInfo
        {
            public int Id { get; set; }
            public FieldElementTypeEnum FieldElementType { get; set; }
        }

        public class FieldElementTypeStack
        {
            public int Id { get; set; }
            public Stack<FieldElementTypeInfo> FiStack { get; set; }
        }
    }
}