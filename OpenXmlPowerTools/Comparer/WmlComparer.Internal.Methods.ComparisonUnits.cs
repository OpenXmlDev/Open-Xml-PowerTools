// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public static partial class WmlComparer
    {
        #region CreateComparisonUnitAtomList

        internal static ComparisonUnitAtom[] CreateComparisonUnitAtomList(
            OpenXmlPart part,
            XElement contentParent,
            WmlComparerSettings settings)
        {
            VerifyNoInvalidContent(contentParent);
            AssignUnidToAllElements(contentParent); // add the Guid id to every element
            MoveLastSectPrIntoLastParagraph(contentParent);
            ComparisonUnitAtom[] cal = CreateComparisonUnitAtomListInternal(part, contentParent, settings).ToArray();

            if (False)
            {
                var sb = new StringBuilder();
                foreach (ComparisonUnitAtom item in cal)
                    sb.Append(item + Environment.NewLine);
                string sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            return cal;
        }

        private static void VerifyNoInvalidContent(XElement contentParent)
        {
            XElement invalidElement = contentParent.Descendants().FirstOrDefault(d => InvalidElements.Contains(d.Name));
            if (invalidElement == null)
                return;

            throw new NotSupportedException("Document contains " + invalidElement.Name.LocalName);
        }

        private static void MoveLastSectPrIntoLastParagraph(XElement contentParent)
        {
            List<XElement> lastSectPrList = contentParent.Elements(W.sectPr).ToList();
            if (lastSectPrList.Count() > 1)
                throw new OpenXmlPowerToolsException("Invalid document");

            XElement lastSectPr = lastSectPrList.FirstOrDefault();
            if (lastSectPr != null)
            {
                XElement lastParagraph = contentParent.Elements(W.p).LastOrDefault();
                if (lastParagraph == null)
                    throw new OpenXmlPowerToolsException("Invalid document");

                XElement pPr = lastParagraph.Element(W.pPr);
                if (pPr == null)
                {
                    pPr = new XElement(W.pPr);
                    lastParagraph.AddFirst(W.pPr);
                }

                pPr.Add(lastSectPr);
                contentParent.Elements(W.sectPr).Remove();
            }
        }

        private static List<ComparisonUnitAtom> CreateComparisonUnitAtomListInternal(
            OpenXmlPart part,
            XElement contentParent,
            WmlComparerSettings settings)
        {
            var comparisonUnitAtomList = new List<ComparisonUnitAtom>();
            CreateComparisonUnitAtomListRecurse(part, contentParent, comparisonUnitAtomList, settings);
            return comparisonUnitAtomList;
        }

        private static void CreateComparisonUnitAtomListRecurse(
            OpenXmlPart part,
            XElement element,
            List<ComparisonUnitAtom> comparisonUnitAtomList,
            WmlComparerSettings settings)
        {
            if (element.Name == W.body || element.Name == W.footnote || element.Name == W.endnote)
            {
                foreach (XElement item in element.Elements())
                    CreateComparisonUnitAtomListRecurse(part, item, comparisonUnitAtomList, settings);
                return;
            }

            if (element.Name == W.p)
            {
                IEnumerable<XElement> paraChildrenToProcess = element
                    .Elements()
                    .Where(e => e.Name != W.pPr);
                foreach (XElement item in paraChildrenToProcess)
                    CreateComparisonUnitAtomListRecurse(part, item, comparisonUnitAtomList, settings);
                XElement paraProps = element.Element(W.pPr);
                if (paraProps == null)
                {
                    var pPrComparisonUnitAtom = new ComparisonUnitAtom(
                        new XElement(W.pPr),
                        element.AncestorsAndSelf()
                            .TakeWhile(a => a.Name != W.body && a.Name != W.footnotes && a.Name != W.endnotes).Reverse()
                            .ToArray(),
                        part,
                        settings);
                    comparisonUnitAtomList.Add(pPrComparisonUnitAtom);
                }
                else
                {
                    var pPrComparisonUnitAtom = new ComparisonUnitAtom(
                        paraProps,
                        element.AncestorsAndSelf()
                            .TakeWhile(a => a.Name != W.body && a.Name != W.footnotes && a.Name != W.endnotes).Reverse()
                            .ToArray(),
                        part,
                        settings);
                    comparisonUnitAtomList.Add(pPrComparisonUnitAtom);
                }

                return;
            }

            if (element.Name == W.r)
            {
                IEnumerable<XElement> runChildrenToProcess = element
                    .Elements()
                    .Where(e => e.Name != W.rPr);
                foreach (XElement item in runChildrenToProcess)
                    CreateComparisonUnitAtomListRecurse(part, item, comparisonUnitAtomList, settings);
                return;
            }

            if (element.Name == W.t || element.Name == W.delText)
            {
                string val = element.Value;
                foreach (char ch in val)
                {
                    var sr = new ComparisonUnitAtom(
                        new XElement(element.Name, ch),
                        element.AncestorsAndSelf()
                            .TakeWhile(a => a.Name != W.body && a.Name != W.footnotes && a.Name != W.endnotes).Reverse()
                            .ToArray(),
                        part,
                        settings);
                    comparisonUnitAtomList.Add(sr);
                }

                return;
            }

            if (AllowableRunChildren.Contains(element.Name) || element.Name == W._object)
            {
                var sr3 = new ComparisonUnitAtom(
                    element,
                    element.AncestorsAndSelf().TakeWhile(a => a.Name != W.body && a.Name != W.footnotes && a.Name != W.endnotes)
                        .Reverse().ToArray(),
                    part,
                    settings);
                comparisonUnitAtomList.Add(sr3);
                return;
            }

            RecursionInfo re = RecursionElements.FirstOrDefault(z => z.ElementName == element.Name);
            if (re != null)
            {
                AnnotateElementWithProps(part, element, comparisonUnitAtomList, re.ChildElementPropertyNames, settings);
                return;
            }

            if (ElementsToThrowAway.Contains(element.Name))
                return;

            AnnotateElementWithProps(part, element, comparisonUnitAtomList, null, settings);
        }

        private static void AnnotateElementWithProps(
            OpenXmlPart part,
            XElement element,
            List<ComparisonUnitAtom> comparisonUnitAtomList,
            XName[] childElementPropertyNames,
            WmlComparerSettings settings)
        {
            IEnumerable<XElement> runChildrenToProcess;
            if (childElementPropertyNames == null)
                runChildrenToProcess = element.Elements();
            else
                runChildrenToProcess = element
                    .Elements()
                    .Where(e => !childElementPropertyNames.Contains(e.Name));

            foreach (XElement item in runChildrenToProcess)
                CreateComparisonUnitAtomListRecurse(part, item, comparisonUnitAtomList, settings);
        }

        #endregion CreateComparisonUnitAtomList

        #region GetComparisonUnitList

        // The following method must be made internal if we ever turn this part of the partial class
        // into its own class.
        private static ComparisonUnit[] GetComparisonUnitList(
            ComparisonUnitAtom[] comparisonUnitAtomList,
            WmlComparerSettings settings)
        {
            var seed = new Atgbw
            {
                Key = null,
                ComparisonUnitAtomMember = null,
                NextIndex = 0
            };

            IEnumerable<Atgbw> groupingKey = comparisonUnitAtomList
                .Rollup(seed, (sr, prevAtgbw, i) =>
                {
                    int? key;
                    int nextIndex = prevAtgbw.NextIndex;
                    if (sr.ContentElement.Name == W.t)
                    {
                        string chr = sr.ContentElement.Value;
                        char ch = chr[0];
                        if (ch == '.' || ch == ',')
                        {
                            var beforeIsDigit = false;
                            if (i > 0)
                            {
                                ComparisonUnitAtom prev = comparisonUnitAtomList[i - 1];
                                if (prev.ContentElement.Name == W.t && char.IsDigit(prev.ContentElement.Value[0]))
                                    beforeIsDigit = true;
                            }

                            var afterIsDigit = false;
                            if (i < comparisonUnitAtomList.Length - 1)
                            {
                                ComparisonUnitAtom next = comparisonUnitAtomList[i + 1];
                                if (next.ContentElement.Name == W.t && char.IsDigit(next.ContentElement.Value[0]))
                                    afterIsDigit = true;
                            }

                            if (beforeIsDigit || afterIsDigit)
                            {
                                key = nextIndex;
                            }
                            else
                            {
                                nextIndex++;
                                key = nextIndex;
                                nextIndex++;
                            }
                        }
                        else if (settings.WordSeparators.Contains(ch))
                        {
                            nextIndex++;
                            key = nextIndex;
                            nextIndex++;
                        }
                        else
                        {
                            key = nextIndex;
                        }
                    }
                    else if (WordBreakElements.Contains(sr.ContentElement.Name))
                    {
                        nextIndex++;
                        key = nextIndex;
                        nextIndex++;
                    }
                    else
                    {
                        key = nextIndex;
                    }

                    return new Atgbw
                    {
                        Key = key,
                        ComparisonUnitAtomMember = sr,
                        NextIndex = nextIndex
                    };
                })
                .ToArray();

            if (False)
            {
                var sb = new StringBuilder();
                foreach (Atgbw item in groupingKey)
                {
                    sb.Append(item.Key + Environment.NewLine);
                    sb.Append("    " + item.ComparisonUnitAtomMember.ToString(0) + Environment.NewLine);
                }

                string sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            IEnumerable<IGrouping<int?, Atgbw>> groupedByWords = groupingKey
                .GroupAdjacent(gc => gc.Key)
                .ToArray();

            if (False)
            {
                var sb = new StringBuilder();
                foreach (IGrouping<int?, Atgbw> group in groupedByWords)
                {
                    sb.Append("Group ===== " + @group.Key + Environment.NewLine);
                    foreach (Atgbw gc in @group)
                    {
                        sb.Append("    " + gc.ComparisonUnitAtomMember.ToString(0) + Environment.NewLine);
                    }
                }

                string sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            WithHierarchicalGroupingKey[] withHierarchicalGroupingKey = groupedByWords
                .Select(g =>
                    {
                        string[] hierarchicalGroupingArray = g
                            .First()
                            .ComparisonUnitAtomMember
                            .AncestorElements
                            .Where(a => ComparisonGroupingElements.Contains(a.Name))
                            .Select(a => a.Name.LocalName + ":" + (string) a.Attribute(PtOpenXml.Unid))
                            .ToArray();

                        return new WithHierarchicalGroupingKey
                        {
                            ComparisonUnitWord = new ComparisonUnitWord(g.Select(gc => gc.ComparisonUnitAtomMember)),
                            HierarchicalGroupingArray = hierarchicalGroupingArray
                        };
                    }
                )
                .ToArray();

            if (False)
            {
                var sb = new StringBuilder();
                foreach (WithHierarchicalGroupingKey group in withHierarchicalGroupingKey)
                {
                    sb.Append("Grouping Array: " +
                              @group.HierarchicalGroupingArray.Select(gam => gam + " - ").StringConcatenate() +
                              Environment.NewLine);
                    foreach (ComparisonUnit gc in @group.ComparisonUnitWord.Contents)
                    {
                        sb.Append("    " + gc.ToString(0) + Environment.NewLine);
                    }
                }

                string sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            ComparisonUnit[] cul = GetHierarchicalComparisonUnits(withHierarchicalGroupingKey, 0).ToArray();

            if (False)
            {
                string str = ComparisonUnit.ComparisonUnitListToString(cul);
                TestUtil.NotePad(str);
            }

            return cul;
        }

        private static IEnumerable<ComparisonUnit> GetHierarchicalComparisonUnits(
            IEnumerable<WithHierarchicalGroupingKey> input,
            int level)
        {
            IEnumerable<IGrouping<string, WithHierarchicalGroupingKey>> grouped = input
                .GroupAdjacent(
                    whgk => level >= whgk.HierarchicalGroupingArray.Length ? "" : whgk.HierarchicalGroupingArray[level]);

            List<ComparisonUnit> retList = grouped
                .Select(gc =>
                {
                    if (gc.Key == "")
                    {
                        return (IEnumerable<ComparisonUnit>) gc.Select(whgk => whgk.ComparisonUnitWord).ToList();
                    }

                    string[] spl = gc.Key.Split(':');
                    ComparisonUnitGroupType groupType = WmlComparerUtil.ComparisonUnitGroupTypeFromLocalName(spl[0]);
                    IEnumerable<ComparisonUnit> childHierarchicalComparisonUnits = GetHierarchicalComparisonUnits(gc, level + 1);
                    var newCompUnitGroup = new ComparisonUnitGroup(childHierarchicalComparisonUnits, groupType, level);

                    return new[] { newCompUnitGroup };
                })
                .SelectMany(m => m)
                .ToList();

            return retList;
        }

        #endregion GetComparisonUnitList
    }
}
