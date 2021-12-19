// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Xml.XPath;
using DocumentFormat.OpenXml.Packaging;
using JetBrains.Annotations;

namespace OpenXmlPowerTools
{
    [PublicAPI]
    public static class DocumentAssembler
    {
        private static readonly XName[] MetaToForceToBlock =
        {
            Pa.Conditional,
            Pa.EndConditional,
            Pa.Repeat,
            Pa.EndRepeat,
            Pa.Table,
        };

        private static readonly List<string> AliasList = new()
        {
            "Content",
            "Table",
            "Repeat",
            "EndRepeat",
            "Conditional",
            "EndConditional",
        };

        private static Dictionary<XName, PaSchemaSet> _paSchemaSets;

        public static WmlDocument AssembleDocument(WmlDocument templateDoc, XmlDocument data, out bool templateError)
        {
            XDocument xDoc = data.GetXDocument();
            return AssembleDocument(templateDoc, xDoc.Root, out templateError);
        }

        public static WmlDocument AssembleDocument(WmlDocument templateDoc, XElement data, out bool templateError)
        {
            using var mem = new MemoryStream();
            byte[] byteArray = templateDoc.DocumentByteArray;
            mem.Write(byteArray, 0, byteArray.Length);

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, true))
            {
                if (RevisionAccepter.HasTrackedRevisions(wordDoc))
                {
                    throw new OpenXmlPowerToolsException("Invalid DocumentAssembler template - contains tracked revisions");
                }

                var te = new TemplateError();

                foreach (OpenXmlPart part in wordDoc.ContentParts())
                {
                    ProcessTemplatePart(data, te, part);
                }

                templateError = te.HasError;
            }

            return new WmlDocument("TempFileName.docx", mem.ToArray());
        }

        private static void ProcessTemplatePart(XElement data, TemplateError te, OpenXmlPart part)
        {
            XDocument xDoc = part.GetXDocument();

            XElement xDocRoot = RemoveGoBackBookmarks(xDoc.Root);

            // content controls in cells can surround the W.tc element, so transform so that such content controls are within the cell content
            xDocRoot = (XElement) NormalizeContentControlsInCells(xDocRoot);

            xDocRoot = (XElement) TransformToMetadata(xDocRoot, te);

            // Table might have been placed at run-level, when it should be at block-level, so fix this.
            // Repeat, EndRepeat, Conditional, EndConditional are allowed at run level, but only if there is a matching pair
            // if there is only one Repeat, EndRepeat, Conditional, EndConditional, then move to block level.
            // if there is a matching pair, then is OK.
            xDocRoot = (XElement) ForceBlockLevelAsAppropriate(xDocRoot, te);

            NormalizeTablesRepeatAndConditional(xDocRoot, te);

            // any EndRepeat, EndConditional that remain are orphans, so replace with an error
            ProcessOrphanEndRepeatEndConditional(xDocRoot, te);

            // do the actual content replacement
            xDocRoot = (XElement) ContentReplacementTransform(xDocRoot, data, te);

            xDoc.Elements().First().ReplaceWith(xDocRoot);
            part.SaveXDocument();
        }

        private static object ForceBlockLevelAsAppropriate(XNode node, TemplateError te)
        {
            if (node is XElement element)
            {
                if (element.Name == W.p)
                {
                    List<XElement> childMeta = element.Elements().Where(n => MetaToForceToBlock.Contains(n.Name)).ToList();

                    if (childMeta.Count == 1)
                    {
                        XElement child = childMeta.First();

                        string otherTextInParagraph =
                            element.Elements(W.r).Elements(W.t).Select(t => (string) t).StringConcatenate().Trim();

                        if (otherTextInParagraph != "")
                        {
                            var newPara = new XElement(element);
                            XElement newMeta = newPara.Elements().First(n => MetaToForceToBlock.Contains(n.Name));

                            newMeta.ReplaceWith(
                                CreateRunErrorMessage("Error: Unmatched metadata can't be in paragraph with other text", te));

                            return newPara;
                        }

                        var meta = new XElement(child.Name,
                            child.Attributes(),
                            new XElement(W.p,
                                element.Attributes(),
                                element.Elements(W.pPr),
                                child.Elements()));

                        return meta;
                    }

                    int count = childMeta.Count;

                    if (count % 2 == 0)
                    {
                        if (childMeta.Count(c => c.Name == Pa.Repeat) !=
                            childMeta.Count(c => c.Name == Pa.EndRepeat))
                        {
                            return CreateContextErrorMessage(element, "Error: Mismatch Repeat / EndRepeat at run level", te);
                        }

                        if (childMeta.Count(c => c.Name == Pa.Conditional) !=
                            childMeta.Count(c => c.Name == Pa.EndConditional))
                        {
                            return CreateContextErrorMessage(element, "Error: Mismatch Conditional / EndConditional at run level",
                                te);
                        }

                        return new XElement(element.Name,
                            element.Attributes(),
                            element.Nodes().Select(n => ForceBlockLevelAsAppropriate(n, te)));
                    }

                    return CreateContextErrorMessage(element, "Error: Invalid metadata at run level", te);
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => ForceBlockLevelAsAppropriate(n, te)));
            }

            return node;
        }

        private static void ProcessOrphanEndRepeatEndConditional(XElement xDocRoot, TemplateError te)
        {
            foreach (XElement element in xDocRoot.Descendants(Pa.EndRepeat).ToList())
            {
                object error = CreateContextErrorMessage(element, "Error: EndRepeat without matching Repeat", te);
                element.ReplaceWith(error);
            }

            foreach (XElement element in xDocRoot.Descendants(Pa.EndConditional).ToList())
            {
                object error = CreateContextErrorMessage(element, "Error: EndConditional without matching Conditional", te);
                element.ReplaceWith(error);
            }
        }

        private static XElement RemoveGoBackBookmarks(XElement xElement)
        {
            var cloneXDoc = new XElement(xElement);

            while (true)
            {
                XElement bm = cloneXDoc
                    .DescendantsAndSelf(W.bookmarkStart)
                    .FirstOrDefault(b => (string) b.Attribute(W.name) == "_GoBack");

                if (bm == null)
                {
                    break;
                }

                var id = (string) bm.Attribute(W.id);

                XElement endBm = cloneXDoc
                    .DescendantsAndSelf(W.bookmarkEnd)
                    .FirstOrDefault(b => (string) b.Attribute(W.id) == id);

                bm.Remove();
                endBm?.Remove();
            }

            return cloneXDoc;
        }

        // this transform inverts content controls that surround W.tc elements.  After transforming, the W.tc will contain
        // the content control, which contains the paragraph content of the cell.
        private static object NormalizeContentControlsInCells(XNode node)
        {
            return node is XElement element
                ? element.Name == W.sdt && element.Parent?.Name == W.tr
                    ? new XElement(W.tc,
                        element.Elements(W.tc).Elements(W.tcPr),
                        new XElement(W.sdt,
                            element.Elements(W.sdtPr),
                            element.Elements(W.sdtEndPr),
                            new XElement(W.sdtContent,
                                element.Elements(W.sdtContent).Elements(W.tc).Elements().Where(e => e.Name != W.tcPr))))
                    : new XElement(element.Name,
                        element.Attributes(),
                        element.Nodes().Select(NormalizeContentControlsInCells))
                : node;
        }

        // The following method is written using tree modification, not RPFT, because it is easier to write in this fashion.
        // These types of operations are not as easy to write using RPFT.
        // Unless you are completely clear on the semantics of LINQ to XML DML, do not make modifications to this method.
        private static void NormalizeTablesRepeatAndConditional(XElement xDoc, TemplateError te)
        {
            List<XElement> tables = xDoc.Descendants(Pa.Table).ToList();

            foreach (XElement table in tables)
            {
                XElement followingElement = table
                    .ElementsAfterSelf()
                    .FirstOrDefault(e => e.Name == W.tbl || e.Name == W.p);

                if (followingElement == null || followingElement.Name != W.tbl)
                {
                    table.ReplaceWith(CreateParaErrorMessage("Table metadata is not immediately followed by a table", te));
                    continue;
                }

                // Remove superflous paragraph from Table metadata
                table.RemoveNodes();

                // Detach w:tbl from parent, and add to Table metadata
                followingElement.Remove();
                table.Add(followingElement);
            }

            var repeatDepth = 0;
            var conditionalDepth = 0;

            foreach (XElement metadata in xDoc.Descendants()
                         .Where(d =>
                             d.Name == Pa.Repeat ||
                             d.Name == Pa.Conditional ||
                             d.Name == Pa.EndRepeat ||
                             d.Name == Pa.EndConditional))
            {
                if (metadata.Name == Pa.Repeat)
                {
                    ++repeatDepth;
                    metadata.Add(new XAttribute(Pa.Depth, repeatDepth));
                    continue;
                }

                if (metadata.Name == Pa.EndRepeat)
                {
                    metadata.Add(new XAttribute(Pa.Depth, repeatDepth));
                    --repeatDepth;
                    continue;
                }

                if (metadata.Name == Pa.Conditional)
                {
                    ++conditionalDepth;
                    metadata.Add(new XAttribute(Pa.Depth, conditionalDepth));
                    continue;
                }

                if (metadata.Name == Pa.EndConditional)
                {
                    metadata.Add(new XAttribute(Pa.Depth, conditionalDepth));
                    --conditionalDepth;
                }
            }

            while (true)
            {
                var didReplace = false;

                foreach (XElement metadata in xDoc.Descendants()
                             .Where(d => (d.Name == Pa.Repeat || d.Name == Pa.Conditional) && d.Attribute(Pa.Depth) != null)
                             .ToList())
                {
                    var depth = (int) metadata.Attribute(Pa.Depth);
                    XName matchingEndName = null;

                    if (metadata.Name == Pa.Repeat)
                    {
                        matchingEndName = Pa.EndRepeat;
                    }
                    else if (metadata.Name == Pa.Conditional)
                    {
                        matchingEndName = Pa.EndConditional;
                    }

                    if (matchingEndName == null)
                    {
                        throw new OpenXmlPowerToolsException("Internal error");
                    }

                    XElement matchingEnd = metadata.ElementsAfterSelf(matchingEndName)
                        .FirstOrDefault(end => (int) end.Attribute(Pa.Depth) == depth);

                    if (matchingEnd == null)
                    {
                        metadata.ReplaceWith(CreateParaErrorMessage(
                            $"{metadata.Name.LocalName} does not have matching {matchingEndName.LocalName}",
                            te));

                        continue;
                    }

                    metadata.RemoveNodes();

                    List<XElement> contentBetween =
                        metadata.ElementsAfterSelf().TakeWhile(after => after != matchingEnd).ToList();

                    foreach (XElement item in contentBetween)
                    {
                        item.Remove();
                    }

                    contentBetween = contentBetween.Where(n => n.Name != W.bookmarkStart && n.Name != W.bookmarkEnd).ToList();
                    metadata.Add(contentBetween);
                    metadata.Attributes(Pa.Depth).Remove();
                    matchingEnd.Remove();
                    didReplace = true;
                    break;
                }

                if (!didReplace)
                {
                    break;
                }
            }
        }

        private static object TransformToMetadata(XNode node, TemplateError te)
        {
            if (node is XElement element)
            {
                if (element.Name == W.sdt)
                {
                    var alias = (string) element.Elements(W.sdtPr).Elements(W.alias).Attributes(W.val).FirstOrDefault();

                    if (string.IsNullOrEmpty(alias) || AliasList.Contains(alias))
                    {
                        string ccContents = element
                            .DescendantsTrimmed(W.txbxContent)
                            .Where(e => e.Name == W.t)
                            .Select(t => (string) t)
                            .StringConcatenate()
                            .Trim()
                            .Replace('“', '"')
                            .Replace('”', '"');

                        if (ccContents.StartsWith("<", StringComparison.Ordinal))
                        {
                            XElement xml = TransformXmlTextToMetadata(te, ccContents);

                            if (xml.Name == W.p || xml.Name == W.r) // this means there was an error processing the XML.
                            {
                                if (element.Parent?.Name == W.p)
                                {
                                    return xml.Elements(W.r);
                                }

                                return xml;
                            }

                            if (alias != null && xml.Name.LocalName != alias)
                            {
                                if (element.Parent?.Name == W.p)
                                {
                                    return CreateRunErrorMessage(
                                        "Error: Content control alias does not match metadata element name", te);
                                }

                                return CreateParaErrorMessage(
                                    "Error: Content control alias does not match metadata element name", te);
                            }

                            xml.Add(element.Elements(W.sdtContent).Elements());
                            return xml;
                        }

                        return new XElement(element.Name,
                            element.Attributes(),
                            element.Nodes().Select(n => TransformToMetadata(n, te)));
                    }

                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Nodes().Select(n => TransformToMetadata(n, te)));
                }

                if (element.Name == W.p)
                {
                    string paraContents = element
                        .DescendantsTrimmed(W.txbxContent)
                        .Where(e => e.Name == W.t)
                        .Select(t => (string) t)
                        .StringConcatenate()
                        .Trim();

                    int occurances = paraContents
                        .Select((_, i) => paraContents.Substring(i))
                        .Count(sub => sub.StartsWith("<#", StringComparison.Ordinal));

                    if (paraContents.StartsWith("<#", StringComparison.Ordinal) &&
                        paraContents.EndsWith("#>", StringComparison.Ordinal) &&
                        occurances == 1)
                    {
                        string xmlText = paraContents.Substring(2, paraContents.Length - 4).Trim();
                        XElement xml = TransformXmlTextToMetadata(te, xmlText);

                        if (xml.Name == W.p || xml.Name == W.r)
                        {
                            return xml;
                        }

                        xml.Add(element);
                        return xml;
                    }

                    if (paraContents.Contains("<#"))
                    {
                        var runReplacementInfo = new List<RunReplacementInfo>();
                        var thisGuid = Guid.NewGuid().ToString();
                        var r = new Regex("<#.*?#>");
                        XElement xml;

                        OpenXmlRegex.Replace(new[] { element }, r, thisGuid, (_, match) =>
                        {
                            string matchString = match.Value.Trim();

                            string xmlText = matchString.Substring(2, matchString.Length - 4)
                                .Trim()
                                .Replace('“', '"')
                                .Replace('”', '"');

                            try
                            {
                                xml = XElement.Parse(xmlText);
                            }
                            catch (XmlException e)
                            {
                                var rri = new RunReplacementInfo
                                {
                                    Xml = null,
                                    XmlExceptionMessage = "XmlException: " + e.Message,
                                    SchemaValidationMessage = null,
                                };

                                runReplacementInfo.Add(rri);
                                return true;
                            }

                            string schemaError = ValidatePerSchema(xml);

                            if (schemaError != null)
                            {
                                var rri = new RunReplacementInfo
                                {
                                    Xml = null,
                                    XmlExceptionMessage = null,
                                    SchemaValidationMessage = "Schema Validation Error: " + schemaError,
                                };

                                runReplacementInfo.Add(rri);
                                return true;
                            }

                            var rri2 = new RunReplacementInfo
                            {
                                Xml = xml,
                                XmlExceptionMessage = null,
                                SchemaValidationMessage = null,
                            };

                            runReplacementInfo.Add(rri2);
                            return true;
                        }, false);

                        var newPara = new XElement(element);

                        foreach (RunReplacementInfo rri in runReplacementInfo)
                        {
                            XElement runToReplace = newPara.Descendants(W.r)
                                .FirstOrDefault(rn => rn.Value == thisGuid && rn.Parent?.Name != Pa.Content);

                            if (runToReplace == null)
                            {
                                throw new OpenXmlPowerToolsException("Internal error");
                            }

                            if (rri.XmlExceptionMessage != null)
                            {
                                runToReplace.ReplaceWith(CreateRunErrorMessage(rri.XmlExceptionMessage, te));
                            }
                            else if (rri.SchemaValidationMessage != null)
                            {
                                runToReplace.ReplaceWith(CreateRunErrorMessage(rri.SchemaValidationMessage, te));
                            }
                            else
                            {
                                var newXml = new XElement(rri.Xml);
                                newXml.Add(runToReplace);
                                runToReplace.ReplaceWith(newXml);
                            }
                        }

                        XElement coalescedParagraph = WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(newPara);
                        return coalescedParagraph;
                    }
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => TransformToMetadata(n, te)));
            }

            return node;
        }

        private static XElement TransformXmlTextToMetadata(TemplateError te, string xmlText)
        {
            XElement xml;

            try
            {
                xml = XElement.Parse(xmlText);
            }
            catch (XmlException e)
            {
                return CreateParaErrorMessage("XmlException: " + e.Message, te);
            }

            string schemaError = ValidatePerSchema(xml);

            if (schemaError != null)
            {
                return CreateParaErrorMessage("Schema Validation Error: " + schemaError, te);
            }

            return xml;
        }

        private static string ValidatePerSchema(XElement element)
        {
            if (_paSchemaSets == null)
            {
                _paSchemaSets = new Dictionary<XName, PaSchemaSet>
                {
                    {
                        Pa.Content,
                        new PaSchemaSet
                        {
                            XsdMarkup =
                                @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Content'>
                                    <xs:complexType>
                                      <xs:attribute name='Select' type='xs:string' use='required' />
                                      <xs:attribute name='Optional' type='xs:boolean' use='optional' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:schema>",
                        }
                    },
                    {
                        Pa.Table,
                        new PaSchemaSet
                        {
                            XsdMarkup =
                                @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Table'>
                                    <xs:complexType>
                                      <xs:attribute name='Select' type='xs:string' use='required' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:schema>",
                        }
                    },
                    {
                        Pa.Repeat,
                        new PaSchemaSet
                        {
                            XsdMarkup =
                                @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Repeat'>
                                    <xs:complexType>
                                      <xs:attribute name='Select' type='xs:string' use='required' />
                                      <xs:attribute name='Optional' type='xs:boolean' use='optional' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:schema>",
                        }
                    },
                    {
                        Pa.EndRepeat,
                        new PaSchemaSet
                        {
                            XsdMarkup =
                                @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='EndRepeat' />
                                </xs:schema>",
                        }
                    },
                    {
                        Pa.Conditional,
                        new PaSchemaSet
                        {
                            XsdMarkup =
                                @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Conditional'>
                                    <xs:complexType>
                                      <xs:attribute name='Select' type='xs:string' use='required' />
                                      <xs:attribute name='Match' type='xs:string' use='optional' />
                                      <xs:attribute name='NotMatch' type='xs:string' use='optional' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:schema>",
                        }
                    },
                    {
                        Pa.EndConditional,
                        new PaSchemaSet
                        {
                            XsdMarkup =
                                @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='EndConditional' />
                                </xs:schema>",
                        }
                    },
                };

                foreach (KeyValuePair<XName, PaSchemaSet> item in _paSchemaSets)
                {
                    PaSchemaSet itemPAss = item.Value;
                    var schemas = new XmlSchemaSet();
                    schemas.Add("", XmlReader.Create(new StringReader(itemPAss.XsdMarkup)));
                    itemPAss.SchemaSet = schemas;
                }
            }

            if (!_paSchemaSets.ContainsKey(element.Name))
            {
                return $"Invalid XML: {element.Name.LocalName} is not a valid element";
            }

            PaSchemaSet paSchemaSet = _paSchemaSets[element.Name];
            var d = new XDocument(element);
            string message = null;

            d.Validate(paSchemaSet.SchemaSet, (_, e) => message ??= e.Message, true);

            return message;
        }

        private static object ContentReplacementTransform(XNode node, XElement data, TemplateError templateError)
        {
            if (node is XElement element)
            {
                if (element.Name == Pa.Content)
                {
                    XElement para = element.Descendants(W.p).FirstOrDefault();
                    XElement run = element.Descendants(W.r).FirstOrDefault();

                    var xPath = (string) element.Attribute(Pa.Select);
                    var optionalString = (string) element.Attribute(Pa.Optional);
                    bool optional = optionalString != null && optionalString.ToLower() == "true";

                    string newValue;

                    try
                    {
                        newValue = EvaluateXPathToString(data, xPath, optional);
                    }
                    catch (XPathException e)
                    {
                        return CreateContextErrorMessage(element, "XPathException: " + e.Message, templateError);
                    }

                    if (para != null)
                    {
                        var p = new XElement(W.p, para.Elements(W.pPr));

                        foreach (string line in newValue.Split('\n'))
                        {
                            p.Add(new XElement(W.r,
                                para.Elements(W.r).Elements(W.rPr).FirstOrDefault(),
                                p.Elements().Count() > 1 ? new XElement(W.br) : null,
                                new XElement(W.t, line)));
                        }

                        return p;
                    }

                    var list = new List<XElement>();

                    foreach (string line in newValue.Split('\n'))
                    {
                        list.Add(new XElement(W.r,
                            run!.Elements().Where(e => e.Name != W.t),
                            list.Count > 0 ? new XElement(W.br) : null,
                            new XElement(W.t, line)));
                    }

                    return list;
                }

                if (element.Name == Pa.Repeat)
                {
                    var selector = (string) element.Attribute(Pa.Select);
                    var optionalString = (string) element.Attribute(Pa.Optional);
                    bool optional = optionalString != null && optionalString.ToLower() == "true";

                    IEnumerable<XElement> repeatingData;

                    try
                    {
                        repeatingData = data.XPathSelectElements(selector).ToList();
                    }
                    catch (XPathException e)
                    {
                        return CreateContextErrorMessage(element, "XPathException: " + e.Message, templateError);
                    }

                    if (!repeatingData.Any())
                    {
                        return optional ? null : CreateContextErrorMessage(element, "Repeat: Select returned no data", templateError);
                    }

                    List<List<object>> newContent = repeatingData.Select(d =>
                        {
                            List<object> content = element
                                .Elements()
                                .Select(e => ContentReplacementTransform(e, d, templateError))
                                .ToList();

                            return content;
                        })
                        .ToList();

                    return newContent;
                }

                if (element.Name == Pa.Table)
                {
                    IEnumerable<XElement> tableData;

                    try
                    {
                        tableData = data.XPathSelectElements((string) element.Attribute(Pa.Select)).ToList();
                    }
                    catch (XPathException e)
                    {
                        return CreateContextErrorMessage(element, "XPathException: " + e.Message, templateError);
                    }

                    if (!tableData.Any())
                    {
                        return CreateContextErrorMessage(element, "Table Select returned no data", templateError);
                    }

                    XElement table = element.Element(W.tbl)!;
                    XElement protoRow = table.Elements(W.tr).Skip(1).FirstOrDefault();

                    List<XElement> footerRowsBeforeTransform = table
                        .Elements(W.tr)
                        .Skip(2)
                        .ToList();

                    List<object> footerRows = footerRowsBeforeTransform
                        .Select(x => ContentReplacementTransform(x, data, templateError))
                        .ToList();

                    if (protoRow == null)
                    {
                        return CreateContextErrorMessage(element, "Table does not contain a prototype row", templateError);
                    }

                    protoRow.Descendants(W.bookmarkStart).Remove();
                    protoRow.Descendants(W.bookmarkEnd).Remove();

                    var newTable = new XElement(W.tbl,
                        table.Elements().Where(e => e.Name != W.tr),
                        table.Elements(W.tr).FirstOrDefault(),
                        tableData.Select(d =>
                            new XElement(W.tr,
                                protoRow.Elements().Where(r => r.Name != W.tc),
                                protoRow.Elements(W.tc)
                                    .Select(tc =>
                                    {
                                        XElement paragraph = tc.Elements(W.p).First();
                                        XElement cellRun = paragraph.Elements(W.r).FirstOrDefault();
                                        string xPath = paragraph.Value;
                                        string newValue;

                                        try
                                        {
                                            newValue = EvaluateXPathToString(d, xPath, false);
                                        }
                                        catch (XPathException e)
                                        {
                                            var errorCell = new XElement(W.tc,
                                                tc.Elements().Where(z => z.Name != W.p),
                                                new XElement(W.p,
                                                    paragraph.Element(W.pPr),
                                                    CreateRunErrorMessage(e.Message, templateError)));

                                            return errorCell;
                                        }

                                        var newCell = new XElement(W.tc,
                                            tc.Elements().Where(z => z.Name != W.p),
                                            new XElement(W.p,
                                                paragraph.Element(W.pPr),
                                                new XElement(W.r,
                                                    cellRun != null
                                                        ? cellRun.Element(W.rPr)
                                                        : new XElement(W.rPr), //if the cell was empty there is no cellrun
                                                    new XElement(W.t, newValue))));

                                        return newCell;
                                    }))),
                        footerRows);

                    return newTable;
                }

                if (element.Name == Pa.Conditional)
                {
                    var xPath = (string) element.Attribute(Pa.Select);
                    var match = (string) element.Attribute(Pa.Match);
                    var notMatch = (string) element.Attribute(Pa.NotMatch);

                    if (match == null && notMatch == null)
                    {
                        return CreateContextErrorMessage(element, "Conditional: Must specify either Match or NotMatch",
                            templateError);
                    }

                    if (match != null && notMatch != null)
                    {
                        return CreateContextErrorMessage(element, "Conditional: Cannot specify both Match and NotMatch",
                            templateError);
                    }

                    string testValue;

                    try
                    {
                        testValue = EvaluateXPathToString(data, xPath, false);
                    }
                    catch (XPathException e)
                    {
                        return CreateContextErrorMessage(element, e.Message, templateError);
                    }

                    if ((match != null && testValue == match) || (notMatch != null && testValue != notMatch))
                    {
                        IEnumerable<object> content = element
                            .Elements()
                            .Select(e => ContentReplacementTransform(e, data, templateError));

                        return content;
                    }

                    return null;
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => ContentReplacementTransform(n, data, templateError)));
            }

            return node;
        }

        private static object CreateContextErrorMessage(XElement element, string errorMessage, TemplateError templateError)
        {
            XElement para = element.Descendants(W.p).FirstOrDefault();
            XElement errorRun = CreateRunErrorMessage(errorMessage, templateError);

            return para != null ? new XElement(W.p, errorRun) : errorRun;
        }

        private static XElement CreateRunErrorMessage(string errorMessage, TemplateError templateError)
        {
            templateError.HasError = true;

            var errorRun = new XElement(W.r,
                new XElement(W.rPr,
                    new XElement(W.color, new XAttribute(W.val, "FF0000")),
                    new XElement(W.highlight, new XAttribute(W.val, "yellow"))),
                new XElement(W.t, errorMessage));

            return errorRun;
        }

        private static XElement CreateParaErrorMessage(string errorMessage, TemplateError templateError)
        {
            templateError.HasError = true;

            var errorPara = new XElement(W.p,
                new XElement(W.r,
                    new XElement(W.rPr,
                        new XElement(W.color, new XAttribute(W.val, "FF0000")),
                        new XElement(W.highlight, new XAttribute(W.val, "yellow"))),
                    new XElement(W.t, errorMessage)));

            return errorPara;
        }

        private static string EvaluateXPathToString(XElement element, string xPath, bool optional)
        {
            object xPathSelectResult;

            try
            {
                // Support some cells in the table may not have an xpath expression.
                if (string.IsNullOrWhiteSpace(xPath))
                {
                    return string.Empty;
                }

                xPathSelectResult = element.XPathEvaluate(xPath);
            }
            catch (XPathException e)
            {
                throw new XPathException("XPathException: " + e.Message, e);
            }

            // TODO: Revisit. Does that make sense?
            if (xPathSelectResult is IEnumerable enumerable and not string)
            {
                List<XObject> selectedData = enumerable.Cast<XObject>().ToList();

                if (!selectedData.Any())
                {
                    return optional ? string.Empty : throw new XPathException($"XPath expression ({xPath}) returned no results");
                }

                if (selectedData.Count > 1)
                {
                    throw new XPathException($"XPath expression ({xPath}) returned more than one node");
                }

                XObject selectedDatum = selectedData.First();

                if (selectedDatum is XElement xElement)
                {
                    return xElement.Value;
                }

                if (selectedDatum is XAttribute xAttribute)
                {
                    return xAttribute.Value;
                }
            }

            return xPathSelectResult.ToString();
        }

        private sealed class RunReplacementInfo
        {
            public XElement Xml;
            public string XmlExceptionMessage;
            public string SchemaValidationMessage;
        }

        private static class Pa
        {
            public static readonly XName Content = "Content";
            public static readonly XName Table = "Table";
            public static readonly XName Repeat = "Repeat";
            public static readonly XName EndRepeat = "EndRepeat";
            public static readonly XName Conditional = "Conditional";
            public static readonly XName EndConditional = "EndConditional";

            public static readonly XName Select = "Select";
            public static readonly XName Optional = "Optional";
            public static readonly XName Match = "Match";
            public static readonly XName NotMatch = "NotMatch";
            public static readonly XName Depth = "Depth";
        }

        private sealed class PaSchemaSet
        {
            public string XsdMarkup;
            public XmlSchemaSet SchemaSet;
        }

        private sealed class TemplateError
        {
            public bool HasError;
        }
    }
}
