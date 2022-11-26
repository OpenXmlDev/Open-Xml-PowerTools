// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Xml.Schema;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System.Collections;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Path = System.IO.Path;

namespace OpenXmlPowerTools
{
    public class DocumentAssembler
    {
        public static WmlDocument AssembleDocument(WmlDocument templateDoc, XmlDocument data, out bool templateError)
        {
            XDocument xDoc = data.GetXDocument();
            return AssembleDocument(templateDoc, xDoc.Root, out templateError);
        }

        public static WmlDocument AssembleDocument(WmlDocument templateDoc, XElement data, out bool templateError)
        {
            byte[] byteArray = templateDoc.DocumentByteArray;
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, (int)byteArray.Length);
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, true))
                {
                    if (RevisionAccepter.HasTrackedRevisions(wordDoc))
                        throw new OpenXmlPowerToolsException("Invalid DocumentAssembler template - contains tracked revisions");

                    // calculate and store the max docPr id for later use when adding image objects
                    var macDocPrId = GetMaxDocPrId(wordDoc);

                    var te = new TemplateError();
                    foreach (var part in wordDoc.ContentParts())
                    {
                        ProcessTemplatePart(data, te, part);
                    }
                    templateError = te.HasError;

                    // update image docPr ids for the whole document
                    FixUpDocPrIds(wordDoc, macDocPrId);
                }
                WmlDocument assembledDocument = new WmlDocument("TempFileName.docx", mem.ToArray());
                return assembledDocument;
            }
        }

        private static void ProcessTemplatePart(XElement data, TemplateError te, OpenXmlPart part)
        {
            XDocument xDoc = part.GetXDocument();

            var xDocRoot = RemoveGoBackBookmarks(xDoc.Root);

            // process diagrams part
            // TODO: consider splitting this method into two for clarity, so, for example, there will be
            // TODO: pipeline-like processing: first the giagram, then the document part, or vice versa
            var diagramPart = part.GetPartsOfType<DiagramDataPart>().FirstOrDefault();
            if (diagramPart != null)
            {
                var diagramDoc = diagramPart.GetXDocument();
                if (diagramDoc != null)
                {
                    var dataPartRoot = diagramDoc.Root;
                    if (dataPartRoot != null)
                    {
                        dataPartRoot = (XElement)TransformToMetadata(dataPartRoot, data, te);
                        // do the actual content replacement
                        dataPartRoot = (XElement)ContentReplacementTransform(dataPartRoot, data, te, part);
                        diagramDoc.Elements().First().ReplaceWith(dataPartRoot);
                        diagramPart.PutXDocument();
                    }                
                }                
            }

            // content controls in cells can surround the W.tc element, so transform so that such content controls are within the cell content
            xDocRoot = (XElement)NormalizeContentControlsInCells(xDocRoot);

            xDocRoot = (XElement)TransformToMetadata(xDocRoot, data, te);

            // Table might have been placed at run-level, when it should be at block-level, so fix this.
            // Repeat, EndRepeat, Conditional, EndConditional are allowed at run level, but only if there is a matching pair
            // if there is only one Repeat, EndRepeat, Conditional, EndConditional, then move to block level.
            // if there is a matching pair, then is OK.
            xDocRoot = (XElement)ForceBlockLevelAsAppropriate(xDocRoot, te);

            NormalizeTablesRepeatAndConditional(xDocRoot, te);

            // any EndRepeat, EndConditional that remain are orphans, so replace with an error
            ProcessOrphanEndRepeatEndConditional(xDocRoot, te);

            // do the actual content replacement
            xDocRoot = (XElement)ContentReplacementTransform(xDocRoot, data, te, part);

            xDoc.Elements().First().ReplaceWith(xDocRoot);
            part.PutXDocument();
            return;
        }

        private static XName[] s_MetaToForceToBlock = new XName[] {
            PA.Conditional,
            PA.EndConditional,
            PA.Repeat,
            PA.EndRepeat,
            PA.Table,
            PA.Image
        };

        private static object ForceBlockLevelAsAppropriate(XNode node, TemplateError te)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.p)
                {
                    var childMeta = element.Elements().Where(n => s_MetaToForceToBlock.Contains(n.Name)).ToList();
                    if (childMeta.Count() == 1)
                    {
                        var child = childMeta.First();
                        var otherTextInParagraph = element.Elements(W.r).Elements(W.t).Select(t => (string)t).StringConcatenate().Trim();
                        if (otherTextInParagraph != "")
                        {
                            var newPara = new XElement(element);
                            var newMeta = newPara.Elements().Where(n => s_MetaToForceToBlock.Contains(n.Name)).First();
                            newMeta.ReplaceWith(CreateRunErrorMessage("Error: Unmatched metadata can't be in paragraph with other text", te));
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
                    var count = childMeta.Count();
                    if (count % 2 == 0)
                    {
                        if (childMeta.Where(c => c.Name == PA.Repeat).Count() != childMeta.Where(c => c.Name == PA.EndRepeat).Count())
                            return CreateContextErrorMessage(element, "Error: Mismatch Repeat / EndRepeat at run level", te);
                        if (childMeta.Where(c => c.Name == PA.Conditional).Count() != childMeta.Where(c => c.Name == PA.EndConditional).Count())
                            return CreateContextErrorMessage(element, "Error: Mismatch Conditional / EndConditional at run level", te);
                        return new XElement(element.Name,
                            element.Attributes(),
                            element.Nodes().Select(n => ForceBlockLevelAsAppropriate(n, te)));
                    }
                    else
                    {
                        return CreateContextErrorMessage(element, "Error: Invalid metadata at run level", te);
                    }
                }
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => ForceBlockLevelAsAppropriate(n, te)));
            }
            return node;
        }

        private static void ProcessOrphanEndRepeatEndConditional(XElement xDocRoot, TemplateError te)
        {
            foreach (var element in xDocRoot.Descendants(PA.EndRepeat).ToList())
            {
                var error = CreateContextErrorMessage(element, "Error: EndRepeat without matching Repeat", te);
                element.ReplaceWith(error);
            }
            foreach (var element in xDocRoot.Descendants(PA.EndConditional).ToList())
            {
                var error = CreateContextErrorMessage(element, "Error: EndConditional without matching Conditional", te);
                element.ReplaceWith(error);
            }
        }

        private static XElement RemoveGoBackBookmarks(XElement xElement)
        {
            var cloneXDoc = new XElement(xElement);
            while (true)
            {
                var bm = cloneXDoc.DescendantsAndSelf(W.bookmarkStart).FirstOrDefault(b => (string)b.Attribute(W.name) == "_GoBack");
                if (bm == null)
                    break;
                var id = (string)bm.Attribute(W.id);
                var endBm = cloneXDoc.DescendantsAndSelf(W.bookmarkEnd).FirstOrDefault(b => (string)b.Attribute(W.id) == id);
                bm.Remove();
                endBm.Remove();
            }
            return cloneXDoc;
        }

        // this transform inverts content controls that surround W.tc elements.  After transforming, the W.tc will contain
        // the content control, which contains the paragraph content of the cell.
        private static object NormalizeContentControlsInCells(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.sdt && element.Parent.Name == W.tr)
                {
                    var newCell = new XElement(W.tc,
                        element.Elements(W.tc).Elements(W.tcPr),
                        new XElement(W.sdt,
                            element.Elements(W.sdtPr),
                            element.Elements(W.sdtEndPr),
                            new XElement(W.sdtContent,
                                element.Elements(W.sdtContent).Elements(W.tc).Elements().Where(e => e.Name != W.tcPr))));
                    return newCell;
                }
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => NormalizeContentControlsInCells(n)));
            }
            return node;
        }

        // The following method is written using tree modification, not RPFT, because it is easier to write in this fashion.
        // These types of operations are not as easy to write using RPFT.
        // Unless you are completely clear on the semantics of LINQ to XML DML, do not make modifications to this method.
        private static void NormalizeTablesRepeatAndConditional(XElement xDoc, TemplateError te)
        {
            var tables = xDoc.Descendants(PA.Table).ToList();
            foreach (var table in tables)
            {
                var followingElement = table.ElementsAfterSelf().Where(e => e.Name == W.tbl || e.Name == W.p).FirstOrDefault();
                if (followingElement == null || followingElement.Name != W.tbl)
                {
                    table.ReplaceWith(CreateParaErrorMessage("Table metadata is not immediately followed by a table", te));
                    continue;
                }
                // remove superflous paragraph from Table metadata
                table.RemoveNodes();
                // detach w:tbl from parent, and add to Table metadata
                followingElement.Remove();
                table.Add(followingElement);
            }

            var images = xDoc.Descendants(PA.Image).ToList();
            foreach (var image in images)
            {
                var followingElement = image.ElementsAfterSelf().FirstOrDefault(e => e.Name == W.sdt || e.Name == W.p);

                if (followingElement == null)
                {
                    image.ReplaceWith(CreateParaErrorMessage("Image metadata is not immediately followed by an image", te));
                    continue;
                }

                // get sdt element (can also be within a paragraph) and check it's contents
                var sdt = followingElement.Name == W.p ? followingElement.Elements().FirstOrDefault(e => e.Name == W.sdt) : followingElement;

                if (sdt != null && sdt.Name == W.sdt)
                {
                    // get sdt properties
                    var sdtPr = sdt.Elements().FirstOrDefault(e => e.Name == W.sdtPr);
                    if (sdtPr != null)
                    {
                        // check for properties if contain picture
                        var picture = sdtPr.Elements().FirstOrDefault(e => e.Name == W.picture);
                        if (picture == null)
                        {
                            image.ReplaceWith(
                                CreateParaErrorMessage("Image metadata does not contain picture element", te));
                            continue;
                        }
                    }                    
                }
                else
                {
                    // there might be the image without surrounding content control
                    image.RemoveNodes();
                    followingElement.Remove();
                    image.Add(followingElement);
                    continue;
                }

                // remove superflous paragraph from Image metadata
                image.RemoveNodes();
                // detach w:sdt from parent, and add to Image metadata
                followingElement.Remove();
                image.Add(followingElement);
            }

            int repeatDepth = 0;
            int conditionalDepth = 0;
            foreach (var metadata in xDoc.Descendants().Where(d =>
                    d.Name == PA.Repeat ||
                    d.Name == PA.Conditional ||
                    d.Name == PA.EndRepeat ||
                    d.Name == PA.EndConditional))
            {
                if (metadata.Name == PA.Repeat)
                {
                    ++repeatDepth;
                    metadata.Add(new XAttribute(PA.Depth, repeatDepth));
                    continue;
                }
                if (metadata.Name == PA.EndRepeat)
                {
                    metadata.Add(new XAttribute(PA.Depth, repeatDepth));
                    --repeatDepth;
                    continue;
                }
                if (metadata.Name == PA.Conditional)
                {
                    ++conditionalDepth;
                    metadata.Add(new XAttribute(PA.Depth, conditionalDepth));
                    continue;
                }
                if (metadata.Name == PA.EndConditional)
                {
                    metadata.Add(new XAttribute(PA.Depth, conditionalDepth));
                    --conditionalDepth;
                    continue;
                }
            }

            while (true)
            {
                bool didReplace = false;
                foreach (var metadata in xDoc.Descendants().Where(d => (d.Name == PA.Repeat || d.Name == PA.Conditional || d.Name == PA.Image) && d.Attribute(PA.Depth) != null).ToList())
                {
                    var depth = (int)metadata.Attribute(PA.Depth);
                    XName matchingEndName = null;
                    if (metadata.Name == PA.Repeat)
                        matchingEndName = PA.EndRepeat;
                    else if (metadata.Name == PA.Conditional)
                        matchingEndName = PA.EndConditional;
                    if (matchingEndName == null)
                        throw new OpenXmlPowerToolsException("Internal error");
                    var matchingEnd = metadata.ElementsAfterSelf(matchingEndName).FirstOrDefault(end => { return (int)end.Attribute(PA.Depth) == depth; });
                    if (matchingEnd == null)
                    {
                        metadata.ReplaceWith(CreateParaErrorMessage(string.Format("{0} does not have matching {1}", metadata.Name.LocalName, matchingEndName.LocalName), te));
                        continue;
                    }
                    metadata.RemoveNodes();
                    var contentBetween = metadata.ElementsAfterSelf().TakeWhile(after => after != matchingEnd).ToList();
                    foreach (var item in contentBetween)
                        item.Remove();
                    contentBetween = contentBetween.Where(n => n.Name != W.bookmarkStart && n.Name != W.bookmarkEnd).ToList();
                    metadata.Add(contentBetween);
                    metadata.Attributes(PA.Depth).Remove();
                    matchingEnd.Remove();
                    didReplace = true;
                    break;
                }
                if (!didReplace)
                    break;
            }
        }

        private static List<string> s_AliasList = new List<string>()
        {
            "Image",
            "Content",
            "Table",
            "Repeat",
            "EndRepeat",
            "Conditional",
            "EndConditional",
        };

        private static object TransformToMetadata(XNode node, XElement data, TemplateError te)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.sdt)
                {
                    var alias = (string)element.Elements(W.sdtPr).Elements(W.alias).Attributes(W.val).FirstOrDefault();
                    if (alias == null || alias == "" || s_AliasList.Contains(alias))
                    {
                        var ccContents = element
                            .DescendantsTrimmed(W.txbxContent)
                            .Where(e => e.Name == W.t)
                            .Select(t => (string)t)
                            .StringConcatenate()
                            .Trim()
                            .Replace('“', '"')
                            .Replace('”', '"');
                        if (ccContents.StartsWith("<"))
                        {
                            XElement xml = TransformXmlTextToMetadata(te, ccContents);
                            if (xml.Name == W.p || xml.Name == W.r)  // this means there was an error processing the XML.
                            {
                                if (element.Parent.Name == W.p)
                                    return xml.Elements(W.r);
                                return xml;
                            }
                            if (alias != null && xml.Name.LocalName != alias)
                            {
                                if (element.Parent.Name == W.p)
                                    return CreateRunErrorMessage("Error: Content control alias does not match metadata element name", te);
                                else
                                    return CreateParaErrorMessage("Error: Content control alias does not match metadata element name", te);
                            }
                            xml.Add(element.Elements(W.sdtContent).Elements());
                            return xml;
                        }
                        return new XElement(element.Name,
                            element.Attributes(),
                            element.Nodes().Select(n => TransformToMetadata(n, data, te)));
                    }
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Nodes().Select(n => TransformToMetadata(n, data, te)));
                }
                if (element.Name == A.r)
                {
                    var paraContents = element
                        .DescendantsTrimmed(W.txbxContent)
                        .Where(e => e.Name == A.t)
                        .Select(t => (string)t)
                        .StringConcatenate()
                        .Trim();
                    int occurances = paraContents.Select((c, i) => paraContents.Substring(i)).Count(sub => sub.StartsWith("<#"));
                    if (paraContents.StartsWith("<#") && paraContents.EndsWith("#>") && occurances == 1)
                    {
                        var xmlText = paraContents.Substring(2, paraContents.Length - 4).Trim();
                        XElement xml = TransformXmlTextToMetadata(te, xmlText);
                        if (xml.Name == W.p || xml.Name == W.r)
                            return xml;
                        xml.Add(element);
                        return xml;
                    }
                    if (paraContents.Contains("<#"))
                    {
                        List<RunReplacementInfo> runReplacementInfo = new List<RunReplacementInfo>();
                        var thisGuid = Guid.NewGuid().ToString();
                        var r = new Regex("<#.*?#>");
                        XElement xml = null;
                        OpenXmlRegex.Replace(new[] { element }, r, thisGuid, (para, match) =>
                        {
                            var matchString = match.Value.Trim();
                            var xmlText = matchString.Substring(2, matchString.Length - 4).Trim().Replace('“', '"').Replace('”', '"');
                            try
                            {
                                xml = XElement.Parse(xmlText);
                            }
                            catch (XmlException e)
                            {
                                RunReplacementInfo rri = new RunReplacementInfo()
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
                                RunReplacementInfo rri = new RunReplacementInfo()
                                {
                                    Xml = null,
                                    XmlExceptionMessage = null,
                                    SchemaValidationMessage = "Schema Validation Error: " + schemaError,
                                };
                                runReplacementInfo.Add(rri);
                                return true;
                            }
                            RunReplacementInfo rri2 = new RunReplacementInfo()
                            {
                                Xml = xml,
                                XmlExceptionMessage = null,
                                SchemaValidationMessage = null,
                            };
                            runReplacementInfo.Add(rri2);
                            return true;
                        }, false);

                        var newPara = new XElement(element);
                        foreach (var rri in runReplacementInfo)
                        {
                            var runToReplace = newPara.Descendants(W.r).FirstOrDefault(rn => rn.Value == thisGuid && rn.Parent.Name != PA.Content);
                            if (runToReplace == null)
                                throw new OpenXmlPowerToolsException("Internal error");
                            if (rri.XmlExceptionMessage != null)
                                runToReplace.ReplaceWith(CreateRunErrorMessage(rri.XmlExceptionMessage, te));
                            else if (rri.SchemaValidationMessage != null)
                                runToReplace.ReplaceWith(CreateRunErrorMessage(rri.SchemaValidationMessage, te));
                            else
                            {
                                var newXml = new XElement(rri.Xml);
                                newXml.Add(runToReplace);
                                runToReplace.ReplaceWith(newXml);
                            }
                        }
                        var coalescedParagraph = WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(newPara);
                        return coalescedParagraph;
                    }
                }
                if (element.Name == W.p)
                {
                    var paraContents = element
                        .DescendantsTrimmed(W.txbxContent)
                        .Where(e => e.Name == W.t)
                        .Select(t => (string)t)
                        .StringConcatenate()
                        .Trim();
                    int occurances = paraContents.Select((c, i) => paraContents.Substring(i)).Count(sub => sub.StartsWith("<#"));
                    if (paraContents.StartsWith("<#") && paraContents.EndsWith("#>") && occurances == 1)
                    {
                        var xmlText = paraContents.Substring(2, paraContents.Length - 4).Trim();
                        XElement xml = TransformXmlTextToMetadata(te, xmlText);
                        if (xml.Name == W.p || xml.Name == W.r)
                            return xml;
                        xml.Add(element);
                        return xml;
                    }
                    if (paraContents.Contains("<#"))
                    {
                        List<RunReplacementInfo> runReplacementInfo = new List<RunReplacementInfo>();
                        var thisGuid = Guid.NewGuid().ToString();
                        var r = new Regex("<#.*?#>");
                        XElement xml = null;
                        OpenXmlRegex.Replace(new[] { element }, r, thisGuid, (para, match) =>
                        {
                            var matchString = match.Value.Trim();
                            var xmlText = matchString.Substring(2, matchString.Length - 4).Trim().Replace('“', '"').Replace('”', '"');
                            try
                            {
                                xml = XElement.Parse(xmlText);
                            }
                            catch (XmlException e)
                            {
                                RunReplacementInfo rri = new RunReplacementInfo()
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
                                RunReplacementInfo rri = new RunReplacementInfo()
                                {
                                    Xml = null,
                                    XmlExceptionMessage = null,
                                    SchemaValidationMessage = "Schema Validation Error: " + schemaError,
                                };
                                runReplacementInfo.Add(rri);
                                return true;
                            }
                            RunReplacementInfo rri2 = new RunReplacementInfo()
                            {
                                Xml = xml,
                                XmlExceptionMessage = null,
                                SchemaValidationMessage = null,
                            };
                            runReplacementInfo.Add(rri2);
                            return true;
                        }, false);

                        var newPara = new XElement(element);
                        foreach (var rri in runReplacementInfo)
                        {
                            var runToReplace = newPara.Descendants(W.r).FirstOrDefault(rn => rn.Value == thisGuid && rn.Parent.Name != PA.Content);
                            if (runToReplace == null)
                                throw new OpenXmlPowerToolsException("Internal error");
                            if (rri.XmlExceptionMessage != null)
                                runToReplace.ReplaceWith(CreateRunErrorMessage(rri.XmlExceptionMessage, te));
                            else if (rri.SchemaValidationMessage != null)
                                runToReplace.ReplaceWith(CreateRunErrorMessage(rri.SchemaValidationMessage, te));
                            else
                            {
                                var newXml = new XElement(rri.Xml);
                                newXml.Add(runToReplace);
                                runToReplace.ReplaceWith(newXml);
                            }
                        }
                        var coalescedParagraph = WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(newPara);
                        return coalescedParagraph;
                    }
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => TransformToMetadata(n, data, te)));
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
                return CreateParaErrorMessage("Schema Validation Error: " + schemaError, te);
            return xml;
        }

        private class RunReplacementInfo
        {
            public XElement Xml;
            public string XmlExceptionMessage;
            public string SchemaValidationMessage;
        }

        private static string ValidatePerSchema(XElement element)
        {
            if (s_PASchemaSets == null)
            {
                s_PASchemaSets = new Dictionary<XName, PASchemaSet>()
                {
                    {
                        PA.Content,
                        new PASchemaSet() {
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
                        PA.Table,
                        new PASchemaSet() {
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
                        PA.Repeat,
                        new PASchemaSet() {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Repeat'>
                                    <xs:complexType>
                                      <xs:attribute name='Select' type='xs:string' use='required' />
                                      <xs:attribute name='Optional' type='xs:boolean' use='optional' />
                                      <xs:attribute name='Align' type='xs:string' use='optional' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:schema>",
                        }
                    },
                    {
                        PA.EndRepeat,
                        new PASchemaSet() {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='EndRepeat' />
                                </xs:schema>",
                        }
                    },
                    {
                        PA.Conditional,
                        new PASchemaSet() {
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
                        PA.EndConditional,
                        new PASchemaSet() {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='EndConditional' />
                                </xs:schema>",
                        }
                    },
                    {
                        PA.Image,
                        new PASchemaSet() {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Image'>
                                    <xs:complexType>
                                      <xs:attribute name='Select' type='xs:string' use='required' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:schema>",
                        }
                    }
                };
                foreach (var item in s_PASchemaSets)
                {
                    var itemPAss = item.Value;
                    XmlSchemaSet schemas = new XmlSchemaSet();
                    schemas.Add("", XmlReader.Create(new StringReader(itemPAss.XsdMarkup)));
                    itemPAss.SchemaSet = schemas;
                }
            }
            if (!s_PASchemaSets.ContainsKey(element.Name))
            {
                return string.Format("Invalid XML: {0} is not a valid element", element.Name.LocalName);
            }
            var paSchemaSet = s_PASchemaSets[element.Name];
            XDocument d = new XDocument(element);
            string message = null;
            d.Validate(paSchemaSet.SchemaSet, (sender, e) =>
            {
                if (message == null)
                    message = e.Message;
            }, true);
            if (message != null)
                return message;
            return null;
        }

        private class PA
        {
            public static XName Image = "Image";
            public static XName Content = "Content";
            public static XName Table = "Table";
            public static XName Repeat = "Repeat";
            public static XName EndRepeat = "EndRepeat";
            public static XName Conditional = "Conditional";
            public static XName EndConditional = "EndConditional";

            public static XName Select = "Select";
            public static XName Optional = "Optional";
            public static XName Match = "Match";
            public static XName NotMatch = "NotMatch";
            public static XName Depth = "Depth";
            public static XName Align = "Align";
        }

        private class PASchemaSet
        {
            public string XsdMarkup;
            public XmlSchemaSet SchemaSet;
        }

        private static Dictionary<XName, PASchemaSet> s_PASchemaSets = null;

        private class TemplateError
        {
            public bool HasError = false;
        }

        /// <summary>
        /// Gets the next image relationship identifier of given part. The
        /// parts can be either header, footer or main document part. The method
        /// scans for already present relationship identifiers, then increments and
        /// returns the next available value.
        /// </summary>
        /// <param name="part">The part.</param>
        /// <returns>System.String.</returns>
        private static string GetNextImageRelationshipId(OpenXmlPart part)
        {
            var mainDocumentPart = part as MainDocumentPart;
            if (mainDocumentPart != null)
            {
                var imageId = mainDocumentPart.Parts
                    .Select(p => Regex.Match(p.RelationshipId, @"rId(?<rId>\d+)").Groups["rId"].Value)
                    .Max(x => Convert.ToDecimal(x));

                return string.Format("rId{0}", ++imageId);
            }

            var headerPart = part as HeaderPart;
            if (headerPart != null)
            {
                var imageId = headerPart.Parts
                    .Select(p => Regex.Match(p.RelationshipId, @"rId(?<rId>\d+)").Groups["rId"].Value)
                    .Max(x => Convert.ToDecimal(x));

                return string.Format("rId{0}", ++imageId);
            }

            var footerPart = part as FooterPart;
            if (footerPart != null)
            {
                var imageId = footerPart.Parts
                    .Select(p => Regex.Match(p.RelationshipId, @"rId(?<rId>\d+)").Groups["rId"].Value)
                    .Max(x => Convert.ToDecimal(x));

                return string.Format("rId{0}", ++imageId);
            }

            return null;
        }

        /// <summary>
        /// Calculates the maximum docPr id. The identifier is
        /// unique throughout the document. This method
        /// scans the whole document, finds and stores the max number (id is signed
        /// 23 bit integer).
        /// </summary>
        /// <param name="wordDoc">The word document.</param>
        /// <returns>System.Decimal.</returns>
        private static decimal GetMaxDocPrId(WordprocessingDocument wordDoc)
        {
            var idsList = new List<string>();
            foreach (var part in wordDoc.ContentParts())
            {
                idsList.AddRange(part.GetXDocument().Descendants(WP.docPr)
                    .SelectMany(e => e.Attributes().Where(a => a.Name == NoNamespace.id)).Select(v => v.Value));
            }
            return idsList.Count == 0 ? 0 : idsList.Max(x => Convert.ToDecimal(x));
        }

        private const string InvalidImageId = "InvalidImageId";

        /// <summary>
        /// Fixes docPrIds for the document. The identifier is unique throughout the
        /// document. This method scans the whole document, finds and replaces the
        /// image ids which were marked as invalid with incremental id 
        /// (id is signed 23 bit integer).
        /// </summary>
        /// <param name="wDoc">The word processing document.</param>
        /// <param name="maxDocPrId">The current maximum document pr identifier calculated 
        /// before the document has been processed.</param>
        private static void FixUpDocPrIds(WordprocessingDocument wDoc, decimal maxDocPrId)
        {
            var elementToFind = WP.docPr;
            var docPrToChange = wDoc
                .ContentParts()
                .Select(cp => cp.GetXDocument())
                .Select(xd => xd.Descendants().Where(d => d.Name == elementToFind))
                .SelectMany(m => m);
            var nextId = maxDocPrId;
            foreach (var item in docPrToChange)
            {
                var idAtt = item.Attribute(NoNamespace.id);
                if (idAtt != null && idAtt.Value == InvalidImageId)
                    idAtt.Value = string.Format("{0}", ++nextId);
            }
            foreach (var cp in wDoc.ContentParts())
                cp.PutXDocument();
        }

        // shape type identifier
        private static int _shapeTypeId = 1;
        private static int GetNextShapeType()
        {
            return _shapeTypeId++;
        }

        // shape identifier
        private static int _shapeId = 2000;
        private static string GetNextShapeId()
        {
            return string.Format("_x0000_s{0}", _shapeId++);
        }

        /// <summary>
        /// Creates and returns the image part inside the given part. The
        /// part can be either header, footer or main document part.
        /// </summary>
        /// <param name="part">The part.</param>
        /// <param name="imagePartType">Type of the image part.</param>
        /// <param name="relationshipId">The relationship identifier.</param>
        /// <returns>ImagePart.</returns>
        private static ImagePart GetImagePart(OpenXmlPart part, ImagePartType imagePartType, string relationshipId)
        {
            var mainDocumentPart = part as MainDocumentPart;
            if (mainDocumentPart != null)
            {
                return mainDocumentPart.AddImagePart(imagePartType, relationshipId);
            }

            var headerPart = part as HeaderPart;
            if (headerPart != null)
            {
                return headerPart.AddImagePart(imagePartType, relationshipId);
            }

            var footerPart = part as FooterPart;
            if (footerPart != null)
            {
                return footerPart.AddImagePart(imagePartType, relationshipId);
            }

            return null;
        }

        /// <summary>
        /// Method processes the image content and generates image element
        /// </summary>
        /// <param name="element">Source element</param>
        /// <param name="data">Data element with content</param>
        /// <param name="templateError">Error indicator</param>
        /// <param name="part">The part where the image is getting processed.</param>
        /// <returns>Image element</returns>
        private static object ProcessImageContent(XElement element, XElement data, TemplateError templateError, OpenXmlPart part)
        {
            // check for misplaced sdt content, should contain the paragraph and not vice versa
            var sdt = element.Descendants(W.sdt).FirstOrDefault();
            // get the original element with all the formatting
            var orig = sdt == null ? element.Descendants(W.p).FirstOrDefault() : sdt.Descendants(W.p).FirstOrDefault();

            // check for first run having image element in it
            if (orig == null || !orig.Descendants(W.r).FirstOrDefault().Descendants(W.drawing).Any())
            {
                return CreateContextErrorMessage(element, "Image metadata is not immediately followed by an image", templateError);
            }

            // clone the paragraph, so repeating elements won't be overridden
            var para = new XElement(orig);

            // get the xpath of of the element
            var xPath = (string)element.Attribute(PA.Select);
            // get image path
            var imagePath = EvaluateXPathToString(data, xPath, false);

            // assign unique image and paragraph ids. Image id is document property Id  (wp:docPr)
            // and relationship id is rId. Their numbering is different.
            const string imageId = InvalidImageId; // Ids will be replaced with real ones later, after transform is done
            var relationshipId = GetNextImageRelationshipId(part);

            var inline =
                para.Descendants(W.drawing)
                    .Descendants(WP.inline).FirstOrDefault();
            if (inline == null)
            {
                return CreateContextErrorMessage(element, "Image: invalid picture control", templateError);
            }

            // get aspect ratio option
            var ratioAttr = inline
                .Descendants(WP.cNvGraphicFramePr)
                .Descendants(A.graphicFrameLocks).FirstOrDefault().Attribute(NoNamespace.noChangeAspect);

            var keepSourceImageAspect = (ratioAttr == null);
            var keepOriginalImageSizeElement = inline.Descendants(Pic.cNvPicPr).FirstOrDefault();
            var keepOriginalImageSize = false;

            if (keepOriginalImageSizeElement != null)
            {
                var attr = keepOriginalImageSizeElement.Attribute("preferRelativeResize");
                if (attr != null)
                {
                    keepOriginalImageSize = attr.Value == "0";
                }
            }

            // get extent
            var extent = inline
                .Descendants(WP.extent)
                .FirstOrDefault();
            var pictureExtent = inline
                .Descendants(A.graphic)
                .Descendants(A.graphicData)
                .Descendants(Pic._pic)
                .Descendants(Pic.spPr)
                .Descendants(A.xfrm)
                .Descendants(A.ext).
                FirstOrDefault();

            if (extent == null || pictureExtent == null)
            {
                return CreateContextErrorMessage(element, "Image: missing element in picture control - extent(s)", templateError);
            }

            // get docPr
            var docPr = inline.Descendants(WP.docPr).FirstOrDefault();
            if (docPr == null)
            {
                return CreateContextErrorMessage(element, "Image: missing element in picture control - docPtr", templateError);
            }

            docPr.SetAttributeValue(NoNamespace.id, imageId);
            docPr.SetAttributeValue(NoNamespace.name, "Templated Image Content");

            var blip = inline
                    .Descendants(A.graphic)
                    .Descendants(A.graphicData)
                    .Descendants(Pic.blipFill)
                    .Descendants(A.blip)
                    .FirstOrDefault();

            if (blip != null)
            {
                // Add the image to main document part
                ImagePartType imagePartType;
                string error;
                var stream = Image2Stream(imagePath, out imagePartType, out error);
                if (stream != null)
                {
                    var ip = GetImagePart(part, imagePartType, relationshipId);

                    if (ip == null)
                    {
                        error = "Failed to get image part";
                        return CreateContextErrorMessage(element, string.Concat("Image: ", error), templateError);
                    }
                    ip.FeedData(stream);
                    stream.Close();

                    // access the saved image and get the dimensions
                    using (var savedStream = ip.GetStream(FileMode.Open))
                    using (var image = System.Drawing.Image.FromStream(savedStream))
                    {
                        // one inch is 914400 EMUs
                        // 96dpi where dot is pixel
                        var pixelInEMU = 914400 / 96;
                        var width = image.Width;
                        var height = image.Height;

                        if (keepSourceImageAspect)
                        {
                            var ratio = height / (width * 1.0);
                            if (!int.TryParse(extent.Attribute(NoNamespace.cx).Value, out width))
                            {
                                return CreateContextErrorMessage(element, "Image: Invalid image attributes",
                                    templateError);
                            }
                            height = (int)(width * ratio);

                            // replace attributes
                            extent.SetAttributeValue(NoNamespace.cy, height);
                            pictureExtent.SetAttributeValue(NoNamespace.cx, width);
                            pictureExtent.SetAttributeValue(NoNamespace.cy, height);
                        }

                        if (keepOriginalImageSize)
                        {
                            width = image.Width * pixelInEMU;
                            height = image.Height * pixelInEMU;

                            // replace attributes
                            extent.SetAttributeValue(NoNamespace.cx, width);
                            extent.SetAttributeValue(NoNamespace.cy, height);
                            pictureExtent.SetAttributeValue(NoNamespace.cx, width);
                            pictureExtent.SetAttributeValue(NoNamespace.cy, height);
                        }
                    }
                }
                else
                {
                    return CreateContextErrorMessage(element, string.Concat("Image: ", error), templateError);
                }

                blip.SetAttributeValue(R.embed, relationshipId);
            }

            return para;
        }

        /// <summary>
        /// Determines whether the input image is base64 encoded string or path
        /// </summary>
        /// <param name="inputImage">Input image (either image path or base64 encoded string). Base 64 encoded string
        /// should start with MIME data type identifier followed by raw data. Example:
        /// data:image/jpg;base64,/9j/4AAQSkZJRgAB...</param>
        /// <param name="imagePartType">Image Part Type to be embedded in the document and to be
        /// referenced by image control</param>
        /// <param name="error">Error message</param>
        private static Stream Image2Stream(string inputImage, out ImagePartType imagePartType, out string error)
        {
            string imageType;
            Stream stream;
            if (inputImage.StartsWith("data:image"))
            {
                // assume the image is base64 encoded format. See https://en.wikipedia.org/wiki/Data_URI_scheme

                // get the image type and data
                imageType = Regex.Match(inputImage, @"data:image/(?<type>.+?);").Groups["type"].Value;
                var base64Data = Regex.Match(inputImage, @"data:image/(?<type>.+?),(?<data>.+)").Groups["data"].Value;

                try
                {
                    var imageBytes = Base64.ConvertFromBase64(string.Empty, base64Data);

                    stream = new MemoryStream(imageBytes, 0, imageBytes.Length);
                }
                catch (Exception)
                {
                    imagePartType = default(ImagePartType);
                    error = "Invalid Image data format";
                    return null;
                }
            }
            else
            {
                // assume this is path fo file, so get the extension
                imageType = Path.GetExtension(inputImage).Trim('.');

                try
                {
                    stream = File.Open(inputImage, FileMode.Open);
                }
                catch
                {
                    imagePartType = default(ImagePartType);
                    error = "Invalid Image path";
                    return null;
                }
            }

            switch (imageType)
            {
                case "jpg":
                case "jpeg":
                    imagePartType = ImagePartType.Jpeg;
                    break;
                case "png":
                    imagePartType = ImagePartType.Png;
                    break;
                case "tif":
                case "tiff":
                    imagePartType = ImagePartType.Tiff;
                    break;
                case "bmp":
                    imagePartType = ImagePartType.Bmp;
                    break;
                default:
                    imagePartType = default(ImagePartType);
                    error = "Invalid image type";
                    return null;
            }

            error = string.Empty;
            return stream;
        }

        /// <summary>
        /// Method processes internal paragraphs (marked with a prefix)
        /// </summary>
        /// <param name="element">Source element</param>
        /// <param name="data">Data element with content</param>
        /// <param name="templateError">Error indicator</param>
        /// <returns>Processed element</returns>
        private static object ProcessAParagraph(XElement element, XElement data, TemplateError templateError)
        {
            var para = element.Descendants(A.r).FirstOrDefault();

            var xPath = (string)element.Attribute(PA.Select);
            var optionalString = (string)element.Attribute(PA.Optional);
            var optional = (optionalString != null && optionalString.ToLower() == "true");

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
                var p = new XElement(A.r, para.Elements(A.rPr));
                foreach (var line in newValue.Split('\n'))
                {
                    p.Add(new XElement(A.t,
                        para.Elements(A.t).Elements(A.rPr).FirstOrDefault(),
                        line));
                }
                return p;
            }

            return null;
        }

        static object ContentReplacementTransform(XNode node, XElement data, TemplateError templateError, OpenXmlPart part)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                // TODO: need to figure out potentially better place for handling Alternate Content
                if (element.Name == MC.AlternateContent)
                {
                    // assign new DrawingML object id (for repeated content)
                    var docProperties = element
                        .Descendants(W.drawing)
                        .Descendants(WP.docPr)
                        .FirstOrDefault();
                    if (docProperties != null)
                    {
                        docProperties.SetAttributeValue(NoNamespace.id, InvalidImageId);
                    }

                    // get the fallback picture element
                    var picture = element
                        .Descendants(MC.Fallback)
                        .Descendants(W.pict)
                        .FirstOrDefault();
                    if (picture != null)
                    {
                        // get the shape type element (it's okay not to have it, 
                        // as the shape might use the type defined previously and left
                        // in other shape after copy-paste operation in the editor)
                        var shapeType = picture.Descendants(VML.shapetype).FirstOrDefault();
                        var shape = picture.Descendants(VML.shape).FirstOrDefault();

                        if (shape != null)
                        {
                            shape.SetAttributeValue(NoNamespace.id, GetNextShapeId());

                            if (shapeType != null)
                            {
                                // get next available shape type
                                var spt = GetNextShapeType();
                                var shapeTypeId = string.Format("_x0000_t{0}", spt);

                                // replace the attribute in shape type and in the corresponding shapes
                                shapeType.SetAttributeValue(O.spt, string.Format("{0}", spt));
                                shapeType.SetAttributeValue(NoNamespace.id, shapeTypeId);

                                shape.SetAttributeValue(NoNamespace.type, string.Format("#{0}", shapeTypeId));
                            }
                        }
                    }
                }
                if (element.Name == PA.Image)
                {
                    return ProcessImageContent(element, data, templateError, part);
                }
                if (element.Name == PA.Content)
                {
                    if (element.Descendants(A.r).FirstOrDefault() != null)
                    {
                        return ProcessAParagraph(element, data, templateError);
                    }

                    XElement para = element.Descendants(W.p).FirstOrDefault();
                    XElement run = element.Descendants(W.r).FirstOrDefault();

                    var xPath = (string) element.Attribute(PA.Select);
                    var optionalString = (string) element.Attribute(PA.Optional);
                    bool optional = (optionalString != null && optionalString.ToLower() == "true");

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

                        XElement p = new XElement(W.p, para.Elements(W.pPr));
                        foreach(string line in newValue.Split('\n'))
                        {
                            p.Add(new XElement(W.r,
                                    para.Elements(W.r).Elements(W.rPr).FirstOrDefault(),
                                (p.Elements().Count() > 1) ? new XElement(W.br) : null,
                                new XElement(W.t, line)));
                        }
                        return p;
                    }
                    else
                    {
                        List<XElement> list = new List<XElement>();
                        foreach(string line in newValue.Split('\n'))
                        {
                            list.Add(new XElement(W.r,
                                run.Elements().Where(e => e.Name != W.t),
                                (list.Count > 0) ? new XElement(W.br) : null,
                                new XElement(W.t, line)));
                        }
                        return list;
                    }
                }
                if (element.Name == PA.Repeat)
                {
                    string selector = (string)element.Attribute(PA.Select);
                    var optionalString = (string)element.Attribute(PA.Optional);
                    bool optional = (optionalString != null && optionalString.ToLower() == "true");
                    var alignmentOption = (string)element.Attribute(PA.Align) ?? "vertical";

                    IEnumerable<XElement> repeatingData;
                    try
                    {
                        repeatingData = data.XPathSelectElements(selector);
                    }
                    catch (XPathException e)
                    {
                        return CreateContextErrorMessage(element, "XPathException: " + e.Message, templateError);
                    }
                    if (!repeatingData.Any())
                    {
                        if (optional)
                        {
                            return null;
                            //XElement para = element.Descendants(W.p).FirstOrDefault();
                            //if (para != null)
                            //    return new XElement(W.p, new XElement(W.r));
                            //else
                            //    return new XElement(W.r);
                        }
                        return CreateContextErrorMessage(element, "Repeat: Select returned no data", templateError);
                    }
                    var newContent = repeatingData.Select(d =>
                        {
                            var content = element
                                .Elements()
                                .Select(e => ContentReplacementTransform(e, d, templateError, part))
                                .ToList();
                            return content;
                        })
                        .ToList();
                    switch (alignmentOption.ToLower())
                    {
                        case "horizontal":
                            // keep the properties of first paragraph
                            var pPr = new XElement(W.p, newContent.First())
                                .Elements(W.p)
                                .FirstOrDefault()
                                .Elements(W.pPr)
                                .FirstOrDefault();
                            // create runs from repeated content
                            var runs = newContent.Select(x =>
                            {
                                var run = new XElement(W.p, x);
                                return run.Descendants(W.r).FirstOrDefault();
                            });
                            if (pPr == null)
                            {
                                return new XElement(W.p, runs);
                            }
                            return new XElement(W.p, pPr, runs);
                        case "vertical":
                            return newContent;
                        default:
                            return CreateContextErrorMessage(element, "Repeat: Invalid Align option", templateError);
                    }
                }
                if (element.Name == PA.Table)
                {
                    IEnumerable<XElement> tableData;
                    try
                    {
                        tableData = data.XPathSelectElements((string)element.Attribute(PA.Select));
                    }
                    catch (XPathException e)
                    {
                        return CreateContextErrorMessage(element, "XPathException: " + e.Message, templateError);
                    }
                    if (tableData.Count() == 0)
                        return CreateContextErrorMessage(element, "Table Select returned no data", templateError);
                    XElement table = element.Element(W.tbl);
                    XElement protoRow = table.Elements(W.tr).Skip(1).FirstOrDefault();
                    var footerRowsBeforeTransform = table
                        .Elements(W.tr)
                        .Skip(2)
                        .ToList();
                    var footerRows = footerRowsBeforeTransform
                        .Select(x => ContentReplacementTransform(x, data, templateError, part))
                        .ToList();
                    if (protoRow == null)
                        return CreateContextErrorMessage(element, string.Format("Table does not contain a prototype row"), templateError);
                    protoRow.Descendants(W.bookmarkStart).Remove();
                    protoRow.Descendants(W.bookmarkEnd).Remove();
                    XElement newTable = new XElement(W.tbl,
                        table.Elements().Where(e => e.Name != W.tr),
                        table.Elements(W.tr).FirstOrDefault(),
                        tableData.Select(d =>
                            new XElement(W.tr,
                                protoRow.Elements().Where(r => r.Name != W.tc),
                                protoRow.Elements(W.tc)
                                    .Select(tc =>
                                    {
                                        XElement paragraph = tc.Elements(W.p).FirstOrDefault();

                                        // TODO: to check for other types (if needed, of course). Also, would be nice to refactor it, say, with
                                        // TODO: different condition, for example, with switch case which checks the type of content.
                                        if (paragraph == null)
                                        {
                                            // check if this is emebedded image
                                            var image = tc.Elements(PA.Image).FirstOrDefault();
                                            if (image != null)
                                            {
                                                // has to be wrapped as table cell element, since we are re-formatting the table
                                                return new XElement(W.tc, ProcessImageContent(image, d, templateError, part));
                                            }
                                        }

                                        XElement cellRun = paragraph.Elements(W.r).FirstOrDefault();
                                        string xPath = paragraph.Value;
                                        string newValue = null;
                                        try
                                        {
                                            newValue = EvaluateXPathToString(d, xPath, false);
                                        }
                                        catch (XPathException e)
                                        {
                                            XElement errorCell = new XElement(W.tc,
                                                tc.Elements().Where(z => z.Name != W.p),
                                                new XElement(W.p,
                                                    paragraph.Element(W.pPr),
                                                    CreateRunErrorMessage(e.Message, templateError)));
                                            return errorCell;
                                        }

                                        XElement newCell = new XElement(W.tc,
                                                   tc.Elements().Where(z => z.Name != W.p),
                                                   new XElement(W.p,
                                                       paragraph.Element(W.pPr),
                                                       new XElement(W.r,
                                                           cellRun != null ? cellRun.Element(W.rPr) : new XElement(W.rPr),  //if the cell was empty there is no cellrun
                                                           new XElement(W.t, newValue))));
                                        return newCell;
                                    }))),
                                    footerRows
                                    );
                    return newTable;
                }
                if (element.Name == PA.Conditional)
                {
                    string xPath = (string)element.Attribute(PA.Select);
                    var match = (string)element.Attribute(PA.Match);
                    var notMatch = (string)element.Attribute(PA.NotMatch);

                    if (match == null && notMatch == null)
                        return CreateContextErrorMessage(element, "Conditional: Must specify either Match or NotMatch", templateError);
                    if (match != null && notMatch != null)
                        return CreateContextErrorMessage(element, "Conditional: Cannot specify both Match and NotMatch", templateError);

                    string testValue = null; 
                   
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
                        var content = element.Elements().Select(e => ContentReplacementTransform(e, data, templateError, part));
                        return content;
                    }
                    return null;
                }
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => ContentReplacementTransform(n, data, templateError, part)));
            }
            return node;
        }

        private static object CreateContextErrorMessage(XElement element, string errorMessage, TemplateError templateError)
        {
            XElement para = element.Descendants(W.p).FirstOrDefault();
            XElement run = element.Descendants(W.r).FirstOrDefault();
            var errorRun = CreateRunErrorMessage(errorMessage, templateError);
            if (para != null)
                return new XElement(W.p, errorRun);
            else
                return errorRun;
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

        private static string EvaluateXPathToString(XElement element, string xPath, bool optional )
        {
            object xPathSelectResult;
            try
            {
                //support some cells in the table may not have an xpath expression.
                if (String.IsNullOrWhiteSpace(xPath)) return String.Empty;
                
                xPathSelectResult = element.XPathEvaluate(xPath);
            }
            catch (XPathException e)
            {
                throw new XPathException("XPathException: " + e.Message, e);
            }

            if ((xPathSelectResult is IEnumerable) && !(xPathSelectResult is string))
            {
                var selectedData = ((IEnumerable) xPathSelectResult).Cast<XObject>();
                if (!selectedData.Any())
                {
                    if (optional) return string.Empty;
                    throw new XPathException(string.Format("XPath expression ({0}) returned no results", xPath));
                }
                if (selectedData.Count() > 1)
                {
                    throw new XPathException(string.Format("XPath expression ({0}) returned more than one node", xPath));
                }

                XObject selectedDatum = selectedData.First(); 
                
                if (selectedDatum is XElement) return ((XElement) selectedDatum).Value;

                if (selectedDatum is XAttribute) return ((XAttribute) selectedDatum).Value;
            }

            return xPathSelectResult.ToString();

        }
    }
}
