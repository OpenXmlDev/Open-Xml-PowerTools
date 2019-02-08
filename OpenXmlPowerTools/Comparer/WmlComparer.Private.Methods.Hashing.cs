// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public static partial class WmlComparer
    {
        private static WmlDocument HashBlockLevelContent(
            WmlDocument source,
            WmlDocument sourceAfterProc,
            WmlComparerSettings settings)
        {
            using (var msSource = new MemoryStream())
            using (var msAfterProc = new MemoryStream())
            {
                msSource.Write(source.DocumentByteArray, 0, source.DocumentByteArray.Length);
                msAfterProc.Write(sourceAfterProc.DocumentByteArray, 0, sourceAfterProc.DocumentByteArray.Length);

                using (WordprocessingDocument wDocSource = WordprocessingDocument.Open(msSource, true))
                using (WordprocessingDocument wDocAfterProc = WordprocessingDocument.Open(msAfterProc, true))
                {
                    // create Unid dictionary for source
                    XDocument sourceMainXDoc = wDocSource.MainDocumentPart.GetXDocument();
                    XElement sourceMainRoot = sourceMainXDoc.Root ?? throw new ArgumentException();
                    Dictionary<string, XElement> sourceUnidDict = sourceMainRoot
                        .Descendants()
                        .Where(d => d.Name == W.p || d.Name == W.tbl || d.Name == W.tr)
                        .ToDictionary(d => (string) d.Attribute(PtOpenXml.Unid));

                    XDocument afterProcMainXDoc = wDocAfterProc.MainDocumentPart.GetXDocument();
                    XElement afterProcMainRoot = afterProcMainXDoc.Root ?? throw new ArgumentException();
                    IEnumerable<XElement> blockLevelElements = afterProcMainRoot
                        .Descendants()
                        .Where(d => d.Name == W.p || d.Name == W.tbl || d.Name == W.tr);

                    foreach (XElement blockLevelContent in blockLevelElements)
                    {
                        var cloneBlockLevelContentForHashing = (XElement) CloneBlockLevelContentForHashing(
                            wDocAfterProc.MainDocumentPart,
                            blockLevelContent,
                            true,
                            settings);

                        string shaString = cloneBlockLevelContentForHashing
                            .ToString(SaveOptions.DisableFormatting)
                            .Replace(" xmlns=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", "");

                        string sha1Hash = WmlComparerUtil.SHA1HashStringForUTF8String(shaString);
                        var thisUnid = (string) blockLevelContent.Attribute(PtOpenXml.Unid);
                        if (thisUnid != null)
                        {
                            if (sourceUnidDict.ContainsKey(thisUnid))
                            {
                                XElement correlatedBlockLevelContent = sourceUnidDict[thisUnid];
                                correlatedBlockLevelContent.Add(new XAttribute(PtOpenXml.CorrelatedSHA1Hash, sha1Hash));
                            }
                        }
                    }

                    wDocSource.MainDocumentPart.PutXDocument();
                }

                var sourceWithCorrelatedSHA1Hash = new WmlDocument(source.FileName, msSource.ToArray());
                return sourceWithCorrelatedSHA1Hash;
            }
        }

        // prohibit
        // - altChunk
        // - subDoc
        // - contentPart

        // This strips all text nodes from the XML tree, thereby leaving only the structure.

        private static object CloneBlockLevelContentForHashing(
            OpenXmlPart mainDocumentPart,
            XNode node,
            bool includeRelatedParts,
            WmlComparerSettings settings)
        {
            if (node is XElement element)
            {
                if (element.Name == W.bookmarkStart ||
                    element.Name == W.bookmarkEnd ||
                    element.Name == W.pPr ||
                    element.Name == W.rPr)
                {
                    return null;
                }

                if (element.Name == W.p)
                {
                    var clonedPara = new XElement(element.Name,
                        element.Attributes().Where(a => a.Name != W.rsid &&
                                                        a.Name != W.rsidDel &&
                                                        a.Name != W.rsidP &&
                                                        a.Name != W.rsidR &&
                                                        a.Name != W.rsidRDefault &&
                                                        a.Name != W.rsidRPr &&
                                                        a.Name != W.rsidSect &&
                                                        a.Name != W.rsidTr &&
                                                        a.Name.Namespace != PtOpenXml.pt),
                        element.Nodes().Select(n =>
                            CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts, settings)));

                    IEnumerable<IGrouping<bool, XElement>> groupedRuns = clonedPara
                        .Elements()
                        .GroupAdjacent(e => e.Name == W.r &&
                                            e.Elements().Count() == 1 &&
                                            e.Element(W.t) != null);

                    var clonedParaWithGroupedRuns = new XElement(element.Name,
                        groupedRuns.Select(g =>
                        {
                            if (g.Key)
                            {
                                string text = g.Select(t => t.Value).StringConcatenate();
                                if (settings.CaseInsensitive)
                                    text = text.ToUpper(settings.CultureInfo);
                                var newRun = (object) new XElement(W.r,
                                    new XElement(W.t,
                                        text));
                                return newRun;
                            }

                            return g;
                        }));

                    return clonedParaWithGroupedRuns;
                }

                if (element.Name == W.r)
                {
                    IEnumerable<XElement> clonedRuns = element
                        .Elements()
                        .Where(e => e.Name != W.rPr)
                        .Select(rc => new XElement(W.r,
                            CloneBlockLevelContentForHashing(mainDocumentPart, rc, includeRelatedParts, settings)));
                    return clonedRuns;
                }

                if (element.Name == W.tbl)
                {
                    var clonedTable = new XElement(W.tbl,
                        element.Elements(W.tr).Select(n =>
                            CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts, settings)));
                    return clonedTable;
                }

                if (element.Name == W.tr)
                {
                    var clonedRow = new XElement(W.tr,
                        element.Elements(W.tc).Select(n =>
                            CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts, settings)));
                    return clonedRow;
                }

                if (element.Name == W.tc)
                {
                    var clonedCell = new XElement(W.tc,
                        element.Elements().Select(n =>
                            CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts, settings)));
                    return clonedCell;
                }

                if (element.Name == W.tcPr)
                {
                    var clonedCellProps = new XElement(W.tcPr,
                        element.Elements(W.gridSpan).Select(n =>
                            CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts, settings)));
                    return clonedCellProps;
                }

                if (element.Name == W.gridSpan)
                {
                    var clonedGridSpan = new XElement(W.gridSpan,
                        new XAttribute("val", (string) element.Attribute(W.val)));
                    return clonedGridSpan;
                }

                if (element.Name == W.txbxContent)
                {
                    var clonedTextbox = new XElement(W.txbxContent,
                        element.Elements().Select(n =>
                            CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts, settings)));
                    return clonedTextbox;
                }

                if (includeRelatedParts)
                {
                    if (ComparisonUnitWord.ElementsWithRelationshipIds.Contains(element.Name))
                    {
                        var newElement = new XElement(element.Name,
                            element.Attributes()
                                .Where(a => a.Name.Namespace != PtOpenXml.pt)
                                .Where(a => !AttributesToTrimWhenCloning.Contains(a.Name))
                                .Select(a =>
                                {
                                    if (!ComparisonUnitWord.RelationshipAttributeNames.Contains(a.Name))
                                        return a;

                                    var rId = (string) a;

                                    // could be an hyperlink relationship
                                    try
                                    {
                                        OpenXmlPart oxp = mainDocumentPart.GetPartById(rId);
                                        if (oxp == null)
                                            throw new FileFormatException("Invalid WordprocessingML Document");

                                        var anno = oxp.Annotation<PartSHA1HashAnnotation>();
                                        if (anno != null)
                                            return new XAttribute(a.Name, anno.Hash);

                                        if (!oxp.ContentType.EndsWith("xml"))
                                        {
                                            using (Stream str = oxp.GetStream())
                                            {
                                                byte[] ba;
                                                using (var br = new BinaryReader(str))
                                                {
                                                    ba = br.ReadBytes((int) str.Length);
                                                }

                                                string sha1 = WmlComparerUtil.SHA1HashStringForByteArray(ba);
                                                oxp.AddAnnotation(new PartSHA1HashAnnotation(sha1));
                                                return new XAttribute(a.Name, sha1);
                                            }
                                        }
                                    }
                                    catch (ArgumentOutOfRangeException)
                                    {
                                        HyperlinkRelationship hr =
                                            mainDocumentPart.HyperlinkRelationships.FirstOrDefault(z => z.Id == rId);
                                        if (hr != null)
                                        {
                                            string str = hr.Uri.ToString();
                                            return new XAttribute(a.Name, str);
                                        }

                                        // could be an external relationship
                                        ExternalRelationship er =
                                            mainDocumentPart.ExternalRelationships.FirstOrDefault(z => z.Id == rId);
                                        if (er != null)
                                        {
                                            string str = er.Uri.ToString();
                                            return new XAttribute(a.Name, str);
                                        }

                                        return new XAttribute(a.Name, "NULL Relationship");
                                    }

                                    return null;
                                }),
                            element.Nodes().Select(n =>
                                CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts, settings)));
                        return newElement;
                    }
                }

                if (element.Name == VML.shape)
                {
                    return new XElement(element.Name,
                        element.Attributes()
                            .Where(a => a.Name.Namespace != PtOpenXml.pt)
                            .Where(a => a.Name != "style" && a.Name != "id" && a.Name != "type"),
                        element.Nodes().Select(n =>
                            CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts, settings)));
                }

                if (element.Name == O.OLEObject)
                {
                    var o = new XElement(element.Name,
                        element.Attributes()
                            .Where(a => a.Name.Namespace != PtOpenXml.pt)
                            .Where(a => a.Name != "ObjectID" && a.Name != R.id),
                        element.Nodes().Select(n =>
                            CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts, settings)));
                    return o;
                }

                if (element.Name == W._object)
                {
                    var o = new XElement(element.Name,
                        element.Attributes()
                            .Where(a => a.Name.Namespace != PtOpenXml.pt),
                        element.Nodes().Select(n =>
                            CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts, settings)));
                    return o;
                }

                if (element.Name == WP.docPr)
                {
                    return new XElement(element.Name,
                        element.Attributes()
                            .Where(a => a.Name.Namespace != PtOpenXml.pt && a.Name != "id"),
                        element.Nodes().Select(n =>
                            CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts, settings)));
                }

                return new XElement(element.Name,
                    element.Attributes()
                        .Where(a => a.Name.Namespace != PtOpenXml.pt)
                        .Where(a => !AttributesToTrimWhenCloning.Contains(a.Name)),
                    element.Nodes().Select(n =>
                        CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts, settings)));
            }

            if (settings.CaseInsensitive)
            {
                if (node is XText xt)
                {
                    string newText = xt.Value.ToUpper(settings.CultureInfo);
                    return new XText(newText);
                }
            }

            return node;
        }
    }
}
