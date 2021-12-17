﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Xml.Schema;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public partial class WmlDocument
    {
        public WmlDocument SimplifyMarkup(SimplifyMarkupSettings settings)
        {
            return MarkupSimplifier.SimplifyMarkup(this, settings);
        }
    }

    public class SimplifyMarkupSettings
    {
        public bool AcceptRevisions;
        public bool NormalizeXml;
        public bool RemoveBookmarks;
        public bool RemoveComments;
        public bool RemoveContentControls;
        public bool RemoveEndAndFootNotes;
        public bool RemoveFieldCodes;
        public bool RemoveGoBackBookmark;
        public bool RemoveHyperlinks;
        public bool RemoveLastRenderedPageBreak;
        public bool RemoveMarkupForDocumentComparison;
        public bool RemovePermissions;
        public bool RemoveProof;
        public bool RemoveRsidInfo;
        public bool RemoveSmartTags;
        public bool RemoveSoftHyphens;
        public bool RemoveWebHidden;
        public bool ReplaceTabsWithSpaces;
    }

    public static class MarkupSimplifier
    {
        public static WmlDocument SimplifyMarkup(WmlDocument doc, SimplifyMarkupSettings settings)
        {
            using (var streamDoc = new OpenXmlMemoryStreamDocument(doc))
            {
                using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
                {
                    SimplifyMarkup(document, settings);
                }

                return streamDoc.GetModifiedWmlDocument();
            }
        }

        public static void SimplifyMarkup(WordprocessingDocument doc, SimplifyMarkupSettings settings)
        {
            if (settings.RemoveMarkupForDocumentComparison)
            {
                settings.RemoveRsidInfo = true;
                RemoveElementsForDocumentComparison(doc);
            }

            if (settings.RemoveRsidInfo)
            {
                RemoveRsidInfoInSettings(doc);
            }

            if (settings.AcceptRevisions)
            {
                RevisionAccepter.AcceptRevisions(doc);
            }

            foreach (OpenXmlPart part in doc.ContentParts())
            {
                SimplifyMarkupForPart(part, settings);
            }

            if (doc.MainDocumentPart.StyleDefinitionsPart != null)
            {
                SimplifyMarkupForPart(doc.MainDocumentPart.StyleDefinitionsPart, settings);
            }

            if (doc.MainDocumentPart.StylesWithEffectsPart != null)
            {
                SimplifyMarkupForPart(doc.MainDocumentPart.StylesWithEffectsPart, settings);
            }

            if (settings.RemoveComments)
            {
                WordprocessingCommentsPart commentsPart = doc.MainDocumentPart.WordprocessingCommentsPart;
                if (commentsPart != null) doc.MainDocumentPart.DeletePart(commentsPart);

                WordprocessingCommentsExPart commentsExPart = doc.MainDocumentPart.WordprocessingCommentsExPart;
                if (commentsExPart != null) doc.MainDocumentPart.DeletePart(commentsExPart);
            }
        }

        private static void RemoveRsidInfoInSettings(WordprocessingDocument doc)
        {
            DocumentSettingsPart part = doc.MainDocumentPart.DocumentSettingsPart;
            if (part == null) return;

            XDocument settingsXDoc = part.GetXDocument();
            settingsXDoc.Descendants(W.rsids).Remove();
            part.SaveXDocument();
        }

        private static void RemoveElementsForDocumentComparison(WordprocessingDocument doc)
        {
            OpenXmlPart part = doc.ExtendedFilePropertiesPart;
            if (part != null)
            {
                XDocument appPropsXDoc = part.GetXDocument();
                appPropsXDoc.Descendants(EP.TotalTime).Remove();
                part.SaveXDocument();
            }

            part = doc.CoreFilePropertiesPart;
            if (part != null)
            {
                XDocument corePropsXDoc = part.GetXDocument();
                corePropsXDoc.Descendants(CP.revision).Remove();
                corePropsXDoc.Descendants(DCTERMS.created).Remove();
                corePropsXDoc.Descendants(DCTERMS.modified).Remove();
                part.SaveXDocument();
            }

            XDocument mainXDoc = doc.MainDocumentPart.GetXDocument();
            List<XElement> bookmarkStart = mainXDoc
                .Descendants(W.bookmarkStart)
                .Where(b => (string) b.Attribute(W.name) == "_GoBack")
                .ToList();

            foreach (XElement item in bookmarkStart)
            {
                IEnumerable<XElement> bookmarkEnd = mainXDoc
                    .Descendants(W.bookmarkEnd)
                    .Where(be => (int) be.Attribute(W.id) == (int) item.Attribute(W.id));

                bookmarkEnd.Remove();
            }

            bookmarkStart.Remove();
            doc.MainDocumentPart.SaveXDocument();
        }

        public static XElement MergeAdjacentSuperfluousRuns(XElement element)
        {
            return (XElement) MergeAdjacentRunsTransform(element);
        }

        public static XElement TransformElementToSingleCharacterRuns(XElement element)
        {
            return (XElement) SingleCharacterRunTransform(element);
        }

        public static void TransformPartToSingleCharacterRuns(OpenXmlPart part)
        {
            // After transforming to single character runs, Rsid info will be invalid, so
            // remove from the part.
            XElement root = part.GetXElement();
            var newRoot = (XElement) RemoveRsidTransform(root);
            newRoot = (XElement) SingleCharacterRunTransform(newRoot);
            root.ReplaceWith(newRoot);
            part.SaveXElement();
        }

        public static void TransformToSingleCharacterRuns(WordprocessingDocument doc)
        {
            if (RevisionAccepter.HasTrackedRevisions(doc))
            {
                throw new OpenXmlPowerToolsException(
                    "Transforming a document to single character runs is not supported for " +
                    "a document with tracked revisions.");
            }

            foreach (OpenXmlPart part in doc.ContentParts())
            {
                TransformPartToSingleCharacterRuns(part);
            }
        }

        private static object RemoveCustomXmlAndContentControlsTransform(
            XNode node,
            SimplifyMarkupSettings simplifyMarkupSettings)
        {
            if (node is XElement element)
            {
                if (simplifyMarkupSettings.RemoveSmartTags &&
                    element.Name == W.smartTag)
                {
                    return element
                        .Elements()
                        .Select(e =>
                            RemoveCustomXmlAndContentControlsTransform(e,
                                simplifyMarkupSettings));
                }

                if (simplifyMarkupSettings.RemoveContentControls &&
                    element.Name == W.sdt)
                {
                    return element
                        .Elements(W.sdtContent)
                        .Elements()
                        .Select(e =>
                            RemoveCustomXmlAndContentControlsTransform(e,
                                simplifyMarkupSettings));
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => RemoveCustomXmlAndContentControlsTransform(n, simplifyMarkupSettings)));
            }

            return node;
        }

        private static object RemoveRsidTransform(XNode node)
        {
            if (node is XElement element)
            {
                return element.Name == W.rsid
                    ? null
                    : new XElement(element.Name,
                        element
                            .Attributes()
                            .Where(a => a.Name != W.rsid &&
                                        a.Name != W.rsidDel &&
                                        a.Name != W.rsidP &&
                                        a.Name != W.rsidR &&
                                        a.Name != W.rsidRDefault &&
                                        a.Name != W.rsidRPr &&
                                        a.Name != W.rsidSect &&
                                        a.Name != W.rsidTr),
                        element.Nodes().Select(RemoveRsidTransform));
            }

            return node;
        }

        private static object MergeAdjacentRunsTransform(XNode node)
        {
            if (node is XElement element)
            {
                return element.Name == W.p
                    ? WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(element)
                    : new XElement(element.Name,
                        element.Attributes(),
                        element.Nodes().Select(MergeAdjacentRunsTransform));
            }

            return node;
        }

        private static object RemoveEmptyRunsAndRunPropertiesTransform(XNode node)
        {
            if (node is XElement element)
            {
                return (element.Name == W.r || element.Name == W.rPr || element.Name == W.pPr) && !element.Elements().Any()
                    ? null
                    : new XElement(element.Name,
                        element.Attributes(),
                        element.Nodes().Select(RemoveEmptyRunsAndRunPropertiesTransform));
            }

            return node;
        }

        private static object MergeAdjacentInstrText(
            XNode node)
        {
            if (node is XElement element)
            {
                if (element.Name == W.r && element.Elements(W.instrText).Any())
                {
                    IEnumerable<IGrouping<bool, XElement>> grouped = element
                        .Elements()
                        .GroupAdjacent(e => e.Name == W.instrText);

                    return new XElement(W.r,
                        grouped.Select(g =>
                        {
                            if (g.Key == false)
                            {
                                return (object) g;
                            }

                            // If .doc files are converted to .docx by the Binary to Open XML Translator,
                            // the w:instrText elements might be empty, in which case newInstrText would
                            // be an empty string.
                            string newInstrText = g.Select(i => (string) i).StringConcatenate();
                            if (string.IsNullOrEmpty(newInstrText))
                            {
                                return new XElement(W.instrText);
                            }

                            return new XElement(W.instrText,
                                XmlUtil.GetXmlSpaceAttribute(newInstrText),
                                newInstrText);
                        }));
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(MergeAdjacentInstrText));
            }

            return node;
        }

        // lastRenderedPageBreak, permEnd, permStart, proofErr, noProof
        // softHyphen:
        // Remove when simplifying.

        // fldSimple, fldData, fldChar, instrText:
        // For hyperlinks, generate same in XHtml.  Other than hyperlinks, do the following:
        // - collapse fldSimple
        // - remove fldSimple, fldData, fldChar, instrText.

        private static object SimplifyMarkupTransform(
            XNode node,
            SimplifyMarkupSettings settings,
            SimplifyMarkupParameters parameters)
        {
            if (!(node is XElement element)) return node;

            if (settings.RemovePermissions &&
                (element.Name == W.permEnd ||
                 element.Name == W.permStart))
            {
                return null;
            }

            if (settings.RemoveProof &&
                (element.Name == W.proofErr ||
                 element.Name == W.noProof))
            {
                return null;
            }

            if (settings.RemoveSoftHyphens &&
                element.Name == W.softHyphen)
            {
                return null;
            }

            if (settings.RemoveLastRenderedPageBreak &&
                element.Name == W.lastRenderedPageBreak)
            {
                return null;
            }

            if (settings.RemoveBookmarks &&
                (element.Name == W.bookmarkStart ||
                 element.Name == W.bookmarkEnd))
            {
                return null;
            }

            if (settings.RemoveGoBackBookmark &&
                (element.Name == W.bookmarkStart && (int) element.Attribute(W.id) == parameters.GoBackId ||
                 element.Name == W.bookmarkEnd && (int) element.Attribute(W.id) == parameters.GoBackId))
            {
                return null;
            }

            if (settings.RemoveWebHidden &&
                element.Name == W.webHidden)
            {
                return null;
            }

            if (settings.ReplaceTabsWithSpaces &&
                element.Name == W.tab && element.Parent != null && element.Parent.Name == W.r)
            {
                return new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), " ");
            }

            if (settings.RemoveComments &&
                (element.Name == W.commentRangeStart ||
                 element.Name == W.commentRangeEnd ||
                 element.Name == W.commentReference ||
                 element.Name == W.annotationRef))
            {
                return null;
            }

            if (settings.RemoveComments &&
                element.Name == W.rStyle &&
                element.Attribute(W.val)?.Value == "CommentReference")
            {
                return null;
            }

            if (settings.RemoveEndAndFootNotes &&
                (element.Name == W.endnoteReference ||
                 element.Name == W.footnoteReference))
            {
                return null;
            }

            if (settings.RemoveFieldCodes)
            {
                if (element.Name == W.fldSimple)
                {
                    return element.Elements().Select(e => SimplifyMarkupTransform(e, settings, parameters));
                }

                if (element.Name == W.fldData ||
                    element.Name == W.fldChar ||
                    element.Name == W.instrText)
                {
                    return null;
                }
            }

            if (settings.RemoveHyperlinks &&
                element.Name == W.hyperlink)
            {
                return element.Elements();
            }

            return new XElement(element.Name,
                element.Attributes(),
                element.Nodes().Select(n => SimplifyMarkupTransform(n, settings, parameters)));
        }

        private static XDocument Normalize(XDocument source, XmlSchemaSet schema)
        {
            var havePsvi = false;

            // validate, throw errors, add PSVI information
            if (schema != null)
            {
                source.Validate(schema, null, true);
                havePsvi = true;
            }

            // Remove comments, processing instructions, and text nodes that are
            // children of XDocument. Only white space text nodes are allowed as
            // children of a document, so we can remove all text nodes.
            return new XDocument(
                source.Declaration,
                source.Nodes().Select(n =>
                    n is XComment || n is XProcessingInstruction || n is XText
                        ? null
                        : n is XElement e
                            ? NormalizeElement(e, havePsvi)
                            : n));
        }

        // TODO: Check whether this can be removed.
        //private static bool DeepEqualsWithNormalization(XDocument doc1, XDocument doc2, XmlSchemaSet schemaSet)
        //{
        //    XDocument d1 = Normalize(doc1, schemaSet);
        //    XDocument d2 = Normalize(doc2, schemaSet);
        //    return XNode.DeepEquals(d1, d2);
        //}

        private static IEnumerable<XAttribute> NormalizeAttributes(XElement element, bool havePsvi)
        {
            return element
                .Attributes()
                .Where(a => !a.IsNamespaceDeclaration &&
                            a.Name != XSI.schemaLocation &&
                            a.Name != XSI.noNamespaceSchemaLocation)
                .OrderBy(a => a.Name.NamespaceName)
                .ThenBy(a => a.Name.LocalName)
                .Select(a =>
                {
                    if (havePsvi)
                    {
                        IXmlSchemaInfo schemaInfo = a.GetSchemaInfo();
                        XmlSchemaType schemaType = schemaInfo?.SchemaType;
                        XmlTypeCode? typeCode = schemaType?.TypeCode;

                        switch (typeCode)
                        {
                            case XmlTypeCode.Boolean:
                                return new XAttribute(a.Name, (bool) a);
                            case XmlTypeCode.DateTime:
                                return new XAttribute(a.Name, (DateTime) a);
                            case XmlTypeCode.Decimal:
                                return new XAttribute(a.Name, (decimal) a);
                            case XmlTypeCode.Double:
                                return new XAttribute(a.Name, (double) a);
                            case XmlTypeCode.Float:
                                return new XAttribute(a.Name, (float) a);
                            case XmlTypeCode.HexBinary:
                            case XmlTypeCode.Language:
                                return new XAttribute(a.Name, ((string) a).ToLower());
                        }
                    }

                    return a;
                });
        }

        private static XNode NormalizeNode(XNode node, bool havePsvi)
        {
            // trim comments and processing instructions from normalized tree
            if (node is XComment || node is XProcessingInstruction)
            {
                return null;
            }

            if (node is XElement e)
            {
                return NormalizeElement(e, havePsvi);
            }

            // Only thing left is XCData and XText, so clone them
            return node;
        }

        private static XElement NormalizeElement(XElement element, bool havePsvi)
        {
            if (havePsvi)
            {
                IXmlSchemaInfo schemaInfo = element.GetSchemaInfo();
                XmlSchemaType schemaType = schemaInfo?.SchemaType;
                XmlTypeCode? typeCode = schemaType?.TypeCode;

                switch (typeCode)
                {
                    case XmlTypeCode.Boolean:
                        return new XElement(element.Name,
                            NormalizeAttributes(element, true),
                            (bool) element);

                    case XmlTypeCode.DateTime:
                        return new XElement(element.Name,
                            NormalizeAttributes(element, true),
                            (DateTime) element);

                    case XmlTypeCode.Decimal:
                        return new XElement(element.Name,
                            NormalizeAttributes(element, true),
                            (decimal) element);

                    case XmlTypeCode.Double:
                        return new XElement(element.Name,
                            NormalizeAttributes(element, true),
                            (double) element);

                    case XmlTypeCode.Float:
                        return new XElement(element.Name,
                            NormalizeAttributes(element, true),
                            (float) element);

                    case XmlTypeCode.HexBinary:
                    case XmlTypeCode.Language:
                        return new XElement(element.Name,
                            NormalizeAttributes(element, true),
                            ((string) element).ToLower());

                    default:
                        return new XElement(element.Name,
                            NormalizeAttributes(element, true),
                            element.Nodes().Select(n => NormalizeNode(n, true)));
                }
            }

            return new XElement(element.Name,
                NormalizeAttributes(element, false),
                element.Nodes().Select(n => NormalizeNode(n, false)));
        }

        private static void SimplifyMarkupForPart(OpenXmlPart part, SimplifyMarkupSettings settings)
        {
            var parameters = new SimplifyMarkupParameters();
            if (part.ContentType == "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
            {
                var doc = (WordprocessingDocument) part.OpenXmlPackage;
                if (settings.RemoveGoBackBookmark)
                {
                    XElement goBackBookmark = doc
                        .MainDocumentPart
                        .GetXDocument()
                        .Descendants(W.bookmarkStart)
                        .FirstOrDefault(bm => (string) bm.Attribute(W.name) == "_GoBack");

                    if (goBackBookmark != null)
                    {
                        parameters.GoBackId = (int) goBackBookmark.Attribute(W.id);
                    }
                }
            }

            XDocument xdoc = part.GetXDocument();
            XElement newRoot = xdoc.Root;

            // Need to do this first to enable simplifying hyperlinks.
            if (settings.RemoveContentControls || settings.RemoveSmartTags)
            {
                newRoot = (XElement) RemoveCustomXmlAndContentControlsTransform(newRoot, settings);
            }

            // This may touch many elements, so needs to be its own transform.
            if (settings.RemoveRsidInfo)
            {
                newRoot = (XElement) RemoveRsidTransform(newRoot);
            }

            var prevNewRoot = new XDocument(newRoot);
            while (true)
            {
                if (settings.RemoveComments ||
                    settings.RemoveEndAndFootNotes ||
                    settings.ReplaceTabsWithSpaces ||
                    settings.RemoveFieldCodes ||
                    settings.RemovePermissions ||
                    settings.RemoveProof ||
                    settings.RemoveBookmarks ||
                    settings.RemoveWebHidden ||
                    settings.RemoveGoBackBookmark ||
                    settings.RemoveHyperlinks)
                {
                    newRoot = (XElement) SimplifyMarkupTransform(newRoot, settings, parameters);
                }

                // Remove runs and run properties that have become empty due to previous transforms.
                newRoot = (XElement) RemoveEmptyRunsAndRunPropertiesTransform(newRoot);

                // Merge adjacent runs that have identical run properties.
                newRoot = (XElement) MergeAdjacentRunsTransform(newRoot);

                // Merge adjacent instrText elements.
                newRoot = (XElement) MergeAdjacentInstrText(newRoot);

                // Separate run children into separate runs
                newRoot = (XElement) SeparateRunChildrenIntoSeparateRuns(newRoot);

                if (XNode.DeepEquals(prevNewRoot.Root, newRoot))
                {
                    break;
                }

                prevNewRoot = new XDocument(newRoot);
            }

            if (settings.NormalizeXml)
            {
                XAttribute[] nsAttrs =
                {
                    new XAttribute(XNamespace.Xmlns + "wpc", WPC.wpc),
                    new XAttribute(XNamespace.Xmlns + "mc", MC.mc),
                    new XAttribute(XNamespace.Xmlns + "o", O.o),
                    new XAttribute(XNamespace.Xmlns + "r", R.r),
                    new XAttribute(XNamespace.Xmlns + "m", M.m),
                    new XAttribute(XNamespace.Xmlns + "v", VML.vml),
                    new XAttribute(XNamespace.Xmlns + "wp14", WP14.wp14),
                    new XAttribute(XNamespace.Xmlns + "wp", WP.wp),
                    new XAttribute(XNamespace.Xmlns + "w10", W10.w10),
                    new XAttribute(XNamespace.Xmlns + "w", W.w),
                    new XAttribute(XNamespace.Xmlns + "w14", W14.w14),
                    new XAttribute(XNamespace.Xmlns + "w15", W15.w15),
                    new XAttribute(XNamespace.Xmlns + "w16se", W16SE.w16se),
                    new XAttribute(XNamespace.Xmlns + "wpg", WPG.wpg),
                    new XAttribute(XNamespace.Xmlns + "wpi", WPI.wpi),
                    new XAttribute(XNamespace.Xmlns + "wne", WNE.wne),
                    new XAttribute(XNamespace.Xmlns + "wps", WPS.wps),
                    new XAttribute(MC.Ignorable, "w14 wp14 w15 w16se")
                };

                XDocument newXDoc = Normalize(new XDocument(newRoot), null);
                newRoot = newXDoc.Root;
                if (newRoot != null)
                {
                    foreach (XAttribute nsAttr in nsAttrs)
                    {
                        if (newRoot.Attribute(nsAttr.Name) == null)
                        {
                            newRoot.Add(nsAttr);
                        }
                    }
                }

                part.SetXDocument(newXDoc);
            }
            else
            {
                part.SetXDocument(new XDocument(newRoot));
            }
        }

        private static object SeparateRunChildrenIntoSeparateRuns(XNode node)
        {
            if (node is XElement element)
            {
                if (element.Name == W.r)
                {
                    IEnumerable<XElement> runChildren = element.Elements().Where(e => e.Name != W.rPr);
                    XElement rPr = element.Element(W.rPr);
                    return runChildren.Select(rc => new XElement(W.r, rPr, rc));
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(SeparateRunChildrenIntoSeparateRuns));
            }

            return node;
        }

        private static object SingleCharacterRunTransform(XNode node)
        {
            if (node is XElement element)
            {
                if (element.Name == W.r)
                    return element.Elements()
                        .Where(e => e.Name != W.rPr)
                        .GroupAdjacent(sr => sr.Name == W.t)
                        .Select(g =>
                        {
                            if (g.Key)
                            {
                                string s = g.Select(t => (string) t).StringConcatenate();
                                return s.Select(c =>
                                    new XElement(W.r,
                                        element.Elements(W.rPr),
                                        new XElement(W.t,
                                            XmlUtil.GetXmlSpaceAttribute(c),
                                            c)));
                            }

                            return g.Select(sr =>
                                new XElement(W.r,
                                    element.Elements(W.rPr),
                                    new XElement(sr.Name,
                                        sr.Attributes(),
                                        sr.Nodes().Select(SingleCharacterRunTransform))));
                        });

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(SingleCharacterRunTransform));
            }

            return node;
        }

        private class SimplifyMarkupParameters
        {
            public int? GoBackId { get; set; }
        }
    }
}
