/***************************************************************************

Copyright (c) Eric White 2016.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://EricWhite.com
Resource Center and Documentation: http://ericwhite.com/blog/blog/open-xml-powertools-developer-center/

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

***************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.IO.Packaging;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using System.Drawing;
using System.Security.Cryptography;
using OpenXmlPowerTools;

// It is possible to optimize DescendantContentAtoms

namespace OpenXmlPowerTools
{
    public class WmlComparerSettings
    {
        public char[] WordSeparators;
        public string AuthorForRevisions = "Open-Xml-PowerTools";
        public string DateTimeForRevisions = DateTime.Now.ToString("o");
        public double DetailThreshold = 0.15;

        public WmlComparerSettings()
        {
            // note that , and . are processed explicitly to handle cases where they are in a number or word
            WordSeparators = new[] { ' ', '-' }; // todo need to fix this for complete list
        }
    }

    public class WmlRevisedDocumentInfo
    {
        public WmlDocument RevisedDocument;
        public string Revisor;
        public Color Color;
    }

    public static class WmlComparer
    {
        public static bool s_False = false;
        public static bool s_True = true;

        public static WmlDocument Compare(WmlDocument source1, WmlDocument source2, WmlComparerSettings settings)
        {
            return CompareInternal(source1, source2, settings, true);
        }

        private static WmlDocument CompareInternal(WmlDocument source1, WmlDocument source2, WmlComparerSettings settings,
            bool preProcessMarkupInOriginal)
        {
            if (preProcessMarkupInOriginal)
                source1 = PreProcessMarkup(source1);
            source2 = PreProcessMarkup(source2);

            // need to call ChangeFootnoteEndnoteReferencesToGuids before creating the wmlResult document, so that
            // the same GUID ids are used for footnote and endnote references in both the 'after' document, and in the
            // result document.
            WmlDocument wmlResult = new WmlDocument(source1);
            using (MemoryStream ms1 = new MemoryStream())
            using (MemoryStream ms2 = new MemoryStream())
            {
                ms1.Write(source1.DocumentByteArray, 0, source1.DocumentByteArray.Length);
                ms2.Write(source2.DocumentByteArray, 0, source2.DocumentByteArray.Length);
                WmlDocument producedDocument;
                using (WordprocessingDocument wDoc1 = WordprocessingDocument.Open(ms1, true))
                using (WordprocessingDocument wDoc2 = WordprocessingDocument.Open(ms2, true))
                {
                    producedDocument = ProduceDocumentWithTrackedRevisions(settings, wmlResult, wDoc1, wDoc2);
                    return producedDocument;
                }
            }
        }

        private static WmlDocument PreProcessMarkup(WmlDocument source)
        {
            // open and close to get rid of MC content
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(source.DocumentByteArray, 0, source.DocumentByteArray.Length);
                OpenSettings os = new OpenSettings();
                os.MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts,
                    DocumentFormat.OpenXml.FileFormatVersions.Office2007);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true, os))
                {
                    var doc = wDoc.MainDocumentPart.RootElement;
                    if (wDoc.MainDocumentPart.FootnotesPart != null)
                    {
                        // contrary to what you might think, looking at the API, it is necessary to access the root element of each part to cause
                        // the SDK to process MC markup.
                        var fn = wDoc.MainDocumentPart.FootnotesPart.RootElement;
                    }
                    if (wDoc.MainDocumentPart.EndnotesPart != null)
                    {
                        var en = wDoc.MainDocumentPart.EndnotesPart.RootElement;
                    }
                }
                source = new WmlDocument("x.docx", ms.ToArray());
            }

            // open and close to get rid of MC content
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(source.DocumentByteArray, 0, source.DocumentByteArray.Length);
                OpenSettings os = new OpenSettings();
                os.MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts,
                    DocumentFormat.OpenXml.FileFormatVersions.Office2007);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true, os))
                {
                    TestForInvalidContent(wDoc);
                    RemoveExistingPowerToolsMarkup(wDoc);
                    SimplifyMarkupSettings msSettings = new SimplifyMarkupSettings()
                    {
                        RemoveBookmarks = true,
                        AcceptRevisions = true,
                        RemoveComments = true,
                        RemoveContentControls = true,
                        RemoveFieldCodes = true,
                        RemoveGoBackBookmark = true,
                        RemoveLastRenderedPageBreak = true,
                        RemovePermissions = true,
                        RemoveProof = true,
                        RemoveSmartTags = true,
                        RemoveSoftHyphens = true,
                        RemoveHyperlinks = true,
                    };
                    MarkupSimplifier.SimplifyMarkup(wDoc, msSettings);
                    ChangeFootnoteEndnoteReferencesToGuids(wDoc);
                    AddUnidsToMarkupInContentParts(wDoc);
                    AddFootnotesEndnotesParts(wDoc);
                }
                return new WmlDocument("x.docx", ms.ToArray());
            }
        }

        private static void AddUnidsToMarkupInContentParts(WordprocessingDocument wDoc)
        {
            var mdp = wDoc.MainDocumentPart.GetXDocument();
            AssignUnidToAllElements(mdp.Root);
            IgnorePt14Namespace(mdp.Root);
            wDoc.MainDocumentPart.PutXDocument();

            if (wDoc.MainDocumentPart.FootnotesPart != null)
            {
                var p = wDoc.MainDocumentPart.FootnotesPart.GetXDocument();
                AssignUnidToAllElements(p.Root);
                IgnorePt14Namespace(p.Root);
                wDoc.MainDocumentPart.FootnotesPart.PutXDocument();
            }

            if (wDoc.MainDocumentPart.EndnotesPart != null)
            {
                var p = wDoc.MainDocumentPart.EndnotesPart.GetXDocument();
                AssignUnidToAllElements(p.Root);
                IgnorePt14Namespace(p.Root);
                wDoc.MainDocumentPart.EndnotesPart.PutXDocument();
            }
        }

        private class ConsolidationInfo
        {
            public string Revisor;
            public Color Color;
            public XElement RevisionElement;
            public bool InsertBefore = false;
            public string RevisionHash;
        }

        /*****************************************************************************************************************/
        // Consolidate does not process deltas in endnotes and footnotes.  In a future version, this method should go into
        // each footnote and endnote, and process deltas, inserting tables for each delta.  This is not part of the original
        // spec, and there is no time to implement this, so at this point, Consolidate does not implement this.
        /*****************************************************************************************************************/
        public static WmlDocument Consolidate(WmlDocument original,
            List<WmlRevisedDocumentInfo> revisedDocumentInfoList,
            WmlComparerSettings settings)
        {
            // pre-process the original, so that it already has unids for all elements
            // then when comparing all documents to the original, each one will have the prev unid as appropriate
            // for all revision block-level content
            //   set prevUnid to look for
            //   while true
            //     determine where to insert
            //       get the prevUnid for the revision
            //       look it up in the original.  if find it, then insert after that element
            //       if not in the original
            //         look backwards in revised document, set prevUnid to look for, do the loop again
            //       if get to the beginning of the document
            //         insert at beginning of document

            var originalWithUnids = PreProcessMarkup(original);
            WmlDocument consolidated = new WmlDocument(originalWithUnids);

            var revisedDocumentInfoListCount = revisedDocumentInfoList.Count();

            using (MemoryStream consolidatedMs = new MemoryStream())
            {
                consolidatedMs.Write(consolidated.DocumentByteArray, 0, consolidated.DocumentByteArray.Length);
                using (WordprocessingDocument consolidatedWDoc = WordprocessingDocument.Open(consolidatedMs, true))
                {
                    var consolidatedMainDocPart = consolidatedWDoc.MainDocumentPart;
                    var consolidatedMainDocPartXDoc = consolidatedMainDocPart.GetXDocument();

                    // save away last sectPr
                    XElement savedSectPr = consolidatedMainDocPartXDoc
                        .Root
                        .Element(W.body)
                        .Elements(W.sectPr)
                        .LastOrDefault();
                    consolidatedMainDocPartXDoc
                        .Root
                        .Element(W.body)
                        .Elements(W.sectPr)
                        .Remove();

                    var consolidatedByPrevUnid = consolidatedMainDocPartXDoc
                        .Descendants()
                        .Where(d => (d.Name == W.p || d.Name == W.tbl) && d.Attribute(PtOpenXml.PrevUnid) != null)
                        .ToDictionary(d => (string)d.Attribute(PtOpenXml.PrevUnid));

                    foreach (var revisedDocumentInfo in revisedDocumentInfoList)
                    {
                        var delta = WmlComparer.CompareInternal(originalWithUnids, revisedDocumentInfo.RevisedDocument, settings, false);

                        var colorRgb = revisedDocumentInfo.Color.ToArgb();
                        var colorString = colorRgb.ToString("X");
                        if (colorString.Length == 8)
                            colorString = colorString.Substring(2);

                        using (MemoryStream msOriginalWithUnids = new MemoryStream())
                        using (MemoryStream msDelta = new MemoryStream())
                        {
                            msOriginalWithUnids.Write(originalWithUnids.DocumentByteArray, 0, originalWithUnids.DocumentByteArray.Length);
                            msDelta.Write(delta.DocumentByteArray, 0, delta.DocumentByteArray.Length);
                            using (WordprocessingDocument wDocOriginalWithUnids = WordprocessingDocument.Open(msOriginalWithUnids, true))
                            using (WordprocessingDocument wDocDelta = WordprocessingDocument.Open(msDelta, true))
                            {
                                var modMainDocPart = wDocDelta.MainDocumentPart;
                                var modMainDocPartXDoc = modMainDocPart.GetXDocument();
                                var blockLevelContentToMove = modMainDocPartXDoc
                                    .Root
                                    .DescendantsTrimmed(d => d.Name == W.txbxContent || d.Name == W.tr)
                                    .Where(d => d.Name == W.p || d.Name == W.tbl)
                                    .Where(d => d.Descendants().Any(z => z.Name == W.ins || z.Name == W.del))
                                    .ToList();

                                foreach (var revision in blockLevelContentToMove)
                                {
                                    var elementLookingAt = revision;
                                    while (true)
                                    {
                                        var prevUnid = (string)elementLookingAt.Attribute(PtOpenXml.PrevUnid);
                                        if (prevUnid == null)
                                            throw new OpenXmlPowerToolsException("Internal error");

                                        XElement elementToInsertAfter = null;
                                        if (consolidatedByPrevUnid.ContainsKey(prevUnid))
                                            elementToInsertAfter = consolidatedByPrevUnid[prevUnid];

                                        if (elementToInsertAfter != null)
                                        {
                                            ConsolidationInfo ci = new ConsolidationInfo();
                                            ci.Revisor = revisedDocumentInfo.Revisor;
                                            ci.Color = revisedDocumentInfo.Color;
                                            ci.RevisionElement = revision;
                                            AddToAnnotation(
                                                wDocDelta,
                                                consolidatedWDoc,
                                                elementToInsertAfter,
                                                ci);
                                            break;
                                        }
                                        else
                                        {
                                            // find an element to insert after
                                            var elementBeforeRevision = elementLookingAt
                                                .SiblingsBeforeSelfReverseDocumentOrder()
                                                .FirstOrDefault(e => e.Attribute(PtOpenXml.PrevUnid) != null);
                                            if (elementBeforeRevision == null)
                                            {
                                                var firstElement = consolidatedMainDocPartXDoc
                                                    .Root
                                                    .Element(W.body)
                                                    .Elements()
                                                    .FirstOrDefault(e => e.Name == W.p || e.Name == W.tbl);

                                                ConsolidationInfo ci = new ConsolidationInfo();
                                                ci.Revisor = revisedDocumentInfo.Revisor;
                                                ci.Color = revisedDocumentInfo.Color;
                                                ci.RevisionElement = revision;
                                                ci.InsertBefore = true;
                                                AddToAnnotation(
                                                    wDocDelta,
                                                    consolidatedWDoc,
                                                    firstElement,
                                                    ci);
                                                break;
                                            }
                                            else
                                            {
                                                elementLookingAt = elementBeforeRevision;
                                                continue;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // at this point, everything is added as an annotation, from all documents to be merged.
                    // so now the process is to go through and add the annotations to the document
                    var elementsToProcess = consolidatedMainDocPartXDoc
                        .Root
                        .Descendants()
                        .Where(d => d.Annotation<List<ConsolidationInfo>>() != null)
                        .ToList();

                    var emptyParagraph = new XElement(W.p,
                        new XElement(W.pPr,
                            new XElement(W.spacing,
                                new XAttribute(W.after, "0"),
                                new XAttribute(W.line, "240"),
                                new XAttribute(W.lineRule, "auto"))));

                    foreach (var ele in elementsToProcess)
                    {
                        var lci = ele.Annotation<List<ConsolidationInfo>>();

                        // if all revisions from all revisors are exactly the same, then instead of adding multiple tables after
                        // that contains the revisions, then simply replace the paragraph with the one with the revisions.
                        // RC004 documents contain the test data to exercise this.

                        var lciCount = lci.Count();
                        if (lci.Count() > 1 && lciCount == revisedDocumentInfoListCount)
                        {
                            var uniqueRevisionCount = lci
                                .GroupAdjacent(ci => ci.InsertBefore.ToString() + ci.RevisionHash)
                                .Count();
                            if (uniqueRevisionCount == 1)
                            {
                                ele.ReplaceWith(lci.First().RevisionElement);
                                continue;
                            }
                        }
                        var contentToAddBefore = lci
                            .Where(ci => ci.InsertBefore == true)
                            .GroupAdjacent(ci => ci.Revisor + ci.Color.ToString())
                            .Select((groupedCi, idx) => AssembledConjoinedRevisionContent(emptyParagraph, groupedCi, idx));
                        var contentToAddAfter = lci
                            .Where(ci => ci.InsertBefore == false)
                            .GroupAdjacent(ci => ci.Revisor + ci.Color.ToString())
                            .Select((groupedCi, idx) => AssembledConjoinedRevisionContent(emptyParagraph, groupedCi, idx));
                        ele.AddBeforeSelf(contentToAddBefore);
                        ele.AddAfterSelf(contentToAddAfter);
                    }

                    consolidatedMainDocPartXDoc
                        .Root
                        .Element(W.body)
                        .Add(savedSectPr);

                    AddTableGridStyleToStylesPart(consolidatedWDoc.MainDocumentPart.StyleDefinitionsPart);

                    FixUpRevisionIds(consolidatedWDoc, consolidatedMainDocPartXDoc);
                    FixUpEndnoteFootnoteIds(consolidatedWDoc, consolidatedMainDocPartXDoc);
                    FixUpDocPrIds(consolidatedWDoc);
                    FixUpShapeIds(consolidatedWDoc);
                    FixUpShapeTypeIds(consolidatedWDoc);

                    WmlComparer.IgnorePt14Namespace(consolidatedMainDocPartXDoc.Root);

                    consolidatedWDoc.MainDocumentPart.PutXDocument();
                }

                var newConsolidatedDocument = new WmlDocument("consolidated.docx", consolidatedMs.ToArray());
                return newConsolidatedDocument;
            }
        }

        private static void FixUpEndnoteFootnoteIds(WordprocessingDocument wDoc, XDocument mainDocumentXDoc)
        {
            var footnotesPart = wDoc.MainDocumentPart.FootnotesPart;
            var endnotesPart = wDoc.MainDocumentPart.EndnotesPart;

            XDocument footnotesPartXDoc = null;
            if (footnotesPart != null)
                footnotesPartXDoc = footnotesPart.GetXDocument();

            XDocument endnotesPartXDoc = null;
            if (endnotesPart != null)
                endnotesPartXDoc = endnotesPart.GetXDocument();

            var footnotesRefs = mainDocumentXDoc
                .Descendants(W.footnoteReference)
                .Select((fn, idx) =>
                {
                    return new
                    {
                        FootNote = fn,
                        Idx = idx,
                    };
                });

            foreach (var fn in footnotesRefs)
            {
                var oldId = (string)fn.FootNote.Attribute(W.id);
                var newId = (fn.Idx + 1).ToString();
                fn.FootNote.Attribute(W.id).Value = newId;
                var footnote = footnotesPartXDoc
                    .Root
                    .Elements()
                    .FirstOrDefault(e => (string)e.Attribute(W.id) == oldId);
                if (footnote == null)
                    throw new OpenXmlPowerToolsException("Internal error");
                footnote.Attribute(W.id).Value = newId;
            }

            var endnotesRefs = mainDocumentXDoc
                .Descendants(W.endnoteReference)
                .Select((fn, idx) =>
                {
                    return new
                    {
                        EndNote = fn,
                        Idx = idx,
                    };
                });

            foreach (var fn in endnotesRefs)
            {
                var oldId = (string)fn.EndNote.Attribute(W.id);
                var newId = (fn.Idx + 1).ToString();
                fn.EndNote.Attribute(W.id).Value = newId;
                var endnote = endnotesPartXDoc
                    .Root
                    .Elements()
                    .FirstOrDefault(e => (string)e.Attribute(W.id) == oldId);
                if (endnote == null)
                    throw new OpenXmlPowerToolsException("Internal error");
                endnote.Attribute(W.id).Value = newId;
            }

            if (footnotesPart != null)
                footnotesPart.PutXDocument();

            if (endnotesPart != null)
                endnotesPart.PutXDocument();
        }

        private static XElement[] AssembledConjoinedRevisionContent(XElement emptyParagraph, IGrouping<string, ConsolidationInfo> groupedCi, int idx)
        {
            var captionParagraph = new XElement(W.p,
                new XElement(W.pPr,
                    new XElement(W.jc, new XAttribute(W.val, "both")),
                    new XElement(W.rPr,
                        new XElement(W.b),
                        new XElement(W.bCs))),
                new XElement(W.r,
                    new XElement(W.rPr,
                        new XElement(W.b),
                        new XElement(W.bCs)),
                    new XElement(W.t, groupedCi.First().Revisor)));

            var colorRgb = groupedCi.First().Color.ToArgb();
            var colorString = colorRgb.ToString("X");
            if (colorString.Length == 8)
                colorString = colorString.Substring(2);

            var table = new XElement(W.tbl,
                new XElement(W.tblPr,
                    new XElement(W.tblStyle, new XAttribute(W.val, "TableGridForRevisions")),
                    new XElement(W.tblW,
                        new XAttribute(W._w, "0"),
                        new XAttribute(W.type, "auto")),
                    new XElement(W.shd,
                        new XAttribute(W.val, "clear"),
                        new XAttribute(W.color, "auto"),
                        new XAttribute(W.fill, colorString)),
                    new XElement(W.tblLook,
                        new XAttribute(W.firstRow, "0"),
                        new XAttribute(W.lastRow, "0"),
                        new XAttribute(W.firstColumn, "0"),
                        new XAttribute(W.lastColumn, "0"),
                        new XAttribute(W.noHBand, "0"),
                        new XAttribute(W.noVBand, "0"))),
                new XElement(W.tblGrid,
                    new XElement(W.gridCol, new XAttribute(W._w, "9576"))),
                new XElement(W.tr,
                    new XElement(W.tc,
                        new XElement(W.tcPr,
                        new XElement(W.shd,
                            new XAttribute(W.val, "clear"),
                            new XAttribute(W.color, "auto"),
                            new XAttribute(W.fill, colorString))),
                        captionParagraph,
                        groupedCi.Select(ci =>
                            {
                                XElement paraAfter = null;
                                if (ci.RevisionElement.Name == W.tbl)
                                    paraAfter = emptyParagraph;
                                return new [] {
                                    ci.RevisionElement,
                                    paraAfter,
                                };
                            }))));

            var content = new[] {
                                    idx == 0 ? emptyParagraph : null,
                                    table,
                                    emptyParagraph,
                                };

            return content;
        }

        private static void AddToAnnotation(
            WordprocessingDocument wDocDelta,
            WordprocessingDocument consolidatedWDoc,
            XElement elementToInsertAfter,
            ConsolidationInfo consolidationInfo)
        {
            // the following removes footnotes / endnotes; we do not put them in the revision tables that follow revised paragraphs.
            XElement cleanedUpRevision = (XElement)CleanUpRevisionTransform(consolidationInfo.RevisionElement);

            Package packageOfDeletedContent = wDocDelta.MainDocumentPart.OpenXmlPackage.Package;
            Package packageOfNewContent = consolidatedWDoc.MainDocumentPart.OpenXmlPackage.Package;
            PackagePart partInDeletedDocument = packageOfDeletedContent.GetPart(wDocDelta.MainDocumentPart.Uri);
            PackagePart partInNewDocument = packageOfNewContent.GetPart(consolidatedWDoc.MainDocumentPart.Uri);
            consolidationInfo.RevisionElement = MoveRelatedPartsToDestination(partInDeletedDocument, partInNewDocument, cleanedUpRevision);

            var clonedForHashing = (XElement)CloneBlockLevelContentForHashing(consolidatedWDoc.MainDocumentPart, consolidationInfo.RevisionElement, false);
            clonedForHashing.Descendants().Where(d => d.Name == W.ins || d.Name == W.del).Attributes(W.id).Remove();
            var shaString = clonedForHashing.ToString(SaveOptions.DisableFormatting)
                .Replace(" xmlns=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", "");
            var sha1Hash = WmlComparerUtil.SHA1HashStringForUTF8String(shaString);
            consolidationInfo.RevisionHash = sha1Hash;

            var annotationList = elementToInsertAfter.Annotation<List<ConsolidationInfo>>();
            if (annotationList == null)
            {
                annotationList = new List<ConsolidationInfo>();
                elementToInsertAfter.AddAnnotation(annotationList);
            }
            annotationList.Add(consolidationInfo);
        }

        private static object CleanUpRevisionTransform(XNode node)
        {
            var element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.r &&
                    (element.Element(W.footnoteReference) != null || element.Element(W.endnoteReference) != null))
                    return null;

                return new XElement(element.Name,
                    element.Attributes().Where(a => a.Name.NamespaceName != PtOpenXml.pt),
                    element.Nodes().Select(n => CleanUpRevisionTransform(n)));
            }
            return node;
        }

        private static void AddTableGridStyleToStylesPart(StyleDefinitionsPart styleDefinitionsPart)
        {
            var sXDoc = styleDefinitionsPart.GetXDocument();
            var tableGridStyle = sXDoc
                .Root
                .Elements(W.style)
                .FirstOrDefault(s => (string)s.Attribute(W.styleId) == "TableGridForRevisions");
            if (tableGridStyle == null)
            {
                var tableGridForRevisionsStyleMarkup =
@"<w:style w:type=""table""
         w:styleId=""TableGridForRevisions""
         xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:name w:val=""Table Grid For Revisions""/>
  <w:basedOn w:val=""TableNormal""/>
  <w:rsid w:val=""0092121A""/>
  <w:rPr>
    <w:rFonts w:asciiTheme=""minorHAnsi""
              w:eastAsiaTheme=""minorEastAsia""
              w:hAnsiTheme=""minorHAnsi""
              w:cstheme=""minorBidi""/>
    <w:sz w:val=""22""/>
    <w:szCs w:val=""22""/>
  </w:rPr>
  <w:tblPr>
    <w:tblBorders>
      <w:top w:val=""single""
             w:sz=""4""
             w:space=""0""
             w:color=""auto""/>
      <w:left w:val=""single""
              w:sz=""4""
              w:space=""0""
              w:color=""auto""/>
      <w:bottom w:val=""single""
                w:sz=""4""
                w:space=""0""
                w:color=""auto""/>
      <w:right w:val=""single""
               w:sz=""4""
               w:space=""0""
               w:color=""auto""/>
      <w:insideH w:val=""single""
                 w:sz=""4""
                 w:space=""0""
                 w:color=""auto""/>
      <w:insideV w:val=""single""
                 w:sz=""4""
                 w:space=""0""
                 w:color=""auto""/>
    </w:tblBorders>
  </w:tblPr>
</w:style>";
                var tgsElement = XElement.Parse(tableGridForRevisionsStyleMarkup);
                sXDoc.Root.Add(tgsElement);
            }
            var tableNormalStyle = sXDoc
                .Root
                .Elements(W.style)
                .FirstOrDefault(s => (string)s.Attribute(W.styleId) == "TableNormal");
            if (tableNormalStyle == null)
            {
                var tableNormalStyleMarkup =
@"  <w:style w:type=""table""
           w:default=""1""
           w:styleId=""TableNormal"">
    <w:name w:val=""Normal Table""/>
    <w:uiPriority w:val=""99""/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:tblPr>
      <w:tblInd w:w=""0""
                w:type=""dxa""/>
      <w:tblCellMar>
        <w:top w:w=""0""
               w:type=""dxa""/>
        <w:left w:w=""108""
                w:type=""dxa""/>
        <w:bottom w:w=""0""
                  w:type=""dxa""/>
        <w:right w:w=""108""
                 w:type=""dxa""/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>";
                var tnsElement = XElement.Parse(tableNormalStyleMarkup);
                sXDoc.Root.Add(tnsElement);
            }
            styleDefinitionsPart.PutXDocument();
        }

        private static XAttribute[] NamespaceAttributes =
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
            new XAttribute(XNamespace.Xmlns + "wpg", WPG.wpg),
            new XAttribute(XNamespace.Xmlns + "wpi", WPI.wpi),
            new XAttribute(XNamespace.Xmlns + "wne", WNE.wne),
            new XAttribute(XNamespace.Xmlns + "wps", WPS.wps),
            new XAttribute(MC.Ignorable, "w14 wp14"),
        };

        private static void AddFootnotesEndnotesParts(WordprocessingDocument wDoc)
        {
            var mdp = wDoc.MainDocumentPart;
            if (mdp.FootnotesPart == null)
            {
                mdp.AddNewPart<FootnotesPart>();
                var newFootnotes = wDoc.MainDocumentPart.FootnotesPart.GetXDocument();
                newFootnotes.Declaration.Standalone = "yes";
                newFootnotes.Declaration.Encoding = "UTF-8";
                newFootnotes.Add(new XElement(W.footnotes, NamespaceAttributes));
                mdp.FootnotesPart.PutXDocument();
            }
            if (mdp.EndnotesPart == null)
            {
                mdp.AddNewPart<EndnotesPart>();
                var newEndnotes = wDoc.MainDocumentPart.EndnotesPart.GetXDocument();
                newEndnotes.Declaration.Standalone = "yes";
                newEndnotes.Declaration.Encoding = "UTF-8";
                newEndnotes.Add(new XElement(W.endnotes, NamespaceAttributes));
                mdp.EndnotesPart.PutXDocument();
            }
        }

        private static void ChangeFootnoteEndnoteReferencesToGuids(WordprocessingDocument wDoc)
        {
            var mainDocPart = wDoc.MainDocumentPart;
            var footnotesPart = wDoc.MainDocumentPart.FootnotesPart;
            var endnotesPart = wDoc.MainDocumentPart.EndnotesPart;

            var mainDocumentXDoc = mainDocPart.GetXDocument();
            XDocument footnotesPartXDoc = null;
            if (footnotesPart != null)
                footnotesPartXDoc = footnotesPart.GetXDocument();
            XDocument endnotesPartXDoc = null;
            if (endnotesPart != null)
                endnotesPartXDoc = endnotesPart.GetXDocument();

            var references = mainDocumentXDoc
                .Root
                .Descendants()
                .Where(d => d.Name == W.footnoteReference || d.Name == W.endnoteReference);

            foreach (var r in references)
            {
                var oldId = (string)r.Attribute(W.id);
                var newId = Guid.NewGuid().ToString().Replace("-", "");
                r.Attribute(W.id).Value = newId;
                if (r.Name == W.footnoteReference)
                {
                    var fn = footnotesPartXDoc
                        .Root
                        .Elements()
                        .FirstOrDefault(e => (string)e.Attribute(W.id) == oldId);
                    if (fn == null)
                        throw new OpenXmlPowerToolsException("Invalid document");
                    fn.Attribute(W.id).Value = newId;
                }
                else
                {
                    var en = endnotesPartXDoc
                        .Root
                        .Elements()
                        .FirstOrDefault(e => (string)e.Attribute(W.id) == oldId);
                    if (en == null)
                        throw new OpenXmlPowerToolsException("Invalid document");
                    en.Attribute(W.id).Value = newId;
                }
            }

            mainDocPart.PutXDocument();
            if (footnotesPart != null)
                footnotesPart.PutXDocument();
            if (endnotesPart != null)
                endnotesPart.PutXDocument();
        }

        private static WmlDocument ProduceDocumentWithTrackedRevisions(WmlComparerSettings settings, WmlDocument wmlResult, WordprocessingDocument wDoc1, WordprocessingDocument wDoc2)
        {
            var contentParent1 = wDoc1.MainDocumentPart.GetXDocument().Root.Element(W.body);
            AddSha1HashToBlockLevelContent(wDoc1.MainDocumentPart, contentParent1);
            var contentParent2 = wDoc2.MainDocumentPart.GetXDocument().Root.Element(W.body);
            AddSha1HashToBlockLevelContent(wDoc2.MainDocumentPart, contentParent2);

            var cal1 = WmlComparer.CreateComparisonUnitAtomList(wDoc1.MainDocumentPart, wDoc1.MainDocumentPart.GetXDocument().Root.Element(W.body));

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in cal1)
                    sb.Append(item.ToString() + Environment.NewLine);
                var sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            var cus1 = GetComparisonUnitList(cal1, settings);

            if (s_False)
            {
                var sbs = ComparisonUnit.ComparisonUnitListToString(cus1);
                TestUtil.NotePad(sbs);
            }

            var cal2 = WmlComparer.CreateComparisonUnitAtomList(wDoc2.MainDocumentPart, wDoc2.MainDocumentPart.GetXDocument().Root.Element(W.body));

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in cal2)
                    sb.Append(item.ToString() + Environment.NewLine);
                var sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            var cus2 = GetComparisonUnitList(cal2, settings);

            if (s_False)
            {
                var sbs = ComparisonUnit.ComparisonUnitListToString(cus2);
                TestUtil.NotePad(sbs);
            }

            if (s_False)
            {
                var sb3 = new StringBuilder();
                sb3.Append("ComparisonUnitList 1 =====" + Environment.NewLine + Environment.NewLine);
                sb3.Append(ComparisonUnit.ComparisonUnitListToString(cus1));
                sb3.Append(Environment.NewLine);
                sb3.Append("ComparisonUnitList 2 =====" + Environment.NewLine + Environment.NewLine);
                sb3.Append(ComparisonUnit.ComparisonUnitListToString(cus2));
                var sbs3 = sb3.ToString();
                TestUtil.NotePad(sbs3);
            }

            var correlatedSequence = Lcs(cus1, cus2, settings);

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in correlatedSequence)
                    sb.Append(item.ToString() + Environment.NewLine);
                var sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            // for any deleted or inserted rows, we go into the w:trPr properties, and add the appropriate w:ins or w:del element, and therefore
            // when generating the document, the appropriate row will be marked as deleted or inserted.
            MarkRowsAsDeletedOrInserted(settings, correlatedSequence);

            // the following gets a flattened list of ComparisonUnitAtoms, with status indicated in each ComparisonUnitAtom: Deleted, Inserted, or Equal
            var listOfComparisonUnitAtoms = FlattenToComparisonUnitAtomList(correlatedSequence);

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in listOfComparisonUnitAtoms)
                    sb.Append(item.ToString() + Environment.NewLine);
                var sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            // hack = set the guid ID of the table, row, or cell from the 'before' document to be equal to the 'after' document.

            // note - we don't want to do the hack until after flattening all of the groups.  At the end of the flattening, we should simply
            // have a list of ComparisonUnitAtoms, appropriately marked as equal, inserted, or deleted.

            // the table id will be hacked in the normal course of events.
            // in the case where a row is deleted, not necessary to hack - the deleted row ID will do.
            // in the case where a row is inserted, not necessary to hack - the inserted row ID will do as well.
            SetUnidForPreviousDocument(correlatedSequence);

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in listOfComparisonUnitAtoms)
                    sb.Append(item.ToString() + Environment.NewLine);
                var sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            // and then finally can generate the document with revisions
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(wmlResult.DocumentByteArray, 0, wmlResult.DocumentByteArray.Length);
                using (WordprocessingDocument wDocWithRevisions = WordprocessingDocument.Open(ms, true))
                {
                    var xDoc = wDocWithRevisions.MainDocumentPart.GetXDocument();
                    var rootNamespaceAttributes = xDoc
                        .Root
                        .Attributes()
                        .Where(a => a.IsNamespaceDeclaration || a.Name.Namespace == MC.mc)
                        .ToList();

                    // ======================================
                    // The following produces a new valid WordprocessingML document from the listOfComparisonUnitAtoms
                    var newBodyChildren = ProduceNewWmlMarkupFromCorrelatedSequence(wDocWithRevisions.MainDocumentPart,
                        listOfComparisonUnitAtoms, settings);

                    XDocument newXDoc = new XDocument();
                    newXDoc.Add(
                        new XElement(W.document,
                            rootNamespaceAttributes,
                            new XElement(W.body, newBodyChildren)));
                    MarkContentAsDeletedOrInserted(newXDoc, settings);
                    CoalesceAdjacentRunsWithIdenticalFormatting(newXDoc);
                    IgnorePt14Namespace(newXDoc.Root);

                    ProcessFootnoteEndnote(settings,
                        listOfComparisonUnitAtoms,
                        wDoc1.MainDocumentPart,
                        wDoc2.MainDocumentPart,
                        newXDoc);

                    RectifyFootnoteEndnoteIds(
                        wDoc1.MainDocumentPart,
                        wDoc2.MainDocumentPart,
                        wDocWithRevisions.MainDocumentPart,
                        newXDoc,
                        settings);

                    ConjoinDeletedInsertedParagraphMarks(wDocWithRevisions.MainDocumentPart, newXDoc);

                    FixUpRevisionIds(wDocWithRevisions, newXDoc);

                    // little bit of cleanup
                    MoveLastSectPrToChildOfBody(newXDoc);
                    XElement newXDoc2Root = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(newXDoc.Root);
                    xDoc.Root.ReplaceWith(newXDoc2Root);

                    /**********************************************************************************************/
                    // temporary code to remove sections.  When remove this code, get validation errors for some ITU documents.

                    xDoc.Root.Descendants(W.sectPr).Remove();

                    /**********************************************************************************************/

                    wDocWithRevisions.MainDocumentPart.PutXDocument();

                    FixUpRevMarkIds(wDocWithRevisions);
                    FixUpDocPrIds(wDocWithRevisions);
                    FixUpShapeIds(wDocWithRevisions);
                    FixUpShapeTypeIds(wDocWithRevisions);
                }
                foreach (var part in wDoc1.ContentParts())
                    part.PutXDocument();
                var updatedWmlResult = new WmlDocument("Dummy.docx", ms.ToArray());
                return updatedWmlResult;
            }
        }

        // it is possible, per the algorithm, for the algorithm to find that the paragraph mark for a single paragraph has been
        // inserted and deleted.  If the algorithm sets them to equal, then sometimes it will equate paragraph marks that should
        // not be equated.  
        private static void ConjoinDeletedInsertedParagraphMarks(MainDocumentPart mainDocumentPart, XDocument newXDoc)
        {
            ConjoinMultipleParagraphMarks(newXDoc);
            if (mainDocumentPart.FootnotesPart != null)
            {
                var fnXDoc = mainDocumentPart.FootnotesPart.GetXDocument();
                ConjoinMultipleParagraphMarks(fnXDoc);
                mainDocumentPart.FootnotesPart.PutXDocument();
            }
            if (mainDocumentPart.EndnotesPart != null)
            {
                var fnXDoc = mainDocumentPart.EndnotesPart.GetXDocument();
                ConjoinMultipleParagraphMarks(fnXDoc);
                mainDocumentPart.EndnotesPart.PutXDocument();
            }
        }

        private static void ConjoinMultipleParagraphMarks(XDocument xDoc)
        {
            var newRoot = ConjoinTransform(xDoc.Root);
            xDoc.Root.ReplaceWith(newRoot);
        }

        private static object ConjoinTransform(XNode node)
        {
            var element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.p && element.Elements(W.pPr).Count() >= 2)
                {
                    var pPr = new XElement(element.Element(W.pPr));
                    pPr.Elements(W.rPr).Elements().Where(r => r.Name == W.ins || r.Name == W.del).Remove();
                    pPr.Attributes(PtOpenXml.Status).Remove();
                    var newPara = new XElement(W.p,
                        element.Attributes(),
                        pPr,
                        element.Elements().Where(c => c.Name != W.pPr));
                    return newPara;
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => ConjoinTransform(n)));
            }
            return node;
        }

        private static void MarkContentAsDeletedOrInserted(XDocument newXDoc, WmlComparerSettings settings)
        {
            var newRoot = MarkContentAsDeletedOrInsertedTransform(newXDoc.Root, settings);
            newXDoc.Root.ReplaceWith(newRoot);
        }

        private static object MarkContentAsDeletedOrInsertedTransform(XNode node, WmlComparerSettings settings)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.r)
                {
                    var statusList = element
                        .DescendantsTrimmed(W.txbxContent)
                        .Where(d => d.Name == W.t || d.Name == W.delText || AllowableRunChildren.Contains(d.Name))
                        .Attributes(PtOpenXml.Status)
                        .Select(a => (string)a)
                        .Distinct()
                        .ToList();
                    if (statusList.Count() > 1)
                        throw new OpenXmlPowerToolsException("Internal error - have both deleted and inserted text elements in the same run.");
                    if (statusList.Count() == 0)
                        return new XElement(W.r,
                            element.Attributes(),
                            element.Nodes().Select(n => MarkContentAsDeletedOrInsertedTransform(n, settings)));
                    if (statusList.First() == "Deleted")
                    {
                        return new XElement(W.del,
                            new XAttribute(W.author, settings.AuthorForRevisions),
                            new XAttribute(W.id, s_MaxId++),
                            new XAttribute(W.date, settings.DateTimeForRevisions),
                            new XElement(W.r,
                            element.Attributes(),
                            element.Nodes().Select(n => MarkContentAsDeletedOrInsertedTransform(n, settings))));
                    }
                    else if (statusList.First() == "Inserted")
                    {
                        return new XElement(W.ins,
                            new XAttribute(W.author, settings.AuthorForRevisions),
                            new XAttribute(W.id, s_MaxId++),
                            new XAttribute(W.date, settings.DateTimeForRevisions),
                            new XElement(W.r,
                            element.Attributes(),
                            element.Nodes().Select(n => MarkContentAsDeletedOrInsertedTransform(n, settings))));
                    }
                }

                if (element.Name == W.pPr)
                {
                    var status = (string)element.Attribute(PtOpenXml.Status);
                    if (status == null)
                        return new XElement(W.pPr,
                            element.Attributes(),
                            element.Nodes().Select(n => MarkContentAsDeletedOrInsertedTransform(n, settings)));
                    var pPr = new XElement(element);
                    if (status == "Deleted")
                    {
                        XElement rPr = pPr.Element(W.rPr);
                        if (rPr == null)
                            rPr = new XElement(W.rPr);
                        rPr.Add(new XElement(W.del,
                            new XAttribute(W.author, settings.AuthorForRevisions),
                            new XAttribute(W.id, s_MaxId++),
                            new XAttribute(W.date, settings.DateTimeForRevisions)));
                        if (pPr.Element(W.rPr) != null)
                            pPr.Element(W.rPr).ReplaceWith(rPr);
                        else
                            pPr.AddFirst(rPr);
                    }
                    else if (status == "Inserted")
                    {
                        XElement rPr = pPr.Element(W.rPr);
                        if (rPr == null)
                            rPr = new XElement(W.rPr);
                        rPr.Add(new XElement(W.ins,
                            new XAttribute(W.author, settings.AuthorForRevisions),
                            new XAttribute(W.id, s_MaxId++),
                            new XAttribute(W.date, settings.DateTimeForRevisions)));
                        if (pPr.Element(W.rPr) != null)
                            pPr.Element(W.rPr).ReplaceWith(rPr);
                        else
                            pPr.AddFirst(rPr);
                    }
                    else
                        throw new OpenXmlPowerToolsException("Internal error");
                    return pPr;
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => MarkContentAsDeletedOrInsertedTransform(n, settings)));
            }
            return node;
        }

        private static void FixUpRevisionIds(WordprocessingDocument wDocWithRevisions, XDocument newXDoc)
        {
            IEnumerable<XElement> footnoteRevisions = Enumerable.Empty<XElement>();
            if (wDocWithRevisions.MainDocumentPart.FootnotesPart != null)
            {
                var fnxd = wDocWithRevisions.MainDocumentPart.FootnotesPart.GetXDocument();
                footnoteRevisions = fnxd
                    .Descendants()
                    .Where(d => d.Name == W.ins || d.Name == W.del);
            }
            IEnumerable<XElement> endnoteRevisions = Enumerable.Empty<XElement>();
            if (wDocWithRevisions.MainDocumentPart.EndnotesPart != null)
            {
                var fnxd = wDocWithRevisions.MainDocumentPart.EndnotesPart.GetXDocument();
                endnoteRevisions = fnxd
                    .Descendants()
                    .Where(d => d.Name == W.ins || d.Name == W.del);
            }
            var mainRevisions = newXDoc
                .Descendants()
                .Where(d => d.Name == W.ins || d.Name == W.del);
            var allRevisions = mainRevisions
                .Concat(footnoteRevisions)
                .Concat(endnoteRevisions)
                .Select((r, i) =>
                {
                    return new
                    {
                        Rev = r,
                        Idx = i + 1,
                    };
                });
            foreach (var item in allRevisions)
                item.Rev.Attribute(W.id).Value = item.Idx.ToString();
            if (wDocWithRevisions.MainDocumentPart.FootnotesPart != null)
                wDocWithRevisions.MainDocumentPart.FootnotesPart.PutXDocument();
            if (wDocWithRevisions.MainDocumentPart.EndnotesPart != null)
                wDocWithRevisions.MainDocumentPart.EndnotesPart.PutXDocument();
        }

        private static void IgnorePt14Namespace(XElement root)
        {
            if (root.Attribute(XNamespace.Xmlns + "pt14") == null)
            {
                root.Add(new XAttribute(XNamespace.Xmlns + "pt14", PtOpenXml.pt.NamespaceName));
            }
            var ignorable = (string)root.Attribute(MC.Ignorable);
            if (ignorable != null)
            {
                var list = ignorable.Split(' ');
                if (!list.Contains("pt14"))
                {
                    ignorable += " pt14";
                    root.Attribute(MC.Ignorable).Value = ignorable;
                }
            }
            else
            {
                root.Add(new XAttribute(MC.Ignorable, "pt14"));
            }
        }

        private static void CoalesceAdjacentRunsWithIdenticalFormatting(XDocument xDoc)
        {
            var paras = xDoc.Root.DescendantsTrimmed(W.txbxContent).Where(d => d.Name == W.p);
            foreach (var para in paras)
            {
                var newPara = WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(para);
                para.ReplaceNodes(newPara.Nodes());
            }
        }

        private static void ProcessFootnoteEndnote(
            WmlComparerSettings settings,
            List<ComparisonUnitAtom> listOfComparisonUnitAtoms,
            MainDocumentPart mainDocumentPartBefore,
            MainDocumentPart mainDocumentPartAfter,
            XDocument mainDocumentXDoc)
        {
            var footnotesPartBefore = mainDocumentPartBefore.FootnotesPart;
            var endnotesPartBefore = mainDocumentPartBefore.EndnotesPart;
            var footnotesPartAfter = mainDocumentPartAfter.FootnotesPart;
            var endnotesPartAfter = mainDocumentPartAfter.EndnotesPart;

            XDocument footnotesPartBeforeXDoc = null;
            if (footnotesPartBefore != null)
                footnotesPartBeforeXDoc = footnotesPartBefore.GetXDocument();
            XDocument footnotesPartAfterXDoc = null;
            if (footnotesPartAfter != null)
                footnotesPartAfterXDoc = footnotesPartAfter.GetXDocument();
            XDocument endnotesPartBeforeXDoc = null;
            if (endnotesPartBefore != null)
                endnotesPartBeforeXDoc = endnotesPartBefore.GetXDocument();
            XDocument endnotesPartAfterXDoc = null;
            if (endnotesPartAfter != null)
                endnotesPartAfterXDoc = endnotesPartAfter.GetXDocument();

            var possiblyModifiedFootnotesEndNotes = listOfComparisonUnitAtoms
                .Where(cua =>
                    cua.ContentElement.Name == W.footnoteReference ||
                    cua.ContentElement.Name == W.endnoteReference)
                .ToList();

            foreach (var fn in possiblyModifiedFootnotesEndNotes)
            {
                string beforeId = null;
                if (fn.ContentElementBefore != null)
                    beforeId = (string)fn.ContentElementBefore.Attribute(W.id);
                var afterId = (string)fn.ContentElement.Attribute(W.id);

                XElement footnoteEndnoteBefore = null;
                XElement footnoteEndnoteAfter = null;
                OpenXmlPart partToUseBefore = null;
                OpenXmlPart partToUseAfter = null;
                XDocument partToUseBeforeXDoc = null;
                XDocument partToUseAfterXDoc = null;

                if (fn.CorrelationStatus == CorrelationStatus.Equal)
                {
                    if (fn.ContentElement.Name == W.footnoteReference)
                    {
                        footnoteEndnoteBefore = footnotesPartBeforeXDoc
                            .Root
                            .Elements()
                            .FirstOrDefault(fnn => (string)fnn.Attribute(W.id) == beforeId);
                        footnoteEndnoteAfter = footnotesPartAfterXDoc
                            .Root
                            .Elements()
                            .FirstOrDefault(fnn => (string)fnn.Attribute(W.id) == afterId);
                        partToUseBefore = footnotesPartBefore;
                        partToUseAfter = footnotesPartAfter;
                        partToUseBeforeXDoc = footnotesPartBeforeXDoc;
                        partToUseAfterXDoc = footnotesPartAfterXDoc;
                    }
                    else
                    {
                        footnoteEndnoteBefore = endnotesPartBeforeXDoc
                            .Root
                            .Elements()
                            .FirstOrDefault(fnn => (string)fnn.Attribute(W.id) == beforeId);
                        footnoteEndnoteAfter = endnotesPartAfterXDoc
                            .Root
                            .Elements()
                            .FirstOrDefault(fnn => (string)fnn.Attribute(W.id) == afterId);
                        partToUseBefore = endnotesPartBefore;
                        partToUseAfter = endnotesPartAfter;
                        partToUseBeforeXDoc = endnotesPartBeforeXDoc;
                        partToUseAfterXDoc = endnotesPartAfterXDoc;
                    }
                    AddSha1HashToBlockLevelContent(partToUseBefore, footnoteEndnoteBefore);
                    AddSha1HashToBlockLevelContent(partToUseAfter, footnoteEndnoteAfter);

                    var fncal1 = WmlComparer.CreateComparisonUnitAtomList(partToUseBefore, footnoteEndnoteBefore);
                    var fncus1 = GetComparisonUnitList(fncal1, settings);

                    var fncal2 = WmlComparer.CreateComparisonUnitAtomList(partToUseAfter, footnoteEndnoteAfter);
                    var fncus2 = GetComparisonUnitList(fncal2, settings);

                    var fnCorrelatedSequence = Lcs(fncus1, fncus2, settings);

                    if (s_False)
                    {
                        var sb = new StringBuilder();
                        foreach (var item in fnCorrelatedSequence)
                            sb.Append(item.ToString()).Append(Environment.NewLine);
                        var sbs = sb.ToString();
                        TestUtil.NotePad(sbs);
                    }

                    // for any deleted or inserted rows, we go into the w:trPr properties, and add the appropriate w:ins or w:del element, and therefore
                    // when generating the document, the appropriate row will be marked as deleted or inserted.
                    MarkRowsAsDeletedOrInserted(settings, fnCorrelatedSequence);

                    // the following gets a flattened list of ComparisonUnitAtoms, with status indicated in each ComparisonUnitAtom: Deleted, Inserted, or Equal
                    var fnListOfComparisonUnitAtoms = FlattenToComparisonUnitAtomList(fnCorrelatedSequence);

                    if (s_False)
                    {
                        var sb = new StringBuilder();
                        foreach (var item in fnListOfComparisonUnitAtoms)
                            sb.Append(item.ToString() + Environment.NewLine);
                        var sbs = sb.ToString();
                        TestUtil.NotePad(sbs);
                    }

                    // hack = set the guid ID of the table, row, or cell from the 'before' document to be equal to the 'after' document.

                    // note - we don't want to do the hack until after flattening all of the groups.  At the end of the flattening, we should simply
                    // have a list of ComparisonUnitAtoms, appropriately marked as equal, inserted, or deleted.

                    // the table id will be hacked in the normal course of events.
                    // in the case where a row is deleted, not necessary to hack - the deleted row ID will do.
                    // in the case where a row is inserted, not necessary to hack - the inserted row ID will do as well.
                    SetUnidForPreviousDocument(fnCorrelatedSequence);

                    var newFootnoteEndnoteChildren = ProduceNewWmlMarkupFromCorrelatedSequence(partToUseAfter, fnListOfComparisonUnitAtoms, settings);
                    var tempElement = new XElement(W.body, newFootnoteEndnoteChildren);
                    var firstPara = tempElement.Descendants(W.p).FirstOrDefault();
                    if (firstPara != null)
                    {
                        var firstRun = firstPara.Element(W.r);
                        if (firstRun != null)
                        {
                            if (fn.ContentElement.Name == W.footnoteReference)
                                firstRun.AddBeforeSelf(
                                    new XElement(W.r,
                                        new XElement(W.rPr,
                                            new XElement(W.rStyle,
                                                new XAttribute(W.val, "FootnoteReference"))),
                                        new XElement(W.footnoteRef)));
                            else
                                firstRun.AddBeforeSelf(
                                    new XElement(W.r,
                                        new XElement(W.rPr,
                                            new XElement(W.rStyle,
                                                new XAttribute(W.val, "EndnoteReference"))),
                                        new XElement(W.endnoteRef)));
                        }
                    }
                    XElement newTempElement = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(tempElement);
                    var newContentElement = newTempElement.Descendants().FirstOrDefault(d => d.Name == W.footnote || d.Name == W.endnote);
                    if (newContentElement == null)
                        throw new OpenXmlPowerToolsException("Internal error");
                    footnoteEndnoteAfter.ReplaceNodes(newContentElement.Nodes());
                }
                else if (fn.CorrelationStatus == CorrelationStatus.Inserted)
                {
                    if (fn.ContentElement.Name == W.footnoteReference)
                    {
                        footnoteEndnoteAfter = footnotesPartAfterXDoc
                            .Root
                            .Elements()
                            .FirstOrDefault(fnn => (string)fnn.Attribute(W.id) == afterId);
                        partToUseAfter = footnotesPartAfter;
                        partToUseAfterXDoc = footnotesPartAfterXDoc;
                    }
                    else
                    {
                        footnoteEndnoteAfter = endnotesPartAfterXDoc
                            .Root
                            .Elements()
                            .FirstOrDefault(fnn => (string)fnn.Attribute(W.id) == afterId);
                        partToUseAfter = endnotesPartAfter;
                        partToUseAfterXDoc = endnotesPartAfterXDoc;
                    }

                    AddSha1HashToBlockLevelContent(partToUseAfter, footnoteEndnoteAfter);

                    var fncal2 = WmlComparer.CreateComparisonUnitAtomList(partToUseAfter, footnoteEndnoteAfter);
                    var fncus2 = GetComparisonUnitList(fncal2, settings);

                    var insertedCorrSequ = new List<CorrelatedSequence>() {
                        new CorrelatedSequence()
                        {
                            ComparisonUnitArray1 = null,
                            ComparisonUnitArray2 = fncus2,
                            CorrelationStatus = CorrelationStatus.Inserted,
                        },
                    };

                    if (s_False)
                    {
                        var sb = new StringBuilder();
                        foreach (var item in insertedCorrSequ)
                            sb.Append(item.ToString()).Append(Environment.NewLine);
                        var sbs = sb.ToString();
                        TestUtil.NotePad(sbs);
                    }

                    MarkRowsAsDeletedOrInserted(settings, insertedCorrSequ);

                    var fnListOfComparisonUnitAtoms = FlattenToComparisonUnitAtomList(insertedCorrSequ);

                    var newFootnoteEndnoteChildren = ProduceNewWmlMarkupFromCorrelatedSequence(partToUseAfter,
                        fnListOfComparisonUnitAtoms, settings);
                    var tempElement = new XElement(W.body, newFootnoteEndnoteChildren);
                    var firstPara = tempElement.Descendants(W.p).FirstOrDefault();
                    if (firstPara != null)
                    {
                        var firstRun = firstPara.Descendants(W.r).FirstOrDefault();
                        if (firstRun != null)
                        {
                            if (fn.ContentElement.Name == W.footnoteReference)
                                firstRun.AddBeforeSelf(
                                    new XElement(W.r,
                                        new XElement(W.rPr,
                                            new XElement(W.rStyle,
                                                new XAttribute(W.val, "FootnoteReference"))),
                                        new XElement(W.footnoteRef)));
                            else
                                firstRun.AddBeforeSelf(
                                    new XElement(W.r,
                                        new XElement(W.rPr,
                                            new XElement(W.rStyle,
                                                new XAttribute(W.val, "EndnoteReference"))),
                                        new XElement(W.endnoteRef)));
                        }
                    }
                    XElement newTempElement = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(tempElement);
                    var newContentElement = newTempElement
                        .Descendants()
                        .FirstOrDefault(d => d.Name == W.footnote || d.Name == W.endnote);
                    if (newContentElement == null)
                        throw new OpenXmlPowerToolsException("Internal error");
                    footnoteEndnoteAfter.ReplaceNodes(newContentElement.Nodes());
                }
                else if (fn.CorrelationStatus == CorrelationStatus.Deleted)
                {
                    if (fn.ContentElement.Name == W.footnoteReference)
                    {
                        footnoteEndnoteBefore = footnotesPartBeforeXDoc
                            .Root
                            .Elements()
                            .FirstOrDefault(fnn => (string)fnn.Attribute(W.id) == afterId);
                        partToUseAfter = footnotesPartAfter;
                        partToUseAfterXDoc = footnotesPartAfterXDoc;
                    }
                    else
                    {
                        footnoteEndnoteBefore = endnotesPartBeforeXDoc
                            .Root
                            .Elements()
                            .FirstOrDefault(fnn => (string)fnn.Attribute(W.id) == afterId);
                        partToUseBefore = endnotesPartBefore;
                        partToUseBeforeXDoc = endnotesPartBeforeXDoc;
                    }

                    AddSha1HashToBlockLevelContent(partToUseBefore, footnoteEndnoteBefore);

                    var fncal2 = WmlComparer.CreateComparisonUnitAtomList(partToUseBefore, footnoteEndnoteBefore);
                    var fncus2 = GetComparisonUnitList(fncal2, settings);

                    var deletedCorrSequ = new List<CorrelatedSequence>() {
                        new CorrelatedSequence()
                        {
                            ComparisonUnitArray1 = fncus2,
                            ComparisonUnitArray2 = null,
                            CorrelationStatus = CorrelationStatus.Deleted,
                        },
                    };

                    if (s_False)
                    {
                        var sb = new StringBuilder();
                        foreach (var item in deletedCorrSequ)
                            sb.Append(item.ToString()).Append(Environment.NewLine);
                        var sbs = sb.ToString();
                        TestUtil.NotePad(sbs);
                    }

                    MarkRowsAsDeletedOrInserted(settings, deletedCorrSequ);

                    var fnListOfComparisonUnitAtoms = FlattenToComparisonUnitAtomList(deletedCorrSequ);

                    var newFootnoteEndnoteChildren = ProduceNewWmlMarkupFromCorrelatedSequence(partToUseBefore,
                        fnListOfComparisonUnitAtoms, settings);
                    var tempElement = new XElement(W.body, newFootnoteEndnoteChildren);
                    var firstPara = tempElement.Descendants(W.p).FirstOrDefault();
                    if (firstPara != null)
                    {
                        var firstRun = firstPara.Descendants(W.r).FirstOrDefault();
                        if (firstRun != null)
                        {
                            if (fn.ContentElement.Name == W.footnoteReference)
                                firstRun.AddBeforeSelf(
                                    new XElement(W.r,
                                        new XElement(W.rPr,
                                            new XElement(W.rStyle,
                                                new XAttribute(W.val, "FootnoteReference"))),
                                        new XElement(W.footnoteRef)));
                            else
                                firstRun.AddBeforeSelf(
                                    new XElement(W.r,
                                        new XElement(W.rPr,
                                            new XElement(W.rStyle,
                                                new XAttribute(W.val, "EndnoteReference"))),
                                        new XElement(W.endnoteRef)));
                        }
                    }
                    XElement newTempElement = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(tempElement);
                    var newContentElement = newTempElement.Descendants().FirstOrDefault(d => d.Name == W.footnote || d.Name == W.endnote);
                    if (newContentElement == null)
                        throw new OpenXmlPowerToolsException("Internal error");
                    footnoteEndnoteBefore.ReplaceNodes(newContentElement.Nodes());
                }
                else
                    throw new OpenXmlPowerToolsException("Internal error");
            }
        }

        private static void RectifyFootnoteEndnoteIds(
            MainDocumentPart mainDocumentPartBefore,
            MainDocumentPart mainDocumentPartAfter,
            MainDocumentPart mainDocumentPartWithRevisions,
            XDocument mainDocumentXDoc,
            WmlComparerSettings settings)
        {
            var footnotesPartBefore = mainDocumentPartBefore.FootnotesPart;
            var endnotesPartBefore = mainDocumentPartBefore.EndnotesPart;
            var footnotesPartAfter = mainDocumentPartAfter.FootnotesPart;
            var endnotesPartAfter = mainDocumentPartAfter.EndnotesPart;
            var footnotesPartWithRevisions = mainDocumentPartWithRevisions.FootnotesPart;
            var endnotesPartWithRevisions = mainDocumentPartWithRevisions.EndnotesPart;

            XDocument footnotesPartBeforeXDoc = null;
            if (footnotesPartBefore != null)
                footnotesPartBeforeXDoc = footnotesPartBefore.GetXDocument();

            XDocument footnotesPartAfterXDoc = null;
            if (footnotesPartAfter != null)
                footnotesPartAfterXDoc = footnotesPartAfter.GetXDocument();

            XDocument footnotesPartWithRevisionsXDoc = null;
            if (footnotesPartWithRevisions != null)
            {
                footnotesPartWithRevisionsXDoc = footnotesPartWithRevisions.GetXDocument();
                footnotesPartWithRevisionsXDoc
                    .Root
                    .Elements(W.footnote)
                    .Where(e => (string)e.Attribute(W.id) != "-1" && (string)e.Attribute(W.id) != "0")
                    .Remove();
            }

            XDocument endnotesPartBeforeXDoc = null;
            if (endnotesPartBefore != null)
                endnotesPartBeforeXDoc = endnotesPartBefore.GetXDocument();

            XDocument endnotesPartAfterXDoc = null;
            if (endnotesPartAfter != null)
                endnotesPartAfterXDoc = endnotesPartAfter.GetXDocument();

            XDocument endnotesPartWithRevisionsXDoc = null;
            if (endnotesPartWithRevisions != null)
            {
                endnotesPartWithRevisionsXDoc = endnotesPartWithRevisions.GetXDocument();
                endnotesPartWithRevisionsXDoc
                    .Root
                    .Elements(W.endnote)
                    .Where(e => (string)e.Attribute(W.id) != "-1" && (string)e.Attribute(W.id) != "0")
                    .Remove();
            }

            var footnotesRefs = mainDocumentXDoc
                .Descendants(W.footnoteReference)
                .Select((fn, idx) =>
                {
                    return new
                    {
                        FootNote = fn,
                        Idx = idx,
                    };
                });

            foreach (var fn in footnotesRefs)
            {
                var oldId = (string)fn.FootNote.Attribute(W.id);
                var newId = (fn.Idx + 1).ToString();
                fn.FootNote.Attribute(W.id).Value = newId;
                var footnote = footnotesPartAfterXDoc
                    .Root
                    .Elements()
                    .FirstOrDefault(e => (string)e.Attribute(W.id) == oldId);
                if (footnote == null)
                {
                    footnote = footnotesPartBeforeXDoc
                        .Root
                        .Elements()
                        .FirstOrDefault(e => (string)e.Attribute(W.id) == oldId);
                }
                if (footnote == null)
                    throw new OpenXmlPowerToolsException("Internal error");
                footnote.Attribute(W.id).Value = newId;
                footnotesPartWithRevisionsXDoc
                    .Root
                    .Add(footnote);
            }

            var endnotesRefs = mainDocumentXDoc
                .Descendants(W.endnoteReference)
                .Select((fn, idx) =>
                {
                    return new
                    {
                        Endnote = fn,
                        Idx = idx,
                    };
                });

            foreach (var fn in endnotesRefs)
            {
                var oldId = (string)fn.Endnote.Attribute(W.id);
                var newId = (fn.Idx + 1).ToString();
                fn.Endnote.Attribute(W.id).Value = newId;
                var endnote = endnotesPartAfterXDoc
                    .Root
                    .Elements()
                    .FirstOrDefault(e => (string)e.Attribute(W.id) == oldId);
                if (endnote == null)
                {
                    endnote = endnotesPartBeforeXDoc
                        .Root
                        .Elements()
                        .FirstOrDefault(e => (string)e.Attribute(W.id) == oldId);
                }
                if (endnote == null)
                    throw new OpenXmlPowerToolsException("Internal error");
                endnote.Attribute(W.id).Value = newId;
                endnotesPartWithRevisionsXDoc
                    .Root
                    .Add(endnote);
            }

            if (footnotesPartWithRevisionsXDoc != null)
            {
                MarkContentAsDeletedOrInserted(footnotesPartWithRevisionsXDoc, settings);
                CoalesceAdjacentRunsWithIdenticalFormatting(footnotesPartWithRevisionsXDoc);
                XElement newXDocRoot = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(footnotesPartWithRevisionsXDoc.Root);
                footnotesPartWithRevisionsXDoc.Root.ReplaceWith(newXDocRoot);
                IgnorePt14Namespace(footnotesPartWithRevisionsXDoc.Root);
                footnotesPartWithRevisions.PutXDocument();
            }
            if (endnotesPartWithRevisionsXDoc != null)
            {
                MarkContentAsDeletedOrInserted(endnotesPartWithRevisionsXDoc, settings);
                CoalesceAdjacentRunsWithIdenticalFormatting(endnotesPartWithRevisionsXDoc);
                XElement newXDocRoot = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(endnotesPartWithRevisionsXDoc.Root);
                endnotesPartWithRevisionsXDoc.Root.ReplaceWith(newXDocRoot);
                IgnorePt14Namespace(endnotesPartWithRevisionsXDoc.Root);
                endnotesPartWithRevisions.PutXDocument();
            }
        }

        // hack = set the guid ID of the table, row, or cell from the 'before' document to be equal to the 'after' document.

        // note - we don't want to do the hack until after flattening all of the groups.  At the end of the flattening, we should simply
        // have a list of ComparisonUnitAtoms, appropriately marked as equal, inserted, or deleted.

        // the table id will be hacked in the normal course of events.
        // in the case where a row is deleted, not necessary to hack - the deleted row ID will do.
        // in the case where a row is inserted, not necessary to hack - the inserted row ID will do as well.
        private static void SetUnidForPreviousDocument(List<CorrelatedSequence> correlatedSequence)
        {
            HashSet<string> alreadySetUnids = new HashSet<string>();
            HashSet<string> alreadySetPrevUnids = new HashSet<string>();
            foreach (var cs in correlatedSequence.Where(z => z.CorrelationStatus == CorrelationStatus.Equal))
            {
                var zippedComparisonUnitArrays = cs.ComparisonUnitArray1.Zip(cs.ComparisonUnitArray2, (cuBefore, cuAfter) => new
                {
                    CuBefore = cuBefore,
                    CuAfter = cuAfter,
                });
                foreach (var cu in zippedComparisonUnitArrays)
                {
                    var beforeDescendantContentAtoms = cu.CuBefore
                        .DescendantContentAtoms();

                    var afterDescendantContentAtoms = cu.CuAfter
                        .DescendantContentAtoms();

                    var zippedContents = beforeDescendantContentAtoms
                        .Zip(afterDescendantContentAtoms,
                            (conBefore, conAfter) => new
                            {
                                ConBefore = conBefore,
                                ConAfter = conAfter,
                            });

                    foreach (var con in zippedContents)
                    {
                        var zippedAncestors = con.ConBefore.AncestorElements.Zip(con.ConAfter.AncestorElements, (ancBefore, ancAfter) => new
                        {
                            AncestorBefore = ancBefore,
                            AncestorAfter = ancAfter,
                        });
                        foreach (var anc in zippedAncestors)
                        {
                            if (anc.AncestorBefore.Attribute(PtOpenXml.Unid) == null ||
                                anc.AncestorAfter.Attribute(PtOpenXml.Unid) == null ||
                                anc.AncestorBefore.Attribute(PtOpenXml.PrevUnid) == null ||
                                anc.AncestorAfter.Attribute(PtOpenXml.PrevUnid) == null)
                                continue;
                            var beforeUnid = (string)anc.AncestorBefore.Attribute(PtOpenXml.Unid);
                            var afterUnid = (string)anc.AncestorAfter.Attribute(PtOpenXml.Unid);
                            var beforePrevUnid = (string)anc.AncestorBefore.Attribute(PtOpenXml.PrevUnid);
                            var afterPrevUnid = (string)anc.AncestorAfter.Attribute(PtOpenXml.PrevUnid);
                            if (beforeUnid != afterUnid)
                            {
                                if (!alreadySetUnids.Contains(beforeUnid))
                                {
                                    alreadySetUnids.Add(beforeUnid);
                                    anc.AncestorBefore.Attribute(PtOpenXml.Unid).Value = afterUnid;
                                }
                            }
                            if (beforePrevUnid != afterPrevUnid)
                            {
                                if (!alreadySetPrevUnids.Contains(beforePrevUnid))
                                {
                                    alreadySetPrevUnids.Add(beforePrevUnid);
                                    anc.AncestorAfter.Attribute(PtOpenXml.PrevUnid).Value = beforePrevUnid;
                                }
                            }
                        }
                    }
                }
            }
        }

        // the following gets a flattened list of ComparisonUnitAtoms, with status indicated in each ComparisonUnitAtom: Deleted, Inserted, or Equal
        private static List<ComparisonUnitAtom> FlattenToComparisonUnitAtomList(List<CorrelatedSequence> correlatedSequence)
        {
            var listOfComparisonUnitAtoms = correlatedSequence
                .Select(cs =>
                {
                    if (cs.CorrelationStatus == CorrelationStatus.Equal)
                    {
                        var contentAtomsBefore = cs
                            .ComparisonUnitArray1
                            .Select(ca => ca.DescendantContentAtoms())
                            .SelectMany(m => m);

                        var contentAtomsAfter = cs
                            .ComparisonUnitArray2
                            .Select(ca => ca.DescendantContentAtoms())
                            .SelectMany(m => m);

                        var comparisonUnitAtomList = contentAtomsBefore
                            .Zip(contentAtomsAfter,
                                (before, after) =>
                                {
                                    return new ComparisonUnitAtom(after.ContentElement, after.AncestorElements, after.Part)
                                    {
                                        CorrelationStatus = CorrelationStatus.Equal,
                                        ContentElementBefore = before.ContentElement,
                                    };
                                })
                            .ToList();
                        return comparisonUnitAtomList;
                    }
                    else if (cs.CorrelationStatus == CorrelationStatus.Deleted)
                    {
                        var comparisonUnitAtomList = cs
                            .ComparisonUnitArray1
                            .Select(ca => ca.DescendantContentAtoms())
                            .SelectMany(m => m)
                            .Select(ca =>
                                new ComparisonUnitAtom(ca.ContentElement, ca.AncestorElements, ca.Part)
                                {
                                    CorrelationStatus = CorrelationStatus.Deleted,
                                });
                        return comparisonUnitAtomList;
                    }
                    else if (cs.CorrelationStatus == CorrelationStatus.Inserted)
                    {
                        var comparisonUnitAtomList = cs
                            .ComparisonUnitArray2
                            .Select(ca => ca.DescendantContentAtoms())
                            .SelectMany(m => m)
                            .Select(ca =>
                                new ComparisonUnitAtom(ca.ContentElement, ca.AncestorElements, ca.Part)
                                {
                                    CorrelationStatus = CorrelationStatus.Inserted,
                                });
                        return comparisonUnitAtomList;
                    }
                    else
                        throw new OpenXmlPowerToolsException("Internal error");
                })
                .SelectMany(m => m)
                .ToList();

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in listOfComparisonUnitAtoms)
                    sb.Append(item.ToString()).Append(Environment.NewLine);
                var sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            return listOfComparisonUnitAtoms;
        }

        // for any deleted or inserted rows, we go into the w:trPr properties, and add the appropriate w:ins or w:del element, and therefore
        // when generating the document, the appropriate row will be marked as deleted or inserted.
        private static void MarkRowsAsDeletedOrInserted(WmlComparerSettings settings, List<CorrelatedSequence> correlatedSequence)
        {
            foreach (var dcs in correlatedSequence.Where(cs =>
                cs.CorrelationStatus == CorrelationStatus.Deleted || cs.CorrelationStatus == CorrelationStatus.Inserted))
            {
                // iterate through all deleted/inserted items in dcs.ComparisonUnitArray1/ComparisonUnitArray2
                var toIterateThrough = dcs.ComparisonUnitArray1;
                if (dcs.CorrelationStatus == CorrelationStatus.Inserted)
                    toIterateThrough = dcs.ComparisonUnitArray2;

                foreach (var ca in toIterateThrough)
                {
                    var cug = ca as ComparisonUnitGroup;

                    // this works because we will never see a table in this list, only rows.  If tables were in this list, would need to recursively
                    // go into children, but tables are always flattened in the LCS process.

                    // when we have a row, it is only necessary to find the first content atom of the row, then find the row ancestor, and then tweak
                    // the w:trPr

                    if (cug != null && cug.ComparisonUnitGroupType == ComparisonUnitGroupType.Row)
                    {
                        var firstContentAtom = cug.DescendantContentAtoms().FirstOrDefault();
                        if (firstContentAtom == null)
                            throw new OpenXmlPowerToolsException("Internal error");
                        var tr = firstContentAtom
                            .AncestorElements
                            .Reverse()
                            .FirstOrDefault(a => a.Name == W.tr);

                        if (tr == null)
                            throw new OpenXmlPowerToolsException("Internal error");
                        var trPr = tr.Element(W.trPr);
                        if (trPr == null)
                        {
                            trPr = new XElement(W.trPr);
                            tr.AddFirst(trPr);
                        }
                        XName revTrackElementName = null;
                        if (dcs.CorrelationStatus == CorrelationStatus.Deleted)
                            revTrackElementName = W.del;
                        else if (dcs.CorrelationStatus == CorrelationStatus.Inserted)
                            revTrackElementName = W.ins;
                        trPr.Add(new XElement(revTrackElementName,
                            new XAttribute(W.author, settings.AuthorForRevisions),
                            new XAttribute(W.id, s_MaxId++),
                            new XAttribute(W.date, settings.DateTimeForRevisions)));
                    }
                }
            }
        }

        public enum WmlComparerRevisionType
        {
            Inserted,
            Deleted,
        }

        public class WmlComparerRevision
        {
            public WmlComparerRevisionType RevisionType;
            public string Text;
            public string Author;
            public string Date;
            public XElement ContentXElement;
            public XElement RevisionXElement;
            public Uri PartUri;
            public string PartContentType;
        }

        private static XName[] RevElementsWithNoText = new XName[] {
            M.oMath,
            M.oMathPara,
            W.drawing,
        };

        public static List<WmlComparerRevision> GetRevisions(WmlDocument source)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(source.DocumentByteArray, 0, source.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    TestForInvalidContent(wDoc);
                    RemoveExistingPowerToolsMarkup(wDoc);

                    var contentParent = wDoc.MainDocumentPart.GetXDocument().Root.Element(W.body);
                    var atomList = WmlComparer.CreateComparisonUnitAtomList(wDoc.MainDocumentPart, contentParent).ToArray();

                    if (s_False)
                    {
                        var sb = new StringBuilder();
                        foreach (var item in atomList)
                            sb.Append(item.ToString() + Environment.NewLine);
                        var sbs = sb.ToString();
                        TestUtil.NotePad(sbs);
                    }

                    var grouped = atomList
                        .GroupAdjacent(a =>
                            {
                                var key = a.CorrelationStatus.ToString();
                                if (a.CorrelationStatus != CorrelationStatus.Equal)
                                {
                                    var rt = new XElement(a.RevTrackElement.Name,
                                        new XAttribute(XNamespace.Xmlns + "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
                                        a.RevTrackElement.Attributes().Where(a2 => a2.Name != W.id && a2.Name != PtOpenXml.Unid && a2.Name != PtOpenXml.PrevUnid));
                                    key += rt.ToString(SaveOptions.DisableFormatting);
                                }
                                return key;
                            })
                        .ToList();

                    var revisions = grouped
                        .Where(k => k.Key != "Equal")
                        .ToList();

                    if (s_False)
                    {
                        var sb = new StringBuilder();
                        foreach (var item in revisions)
                            sb.Append(item.Key + Environment.NewLine);
                        var sbs = sb.ToString();
                        TestUtil.NotePad(sbs);
                    }

                    var mainDocPartRevisionList = revisions
                        .Select(rg =>
                        {
                            var rev = new WmlComparerRevision();
                            if (rg.Key.StartsWith("Inserted"))
                                rev.RevisionType = WmlComparerRevisionType.Inserted;
                            else if (rg.Key.StartsWith("Deleted"))
                                rev.RevisionType = WmlComparerRevisionType.Deleted;
                            var revTrackElement = rg.First().RevTrackElement;
                            rev.RevisionXElement = revTrackElement;
                            rev.Author = (string)revTrackElement.Attribute(W.author);
                            rev.ContentXElement = rg.First().ContentElement;
                            rev.Date = (string)revTrackElement.Attribute(W.date);
                            rev.PartUri = wDoc.MainDocumentPart.Uri;
                            rev.PartContentType = wDoc.MainDocumentPart.ContentType;
                            if (!RevElementsWithNoText.Contains(rev.ContentXElement.Name))
                            {
                                rev.Text = rg
                                    .Select(rgc =>
                                        {
                                            if (rgc.ContentElement.Name == W.pPr)
                                                return Environment.NewLine;
                                            return rgc.ContentElement.Value;
                                        })
                                    .StringConcatenate();
                            }
                            return rev;
                        })
                        .ToList();

                    var footnotesRevisionList = GetFootnoteEndnoteRevisionList(wDoc.MainDocumentPart.FootnotesPart, W.footnote);
                    var endnotesRevisionList = GetFootnoteEndnoteRevisionList(wDoc.MainDocumentPart.EndnotesPart, W.endnote);
                    var finalRevisionList = mainDocPartRevisionList.Concat(footnotesRevisionList).Concat(endnotesRevisionList).ToList();
                    return finalRevisionList;
                }
            }
        }

        private static IEnumerable<WmlComparerRevision> GetFootnoteEndnoteRevisionList(OpenXmlPart footnotesEndnotesPart,
            XName footnoteEndnoteElementName)
        {
            if (footnotesEndnotesPart == null)
                return Enumerable.Empty<WmlComparerRevision>();

            var xDoc = footnotesEndnotesPart.GetXDocument();
            var footnotesEndnotes = xDoc.Root.Elements(footnoteEndnoteElementName);
            List<WmlComparerRevision> revisionsForPart = new List<WmlComparerRevision>();
            foreach (var fn in footnotesEndnotes)
            {
                var atomList = WmlComparer.CreateComparisonUnitAtomList(footnotesEndnotesPart, fn).ToArray();

                if (s_False)
                {
                    var sb = new StringBuilder();
                    foreach (var item in atomList)
                        sb.Append(item.ToString() + Environment.NewLine);
                    var sbs = sb.ToString();
                    TestUtil.NotePad(sbs);
                }

                var grouped = atomList
                    .GroupAdjacent(a =>
                        {
                            var key = a.CorrelationStatus.ToString();
                            if (a.CorrelationStatus != CorrelationStatus.Equal)
                            {
                                var rt = new XElement(a.RevTrackElement.Name,
                                    new XAttribute(XNamespace.Xmlns + "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
                                    a.RevTrackElement.Attributes().Where(a2 => a2.Name != W.id && a2.Name != PtOpenXml.Unid && a2.Name != PtOpenXml.PrevUnid));
                                key += rt.ToString(SaveOptions.DisableFormatting);
                            }
                            return key;
                        })
                    .ToList();

                var revisions = grouped
                    .Where(k => k.Key != "Equal")
                    .ToList();

                var thisNoteRevisionList = revisions
                    .Select(rg =>
                    {
                        var rev = new WmlComparerRevision();
                        if (rg.Key.StartsWith("Inserted"))
                            rev.RevisionType = WmlComparerRevisionType.Inserted;
                        else if (rg.Key.StartsWith("Deleted"))
                            rev.RevisionType = WmlComparerRevisionType.Deleted;
                        var revTrackElement = rg.First().RevTrackElement;
                        rev.RevisionXElement = revTrackElement;
                        rev.Author = (string)revTrackElement.Attribute(W.author);
                        rev.ContentXElement = rg.First().ContentElement;
                        rev.Date = (string)revTrackElement.Attribute(W.date);
                        rev.PartUri = footnotesEndnotesPart.Uri;
                        rev.PartContentType = footnotesEndnotesPart.ContentType;
                        if (!RevElementsWithNoText.Contains(rev.ContentXElement.Name))
                        {
                            rev.Text = rg
                                .Select(rgc =>
                                    {
                                        if (rgc.ContentElement.Name == W.pPr)
                                            return Environment.NewLine;
                                        return rgc.ContentElement.Value;
                                    })
                                .StringConcatenate();
                        }
                        return rev;
                    });

                foreach (var item in thisNoteRevisionList)
                    revisionsForPart.Add(item);
            }
            return revisionsForPart;
        }

        // prohibit
        // - altChunk
        // - subDoc
        // - contentPart
        // - text boxes and nested tables (for now)
        private static void TestForInvalidContent(WordprocessingDocument wDoc)
        {
            foreach (var part in wDoc.ContentParts())
            {
                var xDoc = part.GetXDocument();
                if (xDoc.Descendants(W.altChunk).Any())
                    throw new OpenXmlPowerToolsException("Unsupported document, contains w:altChunk");
                if (xDoc.Descendants(W.subDoc).Any())
                    throw new OpenXmlPowerToolsException("Unsupported document, contains w:subDoc");
                if (xDoc.Descendants(W.contentPart).Any())
                    throw new OpenXmlPowerToolsException("Unsupported document, contains w:contentPart");
                if (xDoc.Descendants(W.txbxContent).Any())
                    throw new OpenXmlPowerToolsException("Unsupported document, contains text boxes");
                if (xDoc.Descendants(W.tbl).Any(d => d.Ancestors(W.tbl).Any()))
                    throw new OpenXmlPowerToolsException("Unsupported document, contains nested tables");
            }
        }

        private static void RemoveExistingPowerToolsMarkup(WordprocessingDocument wDoc)
        {
            wDoc.MainDocumentPart
                .GetXDocument()
                .Root
                .Descendants()
                .Attributes()
                .Where(a => a.Name.Namespace == PtOpenXml.pt)
                .Where(a => a.Name != PtOpenXml.Unid && a.Name != PtOpenXml.PrevUnid)
                .Remove();
            wDoc.MainDocumentPart.PutXDocument();

            var fnPart = wDoc.MainDocumentPart.FootnotesPart;
            if (fnPart != null)
            {
                var fnXDoc = fnPart.GetXDocument();
                fnXDoc
                    .Root
                    .Descendants()
                    .Attributes()
                    .Where(a => a.Name.Namespace == PtOpenXml.pt)
                    .Where(a => a.Name != PtOpenXml.Unid || a.Name != PtOpenXml.PrevUnid)
                    .Remove();
                fnPart.PutXDocument();
            }

            var enPart = wDoc.MainDocumentPart.EndnotesPart;
            if (enPart != null)
            {
                var enXDoc = enPart.GetXDocument();
                enXDoc
                    .Root
                    .Descendants()
                    .Attributes()
                    .Where(a => a.Name.Namespace == PtOpenXml.pt)
                    .Where(a => a.Name != PtOpenXml.Unid || a.Name != PtOpenXml.PrevUnid)
                    .Remove();
                enPart.PutXDocument();
            }
        }

        private static void AddSha1HashToBlockLevelContent(OpenXmlPart part, XElement contentParent)
        {
            var blockLevelContentToAnnotate = contentParent
                .Descendants()
                .Where(d => ElementsToHaveSha1Hash.Contains(d.Name));

            foreach (var blockLevelContent in blockLevelContentToAnnotate)
            {
                var cloneBlockLevelContentForHashing = (XElement)CloneBlockLevelContentForHashing(part, blockLevelContent, true);
                var shaString = cloneBlockLevelContentForHashing.ToString(SaveOptions.DisableFormatting)
                    .Replace(" xmlns=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", "");
                var sha1Hash = WmlComparerUtil.SHA1HashStringForUTF8String(shaString);
                blockLevelContent.Add(new XAttribute(PtOpenXml.SHA1Hash, sha1Hash));
            }
        }

        static XName[] AttributesToTrimWhenCloning = new XName[] {
            WP14.anchorId,
            WP14.editId,
        };

        private static object CloneBlockLevelContentForHashing(OpenXmlPart mainDocumentPart, XNode node, bool includeRelatedParts)
        {
            var element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.bookmarkStart ||
                    element.Name == W.bookmarkEnd ||
                    element.Name == W.pPr ||
                    element.Name == W.rPr)
                    return null;

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
                        element.Nodes().Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts)));

                    var groupedRuns = clonedPara
                        .Elements()
                        .GroupAdjacent(e => e.Name == W.r &&
                            e.Elements().Count() == 1 &&
                            e.Element(W.t) != null);

                    var clonedParaWithGroupedRuns = new XElement(element.Name,
                        groupedRuns.Select(g =>
                        {
                            if (g.Key)
                            {
                                var newRun = (object)new XElement(W.r,
                                    new XElement(W.t,
                                        g.Select(t => t.Value).StringConcatenate()));
                                return newRun;
                            }
                            return g;
                        }));

                    return clonedParaWithGroupedRuns;
                }

                if (element.Name == W.r)
                {
                    var clonedRuns = element
                        .Elements()
                        .Where(e => e.Name != W.rPr)
                        .Select(rc => new XElement(W.r, CloneBlockLevelContentForHashing(mainDocumentPart, rc, includeRelatedParts)));
                    return clonedRuns;
                }

                if (element.Name == W.tbl)
                {
                    var clonedTable = new XElement(W.tbl,
                        element.Elements(W.tr).Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts)));
                    return clonedTable;
                }

                if (element.Name == W.tr)
                {
                    var clonedRow = new XElement(W.tr,
                        element.Elements(W.tc).Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts)));
                    return clonedRow;
                }

                if (element.Name == W.tc)
                {
                    var clonedCell = new XElement(W.tc,
                        element.Elements().Where(z => z.Name != W.tcPr).Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts)));
                    return clonedCell;
                }

                if (element.Name == W.txbxContent)
                {
                    var clonedTextbox = new XElement(W.txbxContent,
                        element.Elements().Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts)));
                    return clonedTextbox;
                }

                if (includeRelatedParts)
                {
                    if (ComparisonUnitWord.s_ElementsWithRelationshipIds.Contains(element.Name))
                    {
                        var newElement = new XElement(element.Name,
                            element.Attributes()
                                .Where(a => a.Name.Namespace != PtOpenXml.pt)
                                .Where(a => !AttributesToTrimWhenCloning.Contains(a.Name))
                                .Select(a =>
                                {
                                    if (!ComparisonUnitWord.s_RelationshipAttributeNames.Contains(a.Name))
                                        return a;
                                    var rId = (string)a;
                                    OpenXmlPart oxp = mainDocumentPart.GetPartById(rId);
                                    if (oxp == null)
                                        throw new FileFormatException("Invalid WordprocessingML Document");

                                    var anno = oxp.Annotation<PartSHA1HashAnnotation>();
                                    if (anno != null)
                                        return new XAttribute(a.Name, anno.Hash);

                                    if (!oxp.ContentType.EndsWith("xml"))
                                    {
                                        using (var str = oxp.GetStream())
                                        {
                                            byte[] ba;
                                            using (BinaryReader br = new BinaryReader(str))
                                            {
                                                ba = br.ReadBytes((int)str.Length);
                                            }
                                            var sha1 = WmlComparerUtil.SHA1HashStringForByteArray(ba);
                                            oxp.AddAnnotation(new PartSHA1HashAnnotation(sha1));
                                            return new XAttribute(a.Name, sha1);
                                        }
                                    }
                                    return null;
                                }),
                            element.Nodes().Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts)));
                        return newElement;
                    }
                }

                if (element.Name == VML.shape)
                {
                    return new XElement(element.Name,
                        element.Attributes()
                            .Where(a => a.Name.Namespace != PtOpenXml.pt)
                            .Where(a => a.Name != "style"),
                        element.Nodes().Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts)));
                }

                if (element.Name == O.OLEObject)
                {
                    return new XElement(element.Name,
                        element.Attributes()
                            .Where(a => a.Name.Namespace != PtOpenXml.pt)
                            .Where(a => a.Name != "ObjectID" && a.Name != R.id),
                        element.Nodes().Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts)));
                }

                return new XElement(element.Name,
                    element.Attributes()
                        .Where(a => a.Name.Namespace != PtOpenXml.pt)
                        .Where(a => !AttributesToTrimWhenCloning.Contains(a.Name)),
                    element.Nodes().Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n, includeRelatedParts)));
            }
            return node;
        }


        private static List<CorrelatedSequence> FindCommonAtBeginningAndEnd(CorrelatedSequence unknown, WmlComparerSettings settings)
        {
            int lengthToCompare = Math.Min(unknown.ComparisonUnitArray1.Length, unknown.ComparisonUnitArray2.Length);

            var countCommonAtBeginning = unknown
                .ComparisonUnitArray1
                .Take(lengthToCompare)
                .Zip(unknown.ComparisonUnitArray2,
                    (pu1, pu2) =>
                    {
                        return new
                        {
                            Pu1 = pu1,
                            Pu2 = pu2,
                        };
                    })
                    .TakeWhile(pair => pair.Pu1.SHA1Hash == pair.Pu2.SHA1Hash)
                    .Count();

            if (countCommonAtBeginning != 0 && ((double)countCommonAtBeginning / (double)lengthToCompare) < settings.DetailThreshold)
                countCommonAtBeginning = 0;

            var countCommonAtEnd = unknown
                .ComparisonUnitArray1
                .Skip(countCommonAtBeginning)
                .Reverse()
                .Take(lengthToCompare)
                .Zip(unknown
                    .ComparisonUnitArray2
                    .Skip(countCommonAtBeginning)
                    .Reverse()
                    .Take(lengthToCompare),
                    (pu1, pu2) =>
                    {
                        return new
                        {
                            Pu1 = pu1,
                            Pu2 = pu2,
                        };
                    })
                    .TakeWhile(pair => pair.Pu1.SHA1Hash == pair.Pu2.SHA1Hash)
                    .Count();

            // never start a common section with a paragraph mark.  However, it is OK to set two paragraph marks as equal.
            while (true)
            {
                if (countCommonAtEnd <= 1)
                    break;

                var firstCommon = unknown
                    .ComparisonUnitArray1
                    .Reverse()
                    .Take(countCommonAtEnd)
                    .LastOrDefault();

                var firstCommonWord = firstCommon as ComparisonUnitWord;
                if (firstCommonWord == null)
                    break;

                // if the word contains more than one atom, then not a paragraph mark
                if (firstCommonWord.Contents.Count() != 1)
                    break;

                var firstCommonAtom = firstCommonWord.Contents.First() as ComparisonUnitAtom;
                if (firstCommonAtom == null)
                    break;

                if (firstCommonAtom.ContentElement.Name != W.pPr)
                    break;

                countCommonAtEnd--;
            }

            bool isOnlyParagraphMark = false;
            if (countCommonAtEnd == 1)
            {
                var firstCommon = unknown
                    .ComparisonUnitArray1
                    .Reverse()
                    .Take(countCommonAtEnd)
                    .LastOrDefault();

                var firstCommonWord = firstCommon as ComparisonUnitWord;
                if (firstCommonWord != null)
                {
                    // if the word contains more than one atom, then not a paragraph mark
                    if (firstCommonWord.Contents.Count() == 1)
                    {
                        var firstCommonAtom = firstCommonWord.Contents.First() as ComparisonUnitAtom;
                        if (firstCommonAtom != null)
                        {
                            if (firstCommonAtom.ContentElement.Name == W.pPr)
                                isOnlyParagraphMark = true;
                        }
                    }
                }
            }

            if (!isOnlyParagraphMark && countCommonAtEnd != 0 && ((double)countCommonAtEnd / (double)lengthToCompare) < settings.DetailThreshold)
                countCommonAtEnd = 0;

            // If the following test is not there, the test below sets the end paragraph mark of the entire document equal to the end paragraph
            // mark of the first paragraph in the other document, causing lines to be out of order.
            // [InlineData("WC010-Para-Before-Table-Unmodified.docx", "WC010-Para-Before-Table-Mod.docx", 3)]
            if (isOnlyParagraphMark)
                countCommonAtEnd = 0;

            if (countCommonAtBeginning == 0 && countCommonAtEnd == 0)
                return null;

            var newSequence = new List<CorrelatedSequence>();
            if (countCommonAtBeginning != 0)
            {
                CorrelatedSequence cs = new CorrelatedSequence();
                cs.CorrelationStatus = CorrelationStatus.Equal;

                cs.ComparisonUnitArray1 = unknown
                    .ComparisonUnitArray1
                    .Take(countCommonAtBeginning)
                    .ToArray();

                cs.ComparisonUnitArray2 = unknown
                    .ComparisonUnitArray2
                    .Take(countCommonAtBeginning)
                    .ToArray();

                newSequence.Add(cs);
            }

            var middleLeft = unknown
                .ComparisonUnitArray1
                .Skip(countCommonAtBeginning)
                .SkipLast(countCommonAtEnd)
                .ToArray();

            var middleRight = unknown
                .ComparisonUnitArray2
                .Skip(countCommonAtBeginning)
                .SkipLast(countCommonAtEnd)
                .ToArray();

            if (middleLeft.Length > 0 && middleRight.Length == 0)
            {
                CorrelatedSequence cs = new CorrelatedSequence();
                cs.CorrelationStatus = CorrelationStatus.Deleted;
                cs.ComparisonUnitArray1 = middleLeft;
                cs.ComparisonUnitArray2 = null;
                newSequence.Add(cs);
            }
            else if (middleLeft.Length == 0 && middleRight.Length > 0)
            {
                CorrelatedSequence cs = new CorrelatedSequence();
                cs.CorrelationStatus = CorrelationStatus.Inserted;
                cs.ComparisonUnitArray1 = null;
                cs.ComparisonUnitArray2 = middleRight;
                newSequence.Add(cs);
            }
            else if (middleLeft.Length > 0 && middleRight.Length > 0)
            {
                CorrelatedSequence cs = new CorrelatedSequence();
                cs.CorrelationStatus = CorrelationStatus.Unknown;
                cs.ComparisonUnitArray1 = middleLeft;
                cs.ComparisonUnitArray2 = middleRight;
                newSequence.Add(cs);
            }

            if (countCommonAtEnd != 0)
            {
                CorrelatedSequence cs = new CorrelatedSequence();
                cs.CorrelationStatus = CorrelationStatus.Equal;

                cs.ComparisonUnitArray1 = unknown
                    .ComparisonUnitArray1
                    .Skip(countCommonAtBeginning + middleLeft.Length)
                    .ToArray();

                cs.ComparisonUnitArray2 = unknown
                    .ComparisonUnitArray2
                    .Skip(countCommonAtBeginning + middleRight.Length)
                    .ToArray();

                newSequence.Add(cs);
            }
            return newSequence;
        }

        private static void MoveLastSectPrToChildOfBody(XDocument newXDoc)
        {
            var lastParaWithSectPr = newXDoc
                .Root
                .Elements(W.body)
                .Elements(W.p)
                .Where(p => p.Elements(W.pPr).Elements(W.sectPr).Any())
                .LastOrDefault();
            if (lastParaWithSectPr != null)
            {
                newXDoc.Root.Element(W.body).Add(lastParaWithSectPr.Elements(W.pPr).Elements(W.sectPr));
                lastParaWithSectPr.Elements(W.pPr).Elements(W.sectPr).Remove();
            }
        }

        private static int s_MaxId = 0;

        private static object ProduceNewWmlMarkupFromCorrelatedSequence(OpenXmlPart part,
            IEnumerable<ComparisonUnitAtom> comparisonUnitAtomList,
            WmlComparerSettings settings)
        {
            // fabricate new MainDocumentPart from correlatedSequence
            s_MaxId = 0;
            var newBodyChildren = CoalesceRecurse(part, comparisonUnitAtomList, 0, settings);
            return newBodyChildren;
        }

        private static void FixUpDocPrIds(WordprocessingDocument wDoc)
        {
            var elementToFind = WP.docPr;
            var docPrToChange = wDoc
                .ContentParts()
                .Select(cp => cp.GetXDocument())
                .Select(xd => xd.Descendants().Where(d => d.Name == elementToFind))
                .SelectMany(m => m);
            var nextId = 1;
            foreach (var item in docPrToChange)
            {
                var idAtt = item.Attribute("id");
                if (idAtt != null)
                    idAtt.Value = (nextId++).ToString();
            }
            foreach (var cp in wDoc.ContentParts())
                cp.PutXDocument();
        }

        private static void FixUpRevMarkIds(WordprocessingDocument wDoc)
        {
            var revMarksToChange = wDoc
                .ContentParts()
                .Select(cp => cp.GetXDocument())
                .Select(xd => xd.Descendants().Where(d => d.Name == W.ins || d.Name == W.del))
                .SelectMany(m => m);
            var nextId = 0;
            foreach (var item in revMarksToChange)
            {
                var idAtt = item.Attribute(W.id);
                if (idAtt != null)
                    idAtt.Value = (nextId++).ToString();
            }
            foreach (var cp in wDoc.ContentParts())
                cp.PutXDocument();
        }

        private static void FixUpShapeIds(WordprocessingDocument wDoc)
        {
            var elementToFind = VML.shape;
            var shapeIdsToChange = wDoc
                .ContentParts()
                .Select(cp => cp.GetXDocument())
                .Select(xd => xd.Descendants().Where(d => d.Name == elementToFind))
                .SelectMany(m => m);
            var nextId = 1;
            foreach (var item in shapeIdsToChange)
            {
                var idAtt = item.Attribute("id");
                if (idAtt != null)
                    idAtt.Value = (nextId++).ToString();
            }
            foreach (var cp in wDoc.ContentParts())
                cp.PutXDocument();
        }

        private static void FixUpShapeTypeIds(WordprocessingDocument wDoc)
        {
            var elementToFind = VML.shapetype;
            var shapeTypeIdsToChange = wDoc
                .ContentParts()
                .Select(cp => cp.GetXDocument())
                .Select(xd => xd.Descendants().Where(d => d.Name == elementToFind))
                .SelectMany(m => m);
            var nextId = 1;
            foreach (var item in shapeTypeIdsToChange)
            {
                var idAtt = item.Attribute("id");
                if (idAtt != null)
                    idAtt.Value = (nextId++).ToString();
            }
            foreach (var cp in wDoc.ContentParts())
                cp.PutXDocument();
        }

        private static object CoalesceRecurse(OpenXmlPart part, IEnumerable<ComparisonUnitAtom> list, int level, WmlComparerSettings settings)
        {
            var grouped = list.GroupBy(ca =>
                {
                    if (level >= ca.AncestorElements.Length)
                        return "";
                    return (string)ca.AncestorElements[level].Attribute(PtOpenXml.Unid);
                })
                .Where(g => g.Key != "");

            // if there are no deeper children, then we're done.
            if (!grouped.Any())
                return null;

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var group in grouped)
                {
                    sb.AppendFormat("Group Key: {0}", group.Key);
                    sb.Append(Environment.NewLine);
                    foreach (var groupChildItem in group)
                    {
                        sb.Append("  ");
                        sb.Append(groupChildItem.ToString(0));
                        sb.Append(Environment.NewLine);
                    }
                    sb.Append(Environment.NewLine);
                }
                var sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            var elementList = grouped
                .Select(g =>
                {
                    var ancestorBeingConstructed = g.First().AncestorElements[level]; // these will all be the same, by definition

                    // need to group by corr stat
                    var groupedChildren = g
                        .GroupAdjacent(gc =>
                            {
                                var key = "";
                                if (level < (gc.AncestorElements.Length - 1))
                                {
                                    var anc = gc.AncestorElements[level + 1];
                                    key = (string)anc.Attribute(PtOpenXml.Unid);
                                }
                                key += "|" + gc.CorrelationStatus.ToString();
                                return key;
                            })
                        .ToList();

                    if (ancestorBeingConstructed.Name == W.p)
                    {
                        var newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                if (gc.Key.Split('|')[0] == "")
                                    return (object)gc.Select(gcc =>
                                    {
                                        var dup = new XElement(gcc.ContentElement);
                                        if (gcc.CorrelationStatus == CorrelationStatus.Deleted)
                                            dup.Add(new XAttribute(PtOpenXml.Status, "Deleted"));
                                        else if (gcc.CorrelationStatus == CorrelationStatus.Inserted)
                                            dup.Add(new XAttribute(PtOpenXml.Status, "Inserted"));
                                        return dup;
                                    });
                                else
                                {
                                    return CoalesceRecurse(part, gc, level + 1, settings);
                                }
                            })
                            .ToList();

                        var newPara = new XElement(W.p,
                            ancestorBeingConstructed.Attributes(),
                            newChildElements);
                        return newPara;
                    }

                    if (ancestorBeingConstructed.Name == W.r)
                    {
                        var newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                if (gc.Key.Split('|')[0] == "")
                                    return (object)gc.Select(gcc =>
                                    {
                                        var dup = new XElement(gcc.ContentElement);
                                        if (gcc.CorrelationStatus == CorrelationStatus.Deleted)
                                            dup.Add(new XAttribute(PtOpenXml.Status, "Deleted"));
                                        else if (gcc.CorrelationStatus == CorrelationStatus.Inserted)
                                            dup.Add(new XAttribute(PtOpenXml.Status, "Inserted"));
                                        return dup;
                                    });
                                else
                                {
                                    return CoalesceRecurse(part, gc, level + 1, settings);
                                }
                            })
                            .ToList();

                        XElement rPr = ancestorBeingConstructed.Element(W.rPr);
                        var newRun = new XElement(W.r,
                            ancestorBeingConstructed.Attributes(),
                            rPr,
                            newChildElements);
                        return newRun;
                    }

                    if (ancestorBeingConstructed.Name == W.t)
                    {
                        var newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                var textOfTextElement = gc.Select(gce => gce.ContentElement.Value).StringConcatenate();
                                var del = gc.First().CorrelationStatus == CorrelationStatus.Deleted;
                                var ins = gc.First().CorrelationStatus == CorrelationStatus.Inserted;
                                if (del)
                                    return (object)(new XElement(W.delText,
                                        new XAttribute(PtOpenXml.Status, "Deleted"),
                                        GetXmlSpaceAttribute(textOfTextElement),
                                        textOfTextElement));
                                else if (ins)
                                    return (object)(new XElement(W.t,
                                        new XAttribute(PtOpenXml.Status, "Inserted"),
                                        GetXmlSpaceAttribute(textOfTextElement),
                                        textOfTextElement));
                                else
                                    return (object)(new XElement(W.t,
                                        GetXmlSpaceAttribute(textOfTextElement),
                                        textOfTextElement));
                            })
                            .ToList();
                        return newChildElements;
                    }

                    if (ancestorBeingConstructed.Name == W.drawing)
                    {
                        var newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                var del = gc.First().CorrelationStatus == CorrelationStatus.Deleted;
                                var ins = gc.First().CorrelationStatus == CorrelationStatus.Inserted;
                                if (del)
                                {
                                    return (object)gc.Select(gcc =>
                                    {
                                        var newDrawing = new XElement(gcc.ContentElement);
                                        newDrawing.Add(new XAttribute(PtOpenXml.Status, "Deleted"));

                                        var openXmlPartOfDeletedContent = gc.First().Part;
                                        var openXmlPartInNewDocument = part;
                                        return gc.Select(gce =>
                                        {
                                            Package packageOfDeletedContent = openXmlPartOfDeletedContent.OpenXmlPackage.Package;
                                            Package packageOfNewContent = openXmlPartInNewDocument.OpenXmlPackage.Package;
                                            PackagePart partInDeletedDocument = packageOfDeletedContent.GetPart(part.Uri);
                                            PackagePart partInNewDocument = packageOfNewContent.GetPart(part.Uri);
                                            return MoveRelatedPartsToDestination(partInDeletedDocument, partInNewDocument, newDrawing);
                                        });
                                    });
                                }
                                else if (ins)
                                {
                                    return gc.Select(gcc =>
                                    {
                                        var newDrawing = new XElement(gcc.ContentElement);
                                        newDrawing.Add(new XAttribute(PtOpenXml.Status, "Inserted"));

                                        var openXmlPartOfInsertedContent = gc.First().Part;
                                        var openXmlPartInNewDocument = part;
                                        return gc.Select(gce =>
                                        {
                                            Package packageOfSourceContent = openXmlPartOfInsertedContent.OpenXmlPackage.Package;
                                            Package packageOfNewContent = openXmlPartInNewDocument.OpenXmlPackage.Package;
                                            PackagePart partInDeletedDocument = packageOfSourceContent.GetPart(part.Uri);
                                            PackagePart partInNewDocument = packageOfNewContent.GetPart(part.Uri);
                                            return MoveRelatedPartsToDestination(partInDeletedDocument, partInNewDocument, newDrawing);
                                        });
                                    });
                                }
                                else
                                {
                                    return gc.Select(gcc =>
                                    {
                                        return gcc.ContentElement;
                                    });
                                }
                            })
                            .ToList();
                        return newChildElements;
                    }

                    if (ancestorBeingConstructed.Name == M.oMath || ancestorBeingConstructed.Name == M.oMathPara)
                    {
                        var newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                var del = gc.First().CorrelationStatus == CorrelationStatus.Deleted;
                                var ins = gc.First().CorrelationStatus == CorrelationStatus.Inserted;
                                if (del)
                                {
                                    return gc.Select(gcc =>
                                    {
                                        return new XElement(W.del,
                                            new XAttribute(W.author, settings.AuthorForRevisions),
                                            new XAttribute(W.id, s_MaxId++),
                                            new XAttribute(W.date, settings.DateTimeForRevisions),
                                            gcc.ContentElement);
                                    });
                                }
                                else if (ins)
                                {
                                    return gc.Select(gcc =>
                                    {
                                        return new XElement(W.ins,
                                            new XAttribute(W.author, settings.AuthorForRevisions),
                                            new XAttribute(W.id, s_MaxId++),
                                            new XAttribute(W.date, settings.DateTimeForRevisions),
                                            gcc.ContentElement);
                                    });
                                }
                                else
                                {
                                    return gc.Select(gcc => gcc.ContentElement);
                                }
                            })
                            .ToList();
                        return newChildElements;
                    }

                    if (AllowableRunChildren.Contains(ancestorBeingConstructed.Name))
                    {
                        var newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                var del = gc.First().CorrelationStatus == CorrelationStatus.Deleted;
                                var ins = gc.First().CorrelationStatus == CorrelationStatus.Inserted;
                                if (del)
                                {
                                    return gc.Select(gcc =>
                                    {
                                        var dup = new XElement(ancestorBeingConstructed.Name,
                                            ancestorBeingConstructed.Attributes(),
                                            new XAttribute(PtOpenXml.Status, "Deleted"));
                                        return dup;
                                    });
                                }
                                else if (ins)
                                {
                                    return gc.Select(gcc =>
                                    {
                                        var dup = new XElement(ancestorBeingConstructed.Name,
                                            ancestorBeingConstructed.Attributes(),
                                            new XAttribute(PtOpenXml.Status, "Inserted"));
                                        return dup;
                                    });
                                }
                                else
                                {
                                    return gc.Select(gcc => gcc.ContentElement);
                                }
                            })
                            .ToList();
                        return newChildElements;
                    }

                    if (ancestorBeingConstructed.Name == W.tbl)
                        return ReconstructElement(part, g, ancestorBeingConstructed, W.tblPr, W.tblGrid, level, settings);
                    if (ancestorBeingConstructed.Name == W.tr)
                        return ReconstructElement(part, g, ancestorBeingConstructed, W.trPr, null, level, settings);
                    if (ancestorBeingConstructed.Name == W.tc)
                        return ReconstructElement(part, g, ancestorBeingConstructed, W.tcPr, null, level, settings);
                    if (ancestorBeingConstructed.Name == W.sdt)
                        return ReconstructElement(part, g, ancestorBeingConstructed, W.sdtPr, W.sdtEndPr, level, settings);
                    if (ancestorBeingConstructed.Name == W.pict)
                        return ReconstructElement(part, g, ancestorBeingConstructed, VML.shapetype, null, level, settings);
                    return (object)ReconstructElement(part, g, ancestorBeingConstructed, null, null, level, settings);
                })
                .ToList();
            return elementList;
        }

        private static XElement MoveRelatedPartsToDestination(PackagePart partOfDeletedContent, PackagePart partInNewDocument,
            XElement contentElement)
        {
            var elementsToUpdate = contentElement
                .Descendants()
                .Where(d => d.Attributes().Any(a => ComparisonUnitWord.s_RelationshipAttributeNames.Contains(a.Name)))
                .ToList();
            foreach (var element in elementsToUpdate)
            {
                var attributesToUpdate = element
                    .Attributes()
                    .Where(a => ComparisonUnitWord.s_RelationshipAttributeNames.Contains(a.Name))
                    .ToList();
                foreach (var att in attributesToUpdate)
                {
                    var rId = (string)att;

                    var relationshipForDeletedPart = partOfDeletedContent.GetRelationship(rId);
                    if (relationshipForDeletedPart == null)
                        throw new FileFormatException("Invalid document");

                    Uri targetUri = PackUriHelper
                        .ResolvePartUri(
                           new Uri(partOfDeletedContent.Uri.ToString(), UriKind.Relative),
                                 relationshipForDeletedPart.TargetUri);

                    var relatedPackagePart = partOfDeletedContent.Package.GetPart(targetUri);
                    var uriSplit = relatedPackagePart.Uri.ToString().Split('/');
                    var last = uriSplit[uriSplit.Length - 1].Split('.');
                    string uriString = null;
                    if (last.Length == 2)
                    {
                        uriString = uriSplit.SkipLast(1).Select(p => p + "/").StringConcatenate() +
                            "P" + Guid.NewGuid().ToString().Replace("-", "") + "." + last[1];
                    }
                    else
                    {
                        uriString = uriSplit.SkipLast(1).Select(p => p + "/").StringConcatenate() +
                            "P" + Guid.NewGuid().ToString().Replace("-", "");
                    }
                    Uri uri = null;
                    if (relatedPackagePart.Uri.IsAbsoluteUri)
                        uri = new Uri(uriString, UriKind.Absolute);
                    else
                        uri = new Uri(uriString, UriKind.Relative);

                    var newPart = partInNewDocument.Package.CreatePart(uri, relatedPackagePart.ContentType);
                    using (var oldPartStream = relatedPackagePart.GetStream())
                    using (var newPartStream = newPart.GetStream())
                        FileUtils.CopyStream(oldPartStream, newPartStream);

                    var newRid = "R" + Guid.NewGuid().ToString().Replace("-", "");
                    partInNewDocument.CreateRelationship(newPart.Uri, TargetMode.Internal, relationshipForDeletedPart.RelationshipType, newRid);
                    att.Value = newRid;

                    if (newPart.ContentType.EndsWith("xml"))
                    {
                        XDocument newPartXDoc = null;
                        using (var stream = newPart.GetStream())
                        {
                            newPartXDoc = XDocument.Load(stream);
                            MoveRelatedPartsToDestination(relatedPackagePart, newPart, newPartXDoc.Root);
                        }
                        using (var stream = newPart.GetStream())
                            newPartXDoc.Save(stream);
                    }
                }
            }
            return contentElement;
        }

        private static XAttribute GetXmlSpaceAttribute(string textOfTextElement)
        {
            if (char.IsWhiteSpace(textOfTextElement[0]) ||
                char.IsWhiteSpace(textOfTextElement[textOfTextElement.Length - 1]))
                return new XAttribute(XNamespace.Xml + "space", "preserve");
            return null;
        }

        private static XElement ReconstructElement(OpenXmlPart part, IGrouping<string, ComparisonUnitAtom> g, XElement ancestorBeingConstructed, XName props1XName,
            XName props2XName, int level, WmlComparerSettings settings)
        {
            var newChildElements = CoalesceRecurse(part, g, level + 1, settings);
            object props1 = null;
            if (props1XName != null)
                props1 = ancestorBeingConstructed.Elements(props1XName);
            object props2 = null;
            if (props2XName != null)
                props2 = ancestorBeingConstructed.Elements(props2XName);

            var reconstructedElement = new XElement(ancestorBeingConstructed.Name,
                ancestorBeingConstructed.Attributes(),
                props1, props2, newChildElements);
            return reconstructedElement;
        }

        private static List<CorrelatedSequence> Lcs(ComparisonUnit[] cu1, ComparisonUnit[] cu2, WmlComparerSettings settings)
        {
            // set up initial state - one CorrelatedSequence, UnKnown, contents == entire sequences (both)
            CorrelatedSequence cs = new CorrelatedSequence()
            {
                CorrelationStatus = OpenXmlPowerTools.CorrelationStatus.Unknown,
                ComparisonUnitArray1 = cu1,
                ComparisonUnitArray2 = cu2,
            };
            List<CorrelatedSequence> csList = new List<CorrelatedSequence>()
            {
                cs
            };

            while (true)
            {
                if (s_False)
                {
                    var sb = new StringBuilder();
                    foreach (var item in csList)
                        sb.Append(item.ToString()).Append(Environment.NewLine);
                    var sbs = sb.ToString();
                    TestUtil.NotePad(sbs);
                }

                var unknown = csList
                    .FirstOrDefault(z => z.CorrelationStatus == CorrelationStatus.Unknown);
                if (unknown != null)
                {
                    if (s_False)
                    {
                        var sb = new StringBuilder();
                        sb.Append(unknown.ToString());
                        var sbs = sb.ToString();
                        TestUtil.NotePad(sbs);
                    }

                    var newSequence = FindCommonAtBeginningAndEnd(unknown, settings);
                    if (newSequence == null)
                    {
                        newSequence = DoLcsAlgorithm(unknown, settings);
                    }

                    var indexOfUnknown = csList.IndexOf(unknown);
                    csList.Remove(unknown);

                    newSequence.Reverse();
                    foreach (var item in newSequence)
                        csList.Insert(indexOfUnknown, item);

                    continue;
                }

                return csList;
            }
        }

        private static List<CorrelatedSequence> DoLcsAlgorithm(CorrelatedSequence unknown, WmlComparerSettings settings)
        {
            var cul1 = unknown.ComparisonUnitArray1;
            var cul2 = unknown.ComparisonUnitArray2;
            int currentLongestCommonSequenceLength = 0;
            int currentI1 = -1;
            int currentI2 = -1;
            for (int i1 = 0; i1 < cul1.Length; i1++)
            {
                for (int i2 = 0; i2 < cul2.Length; i2++)
                {
                    var thisSequenceLength = 0;
                    var thisI1 = i1;
                    var thisI2 = i2;
                    while (true)
                    {
                        if (cul1[thisI1].SHA1Hash == cul2[thisI2].SHA1Hash)
                        {
                            thisI1++;
                            thisI2++;
                            thisSequenceLength++;
                            if (thisI1 == cul1.Length || thisI2 == cul2.Length)
                            {
                                if (thisSequenceLength > currentLongestCommonSequenceLength)
                                {
                                    currentLongestCommonSequenceLength = thisSequenceLength;
                                    currentI1 = i1;
                                    currentI2 = i2;
                                }
                                break;
                            }
                            continue;
                        }
                        else
                        {
                            if (thisSequenceLength > currentLongestCommonSequenceLength)
                            {
                                currentLongestCommonSequenceLength = thisSequenceLength;
                                currentI1 = i1;
                                currentI2 = i2;
                            }
                            break;
                        }
                    }
                }
            }

            // never start a common section with a paragraph mark.
            while (true)
            {
                if (currentLongestCommonSequenceLength <= 1)
                    break;

                var firstCommon = cul1[currentI1];

                var firstCommonWord = firstCommon as ComparisonUnitWord;
                if (firstCommonWord == null)
                    break;

                // if the word contains more than one atom, then not a paragraph mark
                if (firstCommonWord.Contents.Count() != 1)
                    break;

                var firstCommonAtom = firstCommonWord.Contents.First() as ComparisonUnitAtom;
                if (firstCommonAtom == null)
                    break;

                if (firstCommonAtom.ContentElement.Name != W.pPr)
                    break;

                --currentLongestCommonSequenceLength;
                if (currentLongestCommonSequenceLength == 0)
                {
                    currentI1 = -1;
                    currentI2 = -1;
                }
                else
                {
                    ++currentI1;
                    ++currentI2;
                }
            }

            bool isOnlyParagraphMark = false;
            if (currentLongestCommonSequenceLength == 1)
            {
                var firstCommon = cul1[currentI1];

                var firstCommonWord = firstCommon as ComparisonUnitWord;
                if (firstCommonWord != null)
                {
                    // if the word contains more than one atom, then not a paragraph mark
                    if (firstCommonWord.Contents.Count() == 1)
                    {
                        var firstCommonAtom = firstCommonWord.Contents.First() as ComparisonUnitAtom;
                        if (firstCommonAtom != null)
                        {
                            if (firstCommonAtom.ContentElement.Name == W.pPr)
                                isOnlyParagraphMark = true;
                        }
                    }
                }
            }

            // don't match just a single space
            if (currentLongestCommonSequenceLength == 1)
            {
                var cuw2 = cul2[currentI2] as ComparisonUnitAtom;
                if (cuw2 != null)
                {
                    if (cuw2.ContentElement.Name == W.t && cuw2.ContentElement.Value == " ")
                    {
                        currentI1 = -1;
                        currentI2 = -1;
                        currentLongestCommonSequenceLength = 0;
                    }
                }
            }

            // if we are only looking at text, and if the longest common subsequence is less than 15% of the whole, then forget it,
            // don't find that LCS.
            if (!isOnlyParagraphMark && currentLongestCommonSequenceLength > 0)
            {
                var anyButWord1 = cul1.Any(cu => (cu as ComparisonUnitWord) == null);
                var anyButWord2 = cul2.Any(cu => (cu as ComparisonUnitWord) == null);
                if (!anyButWord1 && !anyButWord2)
                {
                    var maxLen = Math.Max(cul1.Length, cul2.Length);
                    if (((double)currentLongestCommonSequenceLength / (double)maxLen) < settings.DetailThreshold)
                    {
                        currentI1 = -1;
                        currentI2 = -1;
                        currentLongestCommonSequenceLength = 0;
                    }
                }
            }

            var newListOfCorrelatedSequence = new List<CorrelatedSequence>();
            if (currentI1 == -1 && currentI2 == -1)
            {
                var leftLength = unknown.ComparisonUnitArray1.Length;
                var leftTables = unknown.ComparisonUnitArray1.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Table).Count();
                var leftRows = unknown.ComparisonUnitArray1.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Row).Count();
                var leftCells = unknown.ComparisonUnitArray1.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Cell).Count();
                var leftParagraphs = unknown.ComparisonUnitArray1.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Paragraph).Count();
                var leftTextboxes = unknown.ComparisonUnitArray1.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Textbox).Count();
                var leftWords = unknown.ComparisonUnitArray1.OfType<ComparisonUnitWord>().Count();

                var rightLength = unknown.ComparisonUnitArray2.Length;
                var rightTables = unknown.ComparisonUnitArray2.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Table).Count();
                var rightRows = unknown.ComparisonUnitArray2.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Row).Count();
                var rightCells = unknown.ComparisonUnitArray2.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Cell).Count();
                var rightParagraphs = unknown.ComparisonUnitArray2.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Paragraph).Count();
                var rightTextboxes = unknown.ComparisonUnitArray2.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Textbox).Count();
                var rightWords = unknown.ComparisonUnitArray2.OfType<ComparisonUnitWord>().Count();

                // if either side has both words, rows and text boxes, then we need to separate out into separate unknown correlated sequences
                // group adjacent based on whether word, row, or textbox
                // in most cases, the count of groups will be the same, but they may differ
                // if the first group on either side is word, then create a deleted or inserted corr sequ for it.
                // then have counter on both sides pointing to the first matched pairs of rows
                // create an unknown corr sequ for it.
                // increment both counters
                // if one is at end but the other is not, then tag the remaining content as inserted or deleted, and done.
                // if both are at the end, then done
                // return the new list of corr sequ

                var leftOnlyWordsRowsTextboxes = leftLength == leftWords + leftRows + leftTextboxes;
                var rightOnlyWordsRowsTextboxes = rightLength == rightWords + rightRows + rightTextboxes;
                if ((leftWords > 0 || rightWords > 0) &&
                    (leftRows > 0 || rightRows > 0 || leftTextboxes > 0 || rightTextboxes > 0) &&
                    (leftOnlyWordsRowsTextboxes && rightOnlyWordsRowsTextboxes))
                {
                    var leftGrouped = unknown
                        .ComparisonUnitArray1
                        .GroupAdjacent(cu =>
                        {
                            if (cu is ComparisonUnitWord)
                            {
                                return "Word";
                            }
                            else
                            {
                                var cug = cu as ComparisonUnitGroup;
                                if (cug.ComparisonUnitGroupType == ComparisonUnitGroupType.Row)
                                    return "Row";
                                if (cug.ComparisonUnitGroupType == ComparisonUnitGroupType.Textbox)
                                    return "Textbox";
                                throw new OpenXmlPowerToolsException("Internal error");
                            }
                        })
                        .ToArray();
                    var rightGrouped = unknown
                        .ComparisonUnitArray2
                        .GroupAdjacent(cu =>
                        {
                            if (cu is ComparisonUnitWord)
                            {
                                return "Word";
                            }
                            else
                            {
                                var cug = cu as ComparisonUnitGroup;
                                if (cug.ComparisonUnitGroupType == ComparisonUnitGroupType.Row)
                                    return "Row";
                                if (cug.ComparisonUnitGroupType == ComparisonUnitGroupType.Textbox)
                                    return "Textbox";
                                throw new OpenXmlPowerToolsException("Internal error");
                            }
                        })
                        .ToArray();
                    int iLeft = 0;
                    int iRight = 0;

                    // create an unknown corr sequ for it.
                    // increment both counters
                    // if one is at end but the other is not, then tag the remaining content as inserted or deleted, and done.
                    // if both are at the end, then done
                    // return the new list of corr sequ

                    while (true)
                    {
                        if (leftGrouped[iLeft].Key == rightGrouped[iRight].Key)
                        {
                            var unknownCorrelatedSequence = new CorrelatedSequence();
                            unknownCorrelatedSequence.ComparisonUnitArray1 = leftGrouped[iLeft].ToArray();
                            unknownCorrelatedSequence.ComparisonUnitArray2 = rightGrouped[iRight].ToArray();
                            unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                            newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
                            ++iLeft;
                            ++iRight;
                        }
                        else if (leftGrouped[iLeft].Key == "Word" && rightGrouped[iRight].Key != "Word")
                        {
                            var deletedCorrelatedSequence = new CorrelatedSequence();
                            deletedCorrelatedSequence.ComparisonUnitArray1 = leftGrouped[iLeft].ToArray();
                            deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                            deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                            newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                            ++iLeft;
                        }
                        else if (leftGrouped[iLeft].Key != "Word" && rightGrouped[iRight].Key == "Word")
                        {
                            var insertedCorrelatedSequence = new CorrelatedSequence();
                            insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                            insertedCorrelatedSequence.ComparisonUnitArray2 = rightGrouped[iRight].ToArray();
                            insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                            ++iRight;
                        }

                        if (iLeft == leftGrouped.Length && iRight == rightGrouped.Length)
                            return newListOfCorrelatedSequence;

                        // if there is content on the left, but not content on the right
                        if (iRight == rightGrouped.Length)
                        {
                            for (int j = iLeft; j < leftGrouped.Length; j++)
                            {
                                var deletedCorrelatedSequence = new CorrelatedSequence();
                                deletedCorrelatedSequence.ComparisonUnitArray1 = leftGrouped[j].ToArray();
                                deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                                deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                            }
                            return newListOfCorrelatedSequence;
                        }
                        // there is content on the right but not on the left
                        else if (iLeft == leftGrouped.Length) 
                        {
                            for (int j = iRight; j < rightGrouped.Length; j++)
                            {
                                var insertedCorrelatedSequence = new CorrelatedSequence();
                                insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                                insertedCorrelatedSequence.ComparisonUnitArray2 = rightGrouped[j].ToArray();
                                insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                            }
                            return newListOfCorrelatedSequence;
                        }
                        // else continue on next round.
                    }
                }

                // if both sides contain tables and paragraphs, then split into multiple unknown corr sequ
                if (leftTables > 0 && rightTables > 0 &&
                    leftParagraphs > 0 && rightParagraphs > 0 &&
                    (leftLength > 1 || rightLength > 1))
                {
                    var leftGrouped = unknown
                        .ComparisonUnitArray1
                        .GroupAdjacent(cu =>
                        {
                            var cug = cu as ComparisonUnitGroup;
                            if (cug.ComparisonUnitGroupType == ComparisonUnitGroupType.Table)
                                return "Table";
                            else
                                return "Para";
                        })
                        .ToArray();
                    var rightGrouped = unknown
                        .ComparisonUnitArray2
                        .GroupAdjacent(cu =>
                        {
                            var cug = cu as ComparisonUnitGroup;
                            if (cug.ComparisonUnitGroupType == ComparisonUnitGroupType.Table)
                                return "Table";
                            else
                                return "Para";
                        })
                        .ToArray();
                    int iLeft = 0;
                    int iRight = 0;

                    // create an unknown corr sequ for it.
                    // increment both counters
                    // if one is at end but the other is not, then tag the remaining content as inserted or deleted, and done.
                    // if both are at the end, then done
                    // return the new list of corr sequ

                    while (true)
                    {
                        if ((leftGrouped[iLeft].Key == "Table" && rightGrouped[iRight].Key == "Table") ||
                            (leftGrouped[iLeft].Key == "Para" && rightGrouped[iRight].Key == "Para"))
                        {
                            var unknownCorrelatedSequence = new CorrelatedSequence();
                            unknownCorrelatedSequence.ComparisonUnitArray1 = leftGrouped[iLeft].ToArray();
                            unknownCorrelatedSequence.ComparisonUnitArray2 = rightGrouped[iRight].ToArray();
                            unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                            newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
                            ++iLeft;
                            ++iRight;
                        }
                        else if (leftGrouped[iLeft].Key == "Para" && rightGrouped[iRight].Key == "Table")
                        {
                            var deletedCorrelatedSequence = new CorrelatedSequence();
                            deletedCorrelatedSequence.ComparisonUnitArray1 = leftGrouped[iLeft].ToArray();
                            deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                            deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                            newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                            ++iLeft;
                        }
                        else if (leftGrouped[iLeft].Key == "Table" && rightGrouped[iRight].Key == "Para")
                        {
                            var insertedCorrelatedSequence = new CorrelatedSequence();
                            insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                            insertedCorrelatedSequence.ComparisonUnitArray2 = rightGrouped[iRight].ToArray();
                            insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                            ++iRight;
                        }

                        if (iLeft == leftGrouped.Length && iRight == rightGrouped.Length)
                            return newListOfCorrelatedSequence;

                        // if there is content on the left, but not content on the right
                        if (iRight == rightGrouped.Length)
                        {
                            for (int j = iLeft; j < leftGrouped.Length; j++)
                            {
                                var deletedCorrelatedSequence = new CorrelatedSequence();
                                deletedCorrelatedSequence.ComparisonUnitArray1 = leftGrouped[j].ToArray();
                                deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                                deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                            }
                            return newListOfCorrelatedSequence;
                        }
                        // there is content on the right but not on the left
                        else if (iLeft == leftGrouped.Length)
                        {
                            for (int j = iRight; j < rightGrouped.Length; j++)
                            {
                                var insertedCorrelatedSequence = new CorrelatedSequence();
                                insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                                insertedCorrelatedSequence.ComparisonUnitArray2 = rightGrouped[j].ToArray();
                                insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                            }
                            return newListOfCorrelatedSequence;
                        }
                        // else continue on next round.
                    }
                }

                // If both sides consists of a single table, and if the table contains merged cells, then mark as deleted/inserted
                if (leftTables == 1 && leftLength == 1 &&
                    rightTables == 1 && rightLength == 1)
                {
                    var cug1 = unknown.ComparisonUnitArray1.First() as ComparisonUnitGroup;
                    var firstContentAtom1 = cug1.DescendantContentAtoms().FirstOrDefault();
                    if (firstContentAtom1 == null)
                        throw new OpenXmlPowerToolsException("Internal error");
                    var tbl1 = firstContentAtom1
                        .AncestorElements
                        .Reverse()
                        .FirstOrDefault(a => a.Name == W.tbl);

                    var cug2 = unknown.ComparisonUnitArray1.First() as ComparisonUnitGroup;
                    var firstContentAtom2 = cug2.DescendantContentAtoms().FirstOrDefault();
                    if (firstContentAtom2 == null)
                        throw new OpenXmlPowerToolsException("Internal error");
                    var tbl2 = firstContentAtom2
                        .AncestorElements
                        .Reverse()
                        .FirstOrDefault(a => a.Name == W.tbl);

                    var leftContainsMerged = tbl1
                        .Descendants()
                        .Any(d => d.Name == W.vMerge || d.Name == W.gridSpan);

                    var rightContainsMerged = tbl2
                        .Descendants()
                        .Any(d => d.Name == W.vMerge || d.Name == W.gridSpan);

                    if (leftContainsMerged || rightContainsMerged)
                    {
                        // flatten to rows
                        var deletedCorrelatedSequence = new CorrelatedSequence();
                        deletedCorrelatedSequence.ComparisonUnitArray1 = unknown
                            .ComparisonUnitArray1
                            .Select(z => z.Contents)
                            .SelectMany(m => m)
                            .ToArray();
                        deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                        deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                        newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);

                        var insertedCorrelatedSequence = new CorrelatedSequence();
                        insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                        insertedCorrelatedSequence.ComparisonUnitArray2 = unknown
                            .ComparisonUnitArray2
                            .Select(z => z.Contents)
                            .SelectMany(m => m)
                            .ToArray();
                        insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                        newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);

                        return newListOfCorrelatedSequence;
                    }
                }

                // If either side contains only paras or tables, then flatten and iterate.
                var leftOnlyParasTablesTextboxes = leftLength == leftTables + leftParagraphs + leftTextboxes;
                var rightOnlyParasTablesTextboxes = rightLength == rightTables + rightParagraphs + rightTextboxes;
                if (leftOnlyParasTablesTextboxes && rightOnlyParasTablesTextboxes)
                {
                    // flatten paras and tables, and iterate
                    var left = unknown
                        .ComparisonUnitArray1
                        .Select(cu => cu.Contents)
                        .SelectMany(m => m)
                        .ToArray();

                    var right = unknown
                        .ComparisonUnitArray2
                        .Select(cu => cu.Contents)
                        .SelectMany(m => m)
                        .ToArray();

                    var unknownCorrelatedSequence = new CorrelatedSequence();
                    unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                    unknownCorrelatedSequence.ComparisonUnitArray1 = left;
                    unknownCorrelatedSequence.ComparisonUnitArray2 = right;
                    newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);

                    return newListOfCorrelatedSequence;
                }

                // if first of left is a row and first of right is a row
                // then flatten the row to cells and iterate.

                var firstLeft = unknown
                    .ComparisonUnitArray1
                    .First() as ComparisonUnitGroup;

                var firstRight = unknown
                    .ComparisonUnitArray2
                    .First() as ComparisonUnitGroup;

                if (firstLeft != null && firstRight != null)
                {
                    if (firstLeft.ComparisonUnitGroupType == ComparisonUnitGroupType.Row &&
                        firstRight.ComparisonUnitGroupType == ComparisonUnitGroupType.Row)
                    {
                        ComparisonUnit[] leftContent = firstLeft.Contents.ToArray();
                        ComparisonUnit[] rightContent = firstRight.Contents.ToArray();

                        var lenLeft = leftContent.Length;
                        var lenRight = rightContent.Length;

                        if (lenLeft < lenRight)
                            leftContent = leftContent.Concat(Enumerable.Repeat<ComparisonUnit>(null, lenRight - lenLeft)).ToArray();
                        else if (lenRight < lenLeft)
                            rightContent = rightContent.Concat(Enumerable.Repeat<ComparisonUnit>(null, lenLeft - lenRight)).ToArray();

                        List<CorrelatedSequence> newCs = leftContent.Zip(rightContent, (l, r) =>
                            {
                                if (l != null && r != null)
                                {
                                    var cellLcs = Lcs(l.Contents.ToArray(), r.Contents.ToArray(), settings);
                                    return cellLcs.ToArray();
                                }
                                if (l == null)
                                {
                                    var insertedCorrelatedSequence = new CorrelatedSequence();
                                    insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                                    insertedCorrelatedSequence.ComparisonUnitArray2 = r.Contents.ToArray();
                                    insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                                    return new[] { insertedCorrelatedSequence };
                                }
                                else if (r == null)
                                {
                                    var deletedCorrelatedSequence = new CorrelatedSequence();
                                    deletedCorrelatedSequence.ComparisonUnitArray1 = l.Contents.ToArray();
                                    deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                                    deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                                    return new[] { deletedCorrelatedSequence };
                                }
                                else
                                    throw new OpenXmlPowerToolsException("Internal error");
                            })
                            .SelectMany(m => m)
                            .ToList();

                        foreach (var cs in newCs)
                            newListOfCorrelatedSequence.Add(cs);

                        var remainderLeft = unknown
                            .ComparisonUnitArray1
                            .Skip(1)
                            .ToArray();

                        var remainderRight = unknown
                            .ComparisonUnitArray2
                            .Skip(1)
                            .ToArray();

                        if (remainderLeft.Length > 0 && remainderRight.Length == 0)
                        {
                            var deletedCorrelatedSequence = new CorrelatedSequence();
                            deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                            deletedCorrelatedSequence.ComparisonUnitArray1 = remainderLeft;
                            deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                            newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                        }
                        else if (remainderRight.Length > 0 && remainderLeft.Length == 0)
                        {
                            var insertedCorrelatedSequence = new CorrelatedSequence();
                            insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                            insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                            insertedCorrelatedSequence.ComparisonUnitArray2 = remainderRight;
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                        }
                        else if (remainderLeft.Length > 0 && remainderRight.Length > 0)
                        {
                            var unknownCorrelatedSequence2 = new CorrelatedSequence();
                            unknownCorrelatedSequence2.CorrelationStatus = CorrelationStatus.Unknown;
                            unknownCorrelatedSequence2.ComparisonUnitArray1 = remainderLeft;
                            unknownCorrelatedSequence2.ComparisonUnitArray2 = remainderRight;
                            newListOfCorrelatedSequence.Add(unknownCorrelatedSequence2);
                        }

                        if (s_False)
                        {
                            var sb = new StringBuilder();
                            foreach (var item in newListOfCorrelatedSequence)
                                sb.Append(item.ToString()).Append(Environment.NewLine);
                            var sbs = sb.ToString();
                            TestUtil.NotePad(sbs);
                        }

                        return newListOfCorrelatedSequence;
                    }
                    if (firstLeft.ComparisonUnitGroupType == ComparisonUnitGroupType.Cell &&
                        firstRight.ComparisonUnitGroupType == ComparisonUnitGroupType.Cell)
                    {
                        var left = firstLeft
                            .Contents
                            .ToArray();

                        var right = firstRight
                            .Contents
                            .ToArray();

                        var unknownCorrelatedSequence = new CorrelatedSequence();
                        unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                        unknownCorrelatedSequence.ComparisonUnitArray1 = left;
                        unknownCorrelatedSequence.ComparisonUnitArray2 = right;
                        newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);

                        var remainderLeft = unknown
                            .ComparisonUnitArray1
                            .Skip(1)
                            .ToArray();

                        var remainderRight = unknown
                            .ComparisonUnitArray2
                            .Skip(1)
                            .ToArray();

                        if (remainderLeft.Length > 0 && remainderRight.Length == 0)
                        {
                            var deletedCorrelatedSequence = new CorrelatedSequence();
                            deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                            deletedCorrelatedSequence.ComparisonUnitArray1 = remainderLeft;
                            deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                            newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                        }
                        else if (remainderRight.Length > 0 && remainderLeft.Length == 0)
                        {
                            var insertedCorrelatedSequence = new CorrelatedSequence();
                            insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                            insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                            insertedCorrelatedSequence.ComparisonUnitArray2 = remainderRight;
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                        }
                        else if (remainderLeft.Length > 0 && remainderRight.Length > 0)
                        {
                            var unknownCorrelatedSequence2 = new CorrelatedSequence();
                            unknownCorrelatedSequence2.CorrelationStatus = CorrelationStatus.Unknown;
                            unknownCorrelatedSequence2.ComparisonUnitArray1 = remainderLeft;
                            unknownCorrelatedSequence2.ComparisonUnitArray2 = remainderRight;
                            newListOfCorrelatedSequence.Add(unknownCorrelatedSequence2);
                        }

                        return newListOfCorrelatedSequence;
                    }
                }

                // otherwise create ins and del

                var deletedCorrelatedSequence3 = new CorrelatedSequence();
                deletedCorrelatedSequence3.CorrelationStatus = CorrelationStatus.Deleted;
                deletedCorrelatedSequence3.ComparisonUnitArray1 = unknown.ComparisonUnitArray1;
                deletedCorrelatedSequence3.ComparisonUnitArray2 = null;
                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence3);

                var insertedCorrelatedSequence3 = new CorrelatedSequence();
                insertedCorrelatedSequence3.CorrelationStatus = CorrelationStatus.Inserted;
                insertedCorrelatedSequence3.ComparisonUnitArray1 = null;
                insertedCorrelatedSequence3.ComparisonUnitArray2 = unknown.ComparisonUnitArray2;
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence3);

                return newListOfCorrelatedSequence;
            }

            if (currentI1 > 0 && currentI2 == 0)
            {
                var deletedCorrelatedSequence = new CorrelatedSequence();
                deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                deletedCorrelatedSequence.ComparisonUnitArray1 = cul1
                    .Take(currentI1)
                    .ToArray();
                deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
            }
            else if (currentI1 == 0 && currentI2 > 0)
            {
                var insertedCorrelatedSequence = new CorrelatedSequence();
                insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                insertedCorrelatedSequence.ComparisonUnitArray2 = cul2
                    .Take(currentI2)
                    .ToArray();
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
            }
            else if (currentI1 > 0 && currentI2 > 0)
            {
                var unknownCorrelatedSequence = new CorrelatedSequence();
                unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                unknownCorrelatedSequence.ComparisonUnitArray1 = cul1
                    .Take(currentI1)
                    .ToArray();
                unknownCorrelatedSequence.ComparisonUnitArray2 = cul2
                    .Take(currentI2)
                    .ToArray();
                newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
            }
            else if (currentI1 == 0 && currentI2 == 0)
            {
                // nothing to do
            }

            var middleEqual = new CorrelatedSequence();
            middleEqual.CorrelationStatus = CorrelationStatus.Equal;
            middleEqual.ComparisonUnitArray1 = cul1
                .Skip(currentI1)
                .Take(currentLongestCommonSequenceLength)
                .ToArray();
            middleEqual.ComparisonUnitArray2 = cul2
                .Skip(currentI2)
                .Take(currentLongestCommonSequenceLength)
                .ToArray();
            newListOfCorrelatedSequence.Add(middleEqual);

            int endI1 = currentI1 + currentLongestCommonSequenceLength;
            int endI2 = currentI2 + currentLongestCommonSequenceLength;

            if (endI1 < cul1.Length && endI2 == cul2.Length)
            {
                var deletedCorrelatedSequence = new CorrelatedSequence();
                deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                deletedCorrelatedSequence.ComparisonUnitArray1 = cul1
                    .Skip(endI1)
                    .ToArray();
                deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
            }
            else if (endI1 == cul1.Length && endI2 < cul2.Length)
            {
                var insertedCorrelatedSequence = new CorrelatedSequence();
                insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                insertedCorrelatedSequence.ComparisonUnitArray2 = cul2
                    .Skip(endI2)
                    .ToArray();
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
            }
            else if (endI1 < cul1.Length && endI2 < cul2.Length)
            {
                var unknownCorrelatedSequence = new CorrelatedSequence();
                unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                unknownCorrelatedSequence.ComparisonUnitArray1 = cul1
                    .Skip(endI1)
                    .ToArray();
                unknownCorrelatedSequence.ComparisonUnitArray2 = cul2
                    .Skip(endI2)
                    .ToArray();
                newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
            }
            else if (endI1 == cul1.Length && endI2 == cul2.Length)
            {
                // nothing to do
            }
            return newListOfCorrelatedSequence;
        }

        private static XName[] WordBreakElements = new XName[] {
            W.pPr,
            W.tab,
            W.br,
            W.continuationSeparator,
            W.cr,
            W.dayLong,
            W.dayShort,
            W.drawing,
            W.pict,
            W.endnoteRef,
            W.footnoteRef,
            W.monthLong,
            W.monthShort,
            W.noBreakHyphen,
            W._object,
            W.ptab,
            W.separator,
            W.sym,
            W.yearLong,
            W.yearShort,
            M.oMathPara,
            M.oMath,
            W.footnoteReference,
            W.endnoteReference,
        };

        private class Atgbw
        {
            public int? Key;
            public ComparisonUnitAtom ComparisonUnitAtomMember;
            public int NextIndex;
        }

        private static ComparisonUnit[] GetComparisonUnitList(ComparisonUnitAtom[] comparisonUnitAtomList, WmlComparerSettings settings)
        {
            var seed = new Atgbw()
            {
                Key = null,
                ComparisonUnitAtomMember = null,
                NextIndex = 0,
            };

            var groupingKey = comparisonUnitAtomList
                .Rollup(seed, (sr, prevAtgbw, i) =>
                {
                    int? key = null;
                    var nextIndex = prevAtgbw.NextIndex;
                    if (sr.ContentElement.Name == W.t)
                    {
                        string chr = sr.ContentElement.Value;
                        var ch = chr[0];
                        if (ch == '.' || ch == ',')
                        {
                            bool beforeIsDigit = false;
                            if (i > 0)
                            {
                                var prev = comparisonUnitAtomList[i - 1];
                                if (prev.ContentElement.Name == W.t && char.IsDigit(prev.ContentElement.Value[0]))
                                    beforeIsDigit = true;
                            }
                            bool afterIsDigit = false;
                            if (i < comparisonUnitAtomList.Length - 1)
                            {
                                var next = comparisonUnitAtomList[i + 1];
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
                    return new Atgbw()
                    {
                        Key = key,
                        ComparisonUnitAtomMember = sr,
                        NextIndex = nextIndex,
                    };
                });

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in groupingKey)
                {
                    sb.Append(item.Key + Environment.NewLine);
                    sb.Append("    " + item.ComparisonUnitAtomMember.ToString(0) + Environment.NewLine);                                                                                                                                                                                                                    
                }
                var sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            var groupedByWords = groupingKey
                .GroupAdjacent(gc => gc.Key);

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var group in groupedByWords)
                {
                    sb.Append("Group ===== " + group.Key + Environment.NewLine);
                    foreach (var gc in group)
                    {
                        sb.Append("    " + gc.ComparisonUnitAtomMember.ToString(0) + Environment.NewLine);
                    }
                }
                var sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

             var withHierarchicalGroupingKey = groupedByWords
                .Select(g =>
                    {
                        var hierarchicalGroupingArray = g
                            .First()
                            .ComparisonUnitAtomMember
                            .AncestorElements
                            .Where(a => ComparisonGroupingElements.Contains(a.Name))
                            .Select(a => a.Name.LocalName + ":" + (string)a.Attribute(PtOpenXml.Unid))
                            .ToArray();

                        return new WithHierarchicalGroupingKey() {
                            ComparisonUnitWord = new ComparisonUnitWord(g.Select(gc => gc.ComparisonUnitAtomMember)),
                            HierarchicalGroupingArray = hierarchicalGroupingArray,
                        };
                    }
                )
                .ToArray();

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var group in withHierarchicalGroupingKey)
                {
                    sb.Append("Grouping Array: " + group.HierarchicalGroupingArray.Select(gam => gam + " - ").StringConcatenate() + Environment.NewLine);
                    foreach (var gc in group.ComparisonUnitWord.Contents)
                    {
                        sb.Append("    " + gc.ToString(0) + Environment.NewLine);
                    }
                }
                var sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            var cul = GetHierarchicalComparisonUnits(withHierarchicalGroupingKey, 0).ToArray();

            if (s_False)
            {
                var str = ComparisonUnit.ComparisonUnitListToString(cul);
                TestUtil.NotePad(str);
            }

            return cul;
        }

        private static IEnumerable<ComparisonUnit> GetHierarchicalComparisonUnits(IEnumerable<WithHierarchicalGroupingKey> input, int level)
        {
            var grouped = input
                .GroupAdjacent(whgk =>
                {
                    if (level >= whgk.HierarchicalGroupingArray.Length)
                        return "";
                    return whgk.HierarchicalGroupingArray[level];
                });
            var retList = grouped
                .Select(gc =>
                {
                    if (gc.Key == "")
                    {
                        return (IEnumerable<ComparisonUnit>)gc.Select(whgk => whgk.ComparisonUnitWord).ToList();
                    }
                    else
                    {
                        ComparisonUnitGroupType? group = null;
                        var spl = gc.Key.Split(':');
                        if (spl[0] == "p")
                            group = ComparisonUnitGroupType.Paragraph;
                        else if (spl[0] == "tbl")
                            group = ComparisonUnitGroupType.Table;
                        else if (spl[0] == "tr")
                            group = ComparisonUnitGroupType.Row;
                        else if (spl[0] == "tc")
                            group = ComparisonUnitGroupType.Cell;
                        else if (spl[0] == "txbxContent")
                            group = ComparisonUnitGroupType.Textbox;
                        var newCompUnitGroup = new ComparisonUnitGroup(GetHierarchicalComparisonUnits(gc, level + 1), (ComparisonUnitGroupType)group);
                        return new[] { newCompUnitGroup };
                    }
                })
                .SelectMany(m => m)
                .ToList();
            return retList;
        }

        private static XName[] AllowableRunChildren = new XName[] {
            W.br,
            W.drawing,
            W.cr,
            W.dayLong,
            W.dayShort,
            W.footnoteReference,
            W.endnoteReference,
            W.monthLong,
            W.monthShort,
            W.noBreakHyphen,
            //W._object,
            W.pgNum,
            W.ptab,
            W.softHyphen,
            W.sym,
            W.tab,
            W.yearLong,
            W.yearShort,
            M.oMathPara,
            M.oMath,
            W.fldChar,
            W.instrText,
        };

        private static XName[] ElementsToThrowAway = new XName[] {
            W.bookmarkStart,
            W.bookmarkEnd,
            W.commentRangeStart,
            W.commentRangeEnd,
            W.lastRenderedPageBreak,
            W.proofErr,
            W.tblPr,
            W.sectPr,
            W.permEnd,
            W.permStart,
            W.footnoteRef,
            W.endnoteRef,
            W.separator,
            W.continuationSeparator,
        };

        private static XName[] ElementsToHaveSha1Hash = new XName[]
        {
            W.p,
            W.tbl,
            W.tr,
            W.tc,
            W.drawing,
            W.pict,
            W.txbxContent,
        };

        private static XName[] InvalidElements = new XName[]
        {
            W.altChunk,
            W.customXml,
            W.customXmlDelRangeEnd,
            W.customXmlDelRangeStart,
            W.customXmlInsRangeEnd,
            W.customXmlInsRangeStart,
            W.customXmlMoveFromRangeEnd,
            W.customXmlMoveFromRangeStart,
            W.customXmlMoveToRangeEnd,
            W.customXmlMoveToRangeStart,
            W.moveFrom,
            W.moveFromRangeStart,
            W.moveFromRangeEnd,
            W.moveTo,
            W.moveToRangeStart,
            W.moveToRangeEnd,
            W.subDoc,
        };

        private class RecursionInfo
        {
            public XName ElementName;
            public XName[] ChildElementPropertyNames;
        }

        private static RecursionInfo[] RecursionElements = new RecursionInfo[]
        {
            new RecursionInfo()
            {
                ElementName = W.del,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = W.ins,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = W.tbl,
                ChildElementPropertyNames = new[] { W.tblPr, W.tblGrid, W.tblPrEx },
            },
            new RecursionInfo()
            {
                ElementName = W.tr,
                ChildElementPropertyNames = new[] { W.trPr, W.tblPrEx },
            },
            new RecursionInfo()
            {
                ElementName = W.tc,
                ChildElementPropertyNames = new[] { W.tcPr, W.tblPrEx },
            },
            new RecursionInfo()
            {
                ElementName = W.pict,
                ChildElementPropertyNames = new[] { VML.shapetype },
            },
            new RecursionInfo()
            {
                ElementName = VML.shape,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = VML.textbox,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = W.txbxContent,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = W10.wrap,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = W.sdt,
                ChildElementPropertyNames = new[] { W.sdtPr, W.sdtEndPr },
            },
            new RecursionInfo()
            {
                ElementName = W.sdtContent,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = W.hyperlink,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = W.fldSimple,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = W.smartTag,
                ChildElementPropertyNames = new[] { W.smartTagPr },
            },
        };

        internal static ComparisonUnitAtom[] CreateComparisonUnitAtomList(OpenXmlPart part, XElement contentParent)
        {
            VerifyNoInvalidContent(contentParent);
            AssignUnidToAllElements(contentParent);  // add the Guid id to every element
            MoveLastSectPrIntoLastParagraph(contentParent);
            var cal = CreateComparisonUnitAtomListInternal(part, contentParent).ToArray();

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in cal)
                    sb.Append(item.ToString() + Environment.NewLine);
                var sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            return cal;
        }

        private static void VerifyNoInvalidContent(XElement contentParent)
        {
            var invalidElement = contentParent.Descendants().FirstOrDefault(d => InvalidElements.Contains(d.Name));
            if (invalidElement == null)
                return;
            throw new NotSupportedException("Document contains " + invalidElement.Name.LocalName);
        }

        internal static XDocument Coalesce(ComparisonUnitAtom[] comparisonUnitAtomList)
        {
            XDocument newXDoc = new XDocument();
            var newBodyChildren = CoalesceRecurse(comparisonUnitAtomList, 0);
            newXDoc.Add(new XElement(W.document,
                new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "pt14", PtOpenXml.pt.NamespaceName),
                new XElement(W.body, newBodyChildren)));

            // little bit of cleanup
            MoveLastSectPrToChildOfBody(newXDoc);
            XElement newXDoc2Root = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(newXDoc.Root);
            newXDoc.Root.ReplaceWith(newXDoc2Root);
            return newXDoc;
        }

        private static object CoalesceRecurse(IEnumerable<ComparisonUnitAtom> list, int level)
        {
            var grouped = list
                .GroupBy(sr =>
                {
                    // per the algorithm, The following condition will never evaluate to true
                    // if it evaluates to true, then the basic mechanism for breaking a hierarchical structure into flat and back is broken.

                    // for a table, we initially get all ComparisonUnitAtoms for the entire table, then process.  When processing a row,
                    // no ComparisonUnitAtoms will have ancestors outside the row.  Ditto for cells, and on down the tree.
                    if (level >= sr.AncestorElements.Length)
                        throw new OpenXmlPowerToolsException("Internal error 4 - why do we have ComparisonUnitAtom objects with fewer ancestors than its siblings?");

                    var unid = (string)sr.AncestorElements[level].Attribute(PtOpenXml.Unid);
                    return unid;
                });

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var group in grouped)
                {
                    sb.AppendFormat("Group Key: {0}", group.Key);
                    sb.Append(Environment.NewLine);
                    foreach (var groupChildItem in group)
                    {
                        sb.Append("  ");
                        sb.Append(groupChildItem.ToString(0));
                        sb.Append(Environment.NewLine);
                    }
                    sb.Append(Environment.NewLine);
                }
                var sbs = sb.ToString();
            }

            var elementList = grouped
                .Select(g =>
                {
                    // see the comment above at the beginning of CoalesceRecurse
                    if (level >= g.First().AncestorElements.Length)
                        throw new OpenXmlPowerToolsException("Internal error 3 - why do we have ComparisonUnitAtom objects with fewer ancestors than its siblings?");

                    var ancestorBeingConstructed = g.First().AncestorElements[level];

                    if (ancestorBeingConstructed.Name == W.p)
                    {
                        var groupedChildren = g
                            .GroupAdjacent(gc => gc.ContentElement.Name.ToString());
                        var newChildElements = groupedChildren
                            .Where(gc => gc.First().ContentElement.Name != W.pPr)
                            .Select(gc =>
                            {
                                return CoalesceRecurse(gc, level + 1);
                            });
                        var newParaProps = groupedChildren
                            .Where(gc => gc.First().ContentElement.Name == W.pPr)
                            .Select(gc => gc.Select(gce => gce.ContentElement));
                        return new XElement(W.p,
                            ancestorBeingConstructed.Attributes(),
                            newParaProps, newChildElements);
                    }

                    if (ancestorBeingConstructed.Name == W.r)
                    {
                        var groupedChildren = g
                            .GroupAdjacent(gc => gc.ContentElement.Name.ToString());
                        var newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                var name = gc.First().ContentElement.Name;
                                if (name == W.t || name == W.delText)
                                {
                                    var textOfTextElement = gc.Select(gce => gce.ContentElement.Value).StringConcatenate();
                                    return (object)(new XElement(name,
                                        GetXmlSpaceAttribute(textOfTextElement),
                                        textOfTextElement));
                                }
                                else
                                    return gc.Select(gce => gce.ContentElement);
                            });
                        var runProps = ancestorBeingConstructed.Elements(W.rPr);
                        return new XElement(W.r, runProps, newChildElements);
                    }

                    var re = RecursionElements.FirstOrDefault(z => z.ElementName == ancestorBeingConstructed.Name);
                    if (re != null)
                    {
                        return ReconstructElement(g, ancestorBeingConstructed, re.ChildElementPropertyNames, level);
                    }

                    var newElement = new XElement(ancestorBeingConstructed.Name,
                        ancestorBeingConstructed.Attributes(),
                        CoalesceRecurse(g, level + 1));
                    return newElement;
                })
                .ToList();
            return elementList;
        }

        private static XElement ReconstructElement(IGrouping<string, ComparisonUnitAtom> g, XElement ancestorBeingConstructed, XName[] childPropElementNames, int level)
        {
            var newChildElements = CoalesceRecurse(g, level + 1);
            IEnumerable<XElement> childProps = null;
            if (childPropElementNames != null)
                childProps = ancestorBeingConstructed.Elements()
                    .Where(a => childPropElementNames.Contains(a.Name));

            var reconstructedElement = new XElement(ancestorBeingConstructed.Name, childProps, newChildElements);
            return reconstructedElement;
        }

        private static void MoveLastSectPrIntoLastParagraph(XElement contentParent)
        {
            var lastSectPrList = contentParent.Elements(W.sectPr).ToList();
            if (lastSectPrList.Count() > 1)
                throw new OpenXmlPowerToolsException("Invalid document");
            var lastSectPr = lastSectPrList.FirstOrDefault();
            if (lastSectPr != null)
            {
                var lastParagraph = contentParent.Elements(W.p).LastOrDefault();
                if (lastParagraph == null)
                    throw new OpenXmlPowerToolsException("Invalid document");
                var pPr = lastParagraph.Element(W.pPr);
                if (pPr == null)
                {
                    pPr = new XElement(W.pPr);
                    lastParagraph.AddFirst(W.pPr);
                }
                pPr.Add(lastSectPr);
                contentParent.Elements(W.sectPr).Remove();
            }
        }

        private static List<ComparisonUnitAtom> CreateComparisonUnitAtomListInternal(OpenXmlPart part, XElement contentParent)
        {
            var comparisonUnitAtomList = new List<ComparisonUnitAtom>();
            CreateComparisonUnitAtomListRecurse(part, contentParent, comparisonUnitAtomList);
            return comparisonUnitAtomList;
        }

        private static XName[] ComparisonGroupingElements = new[] {
            W.p,
            W.tbl,
            W.tr,
            W.tc,
            W.txbxContent,
        };

        private static void CreateComparisonUnitAtomListRecurse(OpenXmlPart part, XElement element, List<ComparisonUnitAtom> comparisonUnitAtomList)
        {
            if (element.Name == W.body || element.Name == W.footnote || element.Name == W.endnote)
            {
                foreach (var item in element.Elements())
                    CreateComparisonUnitAtomListRecurse(part, item, comparisonUnitAtomList);
                return;
            }

            if (element.Name == W.p)
            {
                var paraChildrenToProcess = element
                    .Elements()
                    .Where(e => e.Name != W.pPr);
                foreach (var item in paraChildrenToProcess)
                    CreateComparisonUnitAtomListRecurse(part, item, comparisonUnitAtomList);
                var paraProps = element.Element(W.pPr);
                if (paraProps == null)
                {
                    ComparisonUnitAtom pPrComparisonUnitAtom = new ComparisonUnitAtom(
                        new XElement(W.pPr),
                        element.AncestorsAndSelf().TakeWhile(a => a.Name != W.body && a.Name != W.footnotes && a.Name != W.endnotes).Reverse().ToArray(),
                        part);
                    comparisonUnitAtomList.Add(pPrComparisonUnitAtom);
                }
                else
                {
                    ComparisonUnitAtom pPrComparisonUnitAtom = new ComparisonUnitAtom(
                        paraProps,
                        element.AncestorsAndSelf().TakeWhile(a => a.Name != W.body && a.Name != W.footnotes && a.Name != W.endnotes).Reverse().ToArray(),
                        part);
                    comparisonUnitAtomList.Add(pPrComparisonUnitAtom);
                }
                return;
            }

            if (element.Name == W.r)
            {
                var runChildrenToProcess = element
                    .Elements()
                    .Where(e => e.Name != W.rPr);
                foreach (var item in runChildrenToProcess)
                    CreateComparisonUnitAtomListRecurse(part, item, comparisonUnitAtomList);
                return;
            }

            if (element.Name == W.t || element.Name == W.delText)
            {
                var val = element.Value;
                foreach (var ch in val)
                {
                    ComparisonUnitAtom sr = new ComparisonUnitAtom(
                        new XElement(element.Name, ch),
                        element.AncestorsAndSelf().TakeWhile(a => a.Name != W.body && a.Name != W.footnotes && a.Name != W.endnotes).Reverse().ToArray(),
                        part);
                    comparisonUnitAtomList.Add(sr);
                }
                return;
            }

            if (AllowableRunChildren.Contains(element.Name) || element.Name == W._object)
            {
                ComparisonUnitAtom sr3 = new ComparisonUnitAtom(
                    element,
                    element.AncestorsAndSelf().TakeWhile(a => a.Name != W.body).Reverse().ToArray(),
                    part);
                comparisonUnitAtomList.Add(sr3);
                return;
            }

            var re = RecursionElements.FirstOrDefault(z => z.ElementName == element.Name);
            if (re != null)
            {
                AnnotateElementWithProps(part, element, comparisonUnitAtomList, re.ChildElementPropertyNames);
                return;
            }

            if (ElementsToThrowAway.Contains(element.Name))
                return;

            throw new OpenXmlPowerToolsException("Internal error - unexpected element");
        }

        private static void AnnotateElementWithProps(OpenXmlPart part, XElement element, List<ComparisonUnitAtom> comparisonUnitAtomList, XName[] childElementPropertyNames)
        {
            IEnumerable<XElement> runChildrenToProcess = null;
            if (childElementPropertyNames == null)
                runChildrenToProcess = element.Elements();
            else
                runChildrenToProcess = element
                    .Elements()
                    .Where(e => !childElementPropertyNames.Contains(e.Name));

            foreach (var item in runChildrenToProcess)
                CreateComparisonUnitAtomListRecurse(part, item, comparisonUnitAtomList);
        }

        private static void AssignUnidToAllElements(XElement contentParent)
        {
            if (contentParent.Descendants().Attributes(PtOpenXml.Unid).Any())
                return;
            var content = contentParent.Descendants();
            foreach (var d in content)
            {
                string unid = Guid.NewGuid().ToString().Replace("-", "");
                var newAtt = new XAttribute(PtOpenXml.Unid, unid);
                var newAtt2 = new XAttribute(PtOpenXml.PrevUnid, unid);
                d.Add(newAtt, newAtt2);
            }
        }
    }

    internal class WithHierarchicalGroupingKey
    {
        public string[] HierarchicalGroupingArray;
        public ComparisonUnitWord ComparisonUnitWord;
    }

    public abstract class ComparisonUnit
    {
        public List<ComparisonUnit> Contents;
        public string SHA1Hash;
        public CorrelationStatus CorrelationStatus;

        public IEnumerable<ComparisonUnit> Descendants()
        {
            List<ComparisonUnit> comparisonUnitList = new List<ComparisonUnit>();
            DescendantsInternal(this, comparisonUnitList);
            return comparisonUnitList;
        }

        public IEnumerable<ComparisonUnitAtom> DescendantContentAtoms()
        {
            return Descendants().OfType<ComparisonUnitAtom>();
        }

        private void DescendantsInternal(ComparisonUnit comparisonUnit, List<ComparisonUnit> comparisonUnitList)
        {
            foreach (var cu in comparisonUnit.Contents)
            {
                comparisonUnitList.Add(cu);
                if (cu.Contents != null && cu.Contents.Any())
                    DescendantsInternal(cu, comparisonUnitList);
            }
        }

        public abstract string ToString(int indent);

        internal static string ComparisonUnitListToString(ComparisonUnit[] cul)
        {
            var sb = new StringBuilder();
            sb.Append("Dump Comparision Unit List To String" + Environment.NewLine);
            foreach (var item in cul)
            {
                sb.Append(item.ToString(2) + Environment.NewLine);
            }
            return sb.ToString();
        }
    }

    internal class ComparisonUnitWord : ComparisonUnit
    {
        public ComparisonUnitWord(IEnumerable<ComparisonUnitAtom> comparisonUnitAtomList)
        {
            Contents = comparisonUnitAtomList.OfType<ComparisonUnit>().ToList();
            var sha1String = Contents
                .Select(c => c.SHA1Hash)
                .StringConcatenate();
            SHA1Hash = WmlComparerUtil.SHA1HashStringForUTF8String(sha1String);
        }

        public static XName[] s_ElementsWithRelationshipIds = new XName[] {
            A.blip,
            A.hlinkClick,
            A.relIds,
            C.chart,
            C.externalData,
            C.userShapes,
            DGM.relIds,
            O.OLEObject,
            VML.fill,
            VML.imagedata,
            VML.stroke,
            W.altChunk,
            W.attachedTemplate,
            W.control,
            W.dataSource,
            W.embedBold,
            W.embedBoldItalic,
            W.embedItalic,
            W.embedRegular,
            W.footerReference,
            W.headerReference,
            W.headerSource,
            W.hyperlink,
            W.printerSettings,
            W.recipientData,
            W.saveThroughXslt,
            W.sourceFileName,
            W.src,
            W.subDoc,
            WNE.toolbarData,
        };

        public static XName[] s_RelationshipAttributeNames = new XName[] {
            R.embed,
            R.link,
            R.id,
            R.cs,
            R.dm,
            R.lo,
            R.qs,
            R.href,
            R.pict,
        };

        public override string ToString(int indent)
        {
            var sb = new StringBuilder();
            sb.Append("".PadRight(indent) + "Word SHA1:" + this.SHA1Hash.Substring(0, 8) + Environment.NewLine);
            foreach (var comparisonUnitAtom in Contents)
                sb.Append(comparisonUnitAtom.ToString(indent + 2) + Environment.NewLine);
            return sb.ToString();
        }
    }

    class WmlComparerUtil
    {
        public static string SHA1HashStringForUTF8String(string s)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(s);
            var sha1 = SHA1.Create();
            byte[] hashBytes = sha1.ComputeHash(bytes);
            return HexStringFromBytes(hashBytes);
        }

        public static string SHA1HashStringForByteArray(byte[] bytes)
        {
            var sha1 = SHA1.Create();
            byte[] hashBytes = sha1.ComputeHash(bytes);
            return HexStringFromBytes(hashBytes);
        }

        public static string HexStringFromBytes(byte[] bytes)
        {
            var sb = new StringBuilder();
            foreach (byte b in bytes)
            {
                var hex = b.ToString("x2");
                sb.Append(hex);
            }
            return sb.ToString();
        }
    }

    public class ComparisonUnitAtom : ComparisonUnit
    {
        // AncestorElements are kept in order from the body to the leaf, because this is the order in which we need to access in order
        // to reassemble the document.  However, in many places in the code, it is necessary to find the nearest ancestor, i.e. cell
        // so it is necessary to reverse the order when looking for it, i.e. look from the leaf back to the body element.

        public XElement[] AncestorElements;
        public XElement ContentElement;
        public XElement ContentElementBefore;
        public OpenXmlPart Part;
        public XElement RevTrackElement;

        public ComparisonUnitAtom(XElement contentElement, XElement[] ancestorElements, OpenXmlPart part)
        {
            ContentElement = contentElement;
            AncestorElements = ancestorElements;
            Part = part;
            RevTrackElement = GetRevisionTrackingElementFromAncestors(contentElement, AncestorElements);
            if (RevTrackElement == null)
            {
                CorrelationStatus = CorrelationStatus.Equal;
            }
            else
            {
                if (RevTrackElement.Name == W.del)
                    CorrelationStatus = CorrelationStatus.Deleted;
                else if (RevTrackElement.Name == W.ins)
                    CorrelationStatus = CorrelationStatus.Inserted;
            }
            string sha1Hash = (string)contentElement.Attribute(PtOpenXml.SHA1Hash);
            if (sha1Hash != null)
            {
                SHA1Hash = sha1Hash;
            }
            else
            {
                var shaHashString = GetSha1HashStringForElement(ContentElement);
                SHA1Hash = WmlComparerUtil.SHA1HashStringForUTF8String(shaHashString);
            }
        }

        private string GetSha1HashStringForElement(XElement contentElement)
        {
            return contentElement.Name.LocalName + contentElement.Value;
        }

        private static XElement GetRevisionTrackingElementFromAncestors(XElement contentElement, XElement[] ancestors)
        {
            XElement revTrackElement = null;

            if (contentElement.Name == W.pPr)
            {
                revTrackElement = contentElement
                    .Elements(W.rPr)
                    .Elements()
                    .FirstOrDefault(e => e.Name == W.del || e.Name == W.ins);
                return revTrackElement;
            }

            revTrackElement = ancestors.FirstOrDefault(a => a.Name == W.del || a.Name == W.ins);
            return revTrackElement;
        }
        
        public override string ToString(int indent)
        {
            int xNamePad = 16;
            var indentString = "".PadRight(indent);

            var sb = new StringBuilder();
            sb.Append(indentString);
            string correlationStatus = "";
            if (CorrelationStatus != OpenXmlPowerTools.CorrelationStatus.Nil)
                correlationStatus = string.Format("[{0}] ", CorrelationStatus.ToString().PadRight(8));
            if (ContentElement.Name == W.t || ContentElement.Name == W.delText)
            {
                sb.AppendFormat("Atom {0}: {1} {2} SHA1:{3} ", PadLocalName(xNamePad, this), ContentElement.Value, correlationStatus, this.SHA1Hash.Substring(0, 8));
                AppendAncestorsDump(sb, this);
            }
            else
            {
                sb.AppendFormat("Atom {0}:   {1} SHA1:{2} ", PadLocalName(xNamePad, this), correlationStatus, this.SHA1Hash.Substring(0, 8));
                AppendAncestorsDump(sb, this);
            }
            return sb.ToString();
        }

        public override string ToString()
        {
            return ToString(0);
        }

        private static string PadLocalName(int xNamePad, ComparisonUnitAtom item)
        {
            return (item.ContentElement.Name.LocalName + " ").PadRight(xNamePad, '-') + " ";
        }

        private void AppendAncestorsDump(StringBuilder sb, ComparisonUnitAtom sr)
        {
            var s = sr.AncestorElements.Select(p => p.Name.LocalName + GetUnid(p) + "/").StringConcatenate().TrimEnd('/');
            sb.Append("Ancestors:" + s);
        }

        private string GetUnid(XElement p)
        {
            var unid = (string)p.Attribute(PtOpenXml.Unid);
            if (unid == null)
                return "";
            return "[" + unid.Substring(0, 8) + "]";
        }

        public static string ComparisonUnitAtomListToString(List<ComparisonUnitAtom> comparisonUnitAtomList, int indent)
        {
            StringBuilder sb = new StringBuilder();
            var cal = comparisonUnitAtomList
                .Select((ca, i) => new
                {
                    ComparisonUnitAtom = ca,
                    Index = i,
                });
            foreach (var item in cal)
                sb.Append("".PadRight(indent))
                  .AppendFormat("[{0:000000}] ", item.Index + 1)
                  .Append(item.ComparisonUnitAtom.ToString(0) + Environment.NewLine);
            return sb.ToString();
        }
    }

    internal enum ComparisonUnitGroupType
    {
        Paragraph,
        Table,
        Row,
        Cell,
        Textbox,
    };

    internal class ComparisonUnitGroup : ComparisonUnit
    {
        public ComparisonUnitGroupType ComparisonUnitGroupType;

        public ComparisonUnitGroup(IEnumerable<ComparisonUnit> comparisonUnitList, ComparisonUnitGroupType groupType)
        {
            Contents = comparisonUnitList.ToList();
            ComparisonUnitGroupType = groupType;
            var first = comparisonUnitList.First();
            ComparisonUnitAtom comparisonUnitAtom = GetFirstComparisonUnitAtomOfGroup(first);
            XName ancestorName = null;
            if (groupType == OpenXmlPowerTools.ComparisonUnitGroupType.Table)
                ancestorName = W.tbl;
            else if (groupType == OpenXmlPowerTools.ComparisonUnitGroupType.Row)
                ancestorName = W.tr;
            else if (groupType == OpenXmlPowerTools.ComparisonUnitGroupType.Cell)
                ancestorName = W.tc;
            else if (groupType == OpenXmlPowerTools.ComparisonUnitGroupType.Paragraph)
                ancestorName = W.p;
            else if (groupType == OpenXmlPowerTools.ComparisonUnitGroupType.Textbox)
                ancestorName = W.txbxContent;
            var ancestor = comparisonUnitAtom.AncestorElements.Reverse().FirstOrDefault(a => a.Name == ancestorName);
            if (ancestor == null)
                throw new OpenXmlPowerToolsException("Internal error: ComparisonUnitGroup");
            SHA1Hash = (string)ancestor.Attribute(PtOpenXml.SHA1Hash);
        }

        public static ComparisonUnitAtom GetFirstComparisonUnitAtomOfGroup(ComparisonUnit group)
        {
            var thisGroup = group;
            while (true)
            {
                var tg = thisGroup as ComparisonUnitGroup;
                if (tg != null)
                {
                    thisGroup = tg.Contents.First();
                    continue;
                }
                var tw = thisGroup as ComparisonUnitWord;
                if (tw == null)
                    throw new OpenXmlPowerToolsException("Internal error: GetFirstComparisonUnitAtomOfGroup");
                var ca = (ComparisonUnitAtom)tw.Contents.First();
                return ca;
            }
        }

        public override string ToString(int indent)
        {
            var sb = new StringBuilder();
            sb.Append("".PadRight(indent) + "Group Type: " + ComparisonUnitGroupType.ToString() + " SHA1:" + SHA1Hash + Environment.NewLine);
            foreach (var comparisonUnitAtom in Contents)
                sb.Append(comparisonUnitAtom.ToString(indent + 2));
            return sb.ToString();
        }
    }

    public enum CorrelationStatus
    {
        Nil,
        Normal,
        Unknown,
        Inserted,
        Deleted,
        Equal,
        Group,
    }

    class PartSHA1HashAnnotation
    {
        public string Hash;

        public PartSHA1HashAnnotation(string hash)
        {
            Hash = hash;
        }
    }

    class CorrelatedSequence
    {
        public CorrelationStatus CorrelationStatus;

        // if ComparisonUnitList1 == null and ComparisonUnitList2 contains sequence, then inserted content.
        // if ComparisonUnitList2 == null and ComparisonUnitList1 contains sequence, then deleted content.
        // if ComparisonUnitList2 contains sequence and ComparisonUnitList1 contains sequence, then either is Unknown or Equal.
        public ComparisonUnit[] ComparisonUnitArray1;
        public ComparisonUnit[] ComparisonUnitArray2;
#if DEBUG
        public string SourceFile;
        public int SourceLine;
#endif

        public CorrelatedSequence()
        {
#if DEBUG
            SourceFile = new System.Diagnostics.StackTrace(true).GetFrame(1).GetFileName();
            SourceLine = new System.Diagnostics.StackTrace(true).GetFrame(1).GetFileLineNumber();
#endif
        }

        public override string ToString()
        {
            var sb = new StringBuilder();
            var indentString = "  ";
            var indentString4 = "    ";
            sb.Append("CorrelatedSequence =====" + Environment.NewLine);
#if DEBUG
            sb.Append(indentString + "Created at Line: " + SourceLine.ToString() + Environment.NewLine);
#endif
            sb.Append(indentString + "CorrelatedItem =====" + Environment.NewLine);
            sb.Append(indentString4 + "CorrelationStatus: " + CorrelationStatus.ToString() + Environment.NewLine);
            if (CorrelationStatus == OpenXmlPowerTools.CorrelationStatus.Equal)
            {
                sb.Append(indentString4 + "ComparisonUnitList =====" + Environment.NewLine);
                foreach (var item in ComparisonUnitArray2)
                    sb.Append(item.ToString(6) + Environment.NewLine);
            }
            else
            {
                if (ComparisonUnitArray1 != null)
                {
                    sb.Append(indentString4 + "ComparisonUnitList1 =====" + Environment.NewLine);
                    foreach (var item in ComparisonUnitArray1)
                        sb.Append(item.ToString(6) + Environment.NewLine);
                }
                if (ComparisonUnitArray2 != null)
                {
                    sb.Append(indentString4 + "ComparisonUnitList2 =====" + Environment.NewLine);
                    foreach (var item in ComparisonUnitArray2)
                        sb.Append(item.ToString(6) + Environment.NewLine);
                }
            }
            return sb.ToString();
        }
    }
}

