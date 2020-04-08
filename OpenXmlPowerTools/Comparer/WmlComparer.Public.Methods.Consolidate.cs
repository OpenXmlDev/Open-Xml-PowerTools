// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static partial class WmlComparer
    {
        /*****************************************************************************************************************/
        // Consolidate processes footnotes and endnotes in a particular fashion - if the unmodified document has a footnote
        // reference, and a delta has a footnote reference, we end up with two footnotes - one is unmodified, and is referred to
        // from the unmodified content.  The footnote reference in the delta refers to the modified footnote.  This is as it
        // should be.
        /*****************************************************************************************************************/

        public static WmlDocument Consolidate(
            WmlDocument original,
            List<WmlRevisedDocumentInfo> revisedDocumentInfoList,
            WmlComparerSettings settings)
        {
            var consolidateSettings = new WmlComparerConsolidateSettings();
            return Consolidate(original, revisedDocumentInfoList, settings, consolidateSettings);
        }

        public static WmlDocument Consolidate(
            WmlDocument original,
            List<WmlRevisedDocumentInfo> revisedDocumentInfoList,
            WmlComparerSettings settings,
            WmlComparerConsolidateSettings consolidateSettings)
        {
            // pre-process the original, so that it already has unids for all elements
            // then when comparing all documents to the original, each one will have the unid as appropriate
            // for all revision block-level content
            //   set unid to look for
            //   while true
            //     determine where to insert
            //       get the unid for the revision
            //       look it up in the original.  if find it, then insert after that element
            //       if not in the original
            //         look backwards in revised document, set unid to look for, do the loop again
            //       if get to the beginning of the document
            //         insert at beginning of document

            settings.StartingIdForFootnotesEndnotes = 3000;
            var originalWithUnids = PreProcessMarkup(original, settings.StartingIdForFootnotesEndnotes);
            var consolidated = new WmlDocument(originalWithUnids);

            if (SaveIntermediateFilesForDebugging && settings.DebugTempFileDi != null)
            {
                var name1 = "Original-with-Unids.docx";
                var preProcFi1 = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name1));
                originalWithUnids.SaveAs(preProcFi1.FullName);
            }

            var revisedDocumentInfoListCount = revisedDocumentInfoList.Count();

            using (var consolidatedMs = new MemoryStream())
            {
                consolidatedMs.Write(consolidated.DocumentByteArray, 0, consolidated.DocumentByteArray.Length);
                using (var consolidatedWDoc = WordprocessingDocument.Open(consolidatedMs, true))
                {
                    var consolidatedMainDocPart = consolidatedWDoc.MainDocumentPart;
                    var consolidatedMainDocPartXDoc = consolidatedMainDocPart.GetXDocument();
                    var consolidatedMainDocPartRoot = consolidatedMainDocPartXDoc.Root ?? throw new ArgumentException();

                    // save away last sectPr
                    var savedSectPr = consolidatedMainDocPartRoot
                        .Elements(W.body)
                        .Elements(W.sectPr)
                        .LastOrDefault();

                    consolidatedMainDocPartRoot
                        .Elements(W.body)
                        .Elements(W.sectPr)
                        .Remove();

                    var consolidatedByUnid = consolidatedMainDocPartXDoc
                        .Descendants()
                        .Where(d => (d.Name == W.p || d.Name == W.tbl) && d.Attribute(PtOpenXml.Unid) != null)
                        .ToDictionary(d => (string)d.Attribute(PtOpenXml.Unid));

                    var deltaNbr = 1;
                    foreach (var revisedDocumentInfo in revisedDocumentInfoList)
                    {
                        settings.StartingIdForFootnotesEndnotes = deltaNbr * 2000 + 3000;
                        var delta = CompareInternal(originalWithUnids, revisedDocumentInfo.RevisedDocument, settings,
                            false);

                        if (SaveIntermediateFilesForDebugging && settings.DebugTempFileDi != null)
                        {
                            var name1 = string.Format("Delta-{0}.docx", deltaNbr++);
                            var deltaFi = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name1));
                            delta.SaveAs(deltaFi.FullName);
                        }

                        using (var msOriginalWithUnids = new MemoryStream())
                        using (var msDelta = new MemoryStream())
                        {
                            msOriginalWithUnids.Write(
                                originalWithUnids.DocumentByteArray,
                                0,
                                originalWithUnids.DocumentByteArray.Length);

                            msDelta.Write(delta.DocumentByteArray, 0, delta.DocumentByteArray.Length);

                            using (var wDocOriginalWithUnids = WordprocessingDocument.Open(msOriginalWithUnids, true))
                            using (var wDocDelta = WordprocessingDocument.Open(msDelta, true))
                            {
                                var modMainDocPart = wDocDelta.MainDocumentPart;
                                var modMainDocPartXDoc = modMainDocPart.GetXDocument();
                                var blockLevelContentToMove = modMainDocPartXDoc
                                    .Root
                                    .DescendantsTrimmed(d => d.Name == W.txbxContent || d.Name == W.tr)
                                    .Where(d => d.Name == W.p || d.Name == W.tbl)
                                    .Where(d => d.Descendants().Any(z => z.Name == W.ins || z.Name == W.del) ||
                                                ContentContainsFootnoteEndnoteReferencesThatHaveRevisions(d, wDocDelta))
                                    .ToList();

                                foreach (var revision in blockLevelContentToMove)
                                {
                                    var elementLookingAt = revision;
                                    while (true)
                                    {
                                        var unid = (string)elementLookingAt.Attribute(PtOpenXml.Unid);
                                        if (unid == null)
                                        {
                                            throw new OpenXmlPowerToolsException("Internal error");
                                        }

                                        XElement elementToInsertAfter = null;
                                        if (consolidatedByUnid.ContainsKey(unid))
                                        {
                                            elementToInsertAfter = consolidatedByUnid[unid];
                                        }

                                        if (elementToInsertAfter != null)
                                        {
                                            var ci = new ConsolidationInfo
                                            {
                                                Revisor = revisedDocumentInfo.Revisor,
                                                Color = revisedDocumentInfo.Color,
                                                RevisionElement = revision,
                                                Footnotes = revision
                                                .Descendants(W.footnoteReference)
                                                .Select(fr =>
                                                {
                                                    var id = (int)fr.Attribute(W.id);
                                                    var fnXDoc = wDocDelta.MainDocumentPart.FootnotesPart.GetXDocument();
                                                    var footnote = fnXDoc.Root.Elements(W.footnote)
                                                        .FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                                                    if (footnote == null)
                                                    {
                                                        throw new OpenXmlPowerToolsException("Internal Error");
                                                    }

                                                    return footnote;
                                                })
                                                .ToArray(),
                                                Endnotes = revision
                                                .Descendants(W.endnoteReference)
                                                .Select(er =>
                                                {
                                                    var id = (int)er.Attribute(W.id);
                                                    var enXDoc = wDocDelta.MainDocumentPart.EndnotesPart.GetXDocument();
                                                    var endnote = enXDoc.Root.Elements(W.endnote)
                                                        .FirstOrDefault(en => (int)en.Attribute(W.id) == id);
                                                    if (endnote == null)
                                                    {
                                                        throw new OpenXmlPowerToolsException("Internal Error");
                                                    }

                                                    return endnote;
                                                })
                                                .ToArray()
                                            };
                                            AddToAnnotation(
                                                wDocDelta,
                                                consolidatedWDoc,
                                                elementToInsertAfter,
                                                ci,
                                                settings);
                                            break;
                                        }

                                        // find an element to insert after
                                        var elementBeforeRevision = elementLookingAt
                                            .SiblingsBeforeSelfReverseDocumentOrder()
                                            .FirstOrDefault(e => e.Attribute(PtOpenXml.Unid) != null);
                                        if (elementBeforeRevision == null)
                                        {
                                            var firstElement = consolidatedMainDocPartXDoc
                                                .Root
                                                .Element(W.body)
                                                .Elements()
                                                .FirstOrDefault(e => e.Name == W.p || e.Name == W.tbl);

                                            var ci = new ConsolidationInfo
                                            {
                                                Revisor = revisedDocumentInfo.Revisor,
                                                Color = revisedDocumentInfo.Color,
                                                RevisionElement = revision,
                                                InsertBefore = true,
                                                Footnotes = revision
                                                .Descendants(W.footnoteReference)
                                                .Select(fr =>
                                                {
                                                    var id = (int)fr.Attribute(W.id);
                                                    var fnXDoc = wDocDelta.MainDocumentPart.FootnotesPart.GetXDocument();
                                                    var footnote = fnXDoc.Root.Elements(W.footnote)
                                                        .FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                                                    if (footnote == null)
                                                    {
                                                        throw new OpenXmlPowerToolsException("Internal Error");
                                                    }

                                                    return footnote;
                                                })
                                                .ToArray(),
                                                Endnotes = revision
                                                .Descendants(W.endnoteReference)
                                                .Select(er =>
                                                {
                                                    var id = (int)er.Attribute(W.id);
                                                    var enXDoc = wDocDelta.MainDocumentPart.EndnotesPart.GetXDocument();
                                                    var endnote = enXDoc.Root.Elements(W.endnote)
                                                        .FirstOrDefault(en => (int)en.Attribute(W.id) == id);
                                                    if (endnote == null)
                                                    {
                                                        throw new OpenXmlPowerToolsException("Internal Error");
                                                    }

                                                    return endnote;
                                                })
                                                .ToArray()
                                            };
                                            AddToAnnotation(
                                                wDocDelta,
                                                consolidatedWDoc,
                                                firstElement,
                                                ci,
                                                settings);
                                            break;
                                        }

                                        elementLookingAt = elementBeforeRevision;
                                    }
                                }

                                CopyMissingStylesFromOneDocToAnother(wDocDelta, consolidatedWDoc);
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

                        // process before
                        var contentToAddBefore = lci
                            .Where(ci => ci.InsertBefore)
                            .GroupAdjacent(ci => ci.Revisor + ci.Color.ToString())
                            .Select((groupedCi, idx) => AssembledConjoinedRevisionContent(emptyParagraph, groupedCi, idx,
                                consolidatedWDoc, consolidateSettings));
                        ele.AddBeforeSelf(contentToAddBefore);

                        // process after
                        // if all revisions from all revisors are exactly the same, then instead of adding multiple tables after
                        // that contains the revisions, then simply replace the paragraph with the one with the revisions.
                        // RC004 documents contain the test data to exercise this.

                        var lciCount = lci.Where(ci => ci.InsertBefore == false).Count();

                        if (lciCount > 1 && lciCount == revisedDocumentInfoListCount)
                        {
                            // This is the code that determines if revisions should be consolidated into one.

                            var uniqueRevisions = lci
                                .Where(ci => ci.InsertBefore == false)
                                .GroupBy(ci =>
                                {
                                    // Get a hash after first accepting revisions and compressing the text.
                                    var acceptedRevisionElement =
                                        RevisionProcessor.AcceptRevisionsForElement(ci.RevisionElement);
                                    var sha1Hash = WmlComparerUtil.SHA1HashStringForUTF8String(acceptedRevisionElement.Value
                                        .Replace(" ", "").Replace(" ", "").Replace(" ", "").Replace("\n", "").Replace(".", "")
                                        .Replace(",", "").ToUpper());
                                    return sha1Hash;
                                })
                                .OrderByDescending(g => g.Count())
                                .ToList();
                            var uniqueRevisionCount = uniqueRevisions.Count();

                            if (uniqueRevisionCount == 1)
                            {
                                MoveFootnotesEndnotesForConsolidatedRevisions(lci.First(), consolidatedWDoc);

                                var dummyElement = new XElement("dummy", lci.First().RevisionElement);

                                foreach (var rev in dummyElement.Descendants().Where(d => d.Attribute(W.author) != null))
                                {
                                    var aut = rev.Attribute(W.author);
                                    aut.Value = "ITU";
                                }

                                ele.ReplaceWith(dummyElement.Elements());
                                continue;
                            }

                            // this is the location where we have determined that there are the same number of revisions for this paragraph as there are revision documents.
                            // however, the hash for all of them were not the same.
                            // therefore, they would be added to the consolidated document as separate revisions.

                            // create a log that shows what is different, in detail.
                            if (settings.LogCallback != null)
                            {
                                var sb = new StringBuilder();
                                sb.Append(
                                    "====================================================================================================" +
                                    NewLine);
                                sb.Append("Non-Consolidated Revision" + NewLine);
                                sb.Append(
                                    "====================================================================================================" +
                                    NewLine);
                                foreach (var urList in uniqueRevisions)
                                {
                                    var revisorList = urList.Select(ur => ur.Revisor + " : ").StringConcatenate()
                                        .TrimEnd(' ', ':');
                                    sb.Append("Revisors: " + revisorList + NewLine);
                                    var str = RevisionToLogFormTransform(urList.First().RevisionElement, 0, false);
                                    sb.Append(str);
                                    sb.Append("=========================" + NewLine);
                                }

                                sb.Append(NewLine);
                                settings.LogCallback(sb.ToString());
                            }
                        }

                        var contentToAddAfter = lci
                            .Where(ci => ci.InsertBefore == false)
                            .GroupAdjacent(ci => ci.Revisor + ci.Color.ToString())
                            .Select((groupedCi, idx) => AssembledConjoinedRevisionContent(emptyParagraph, groupedCi, idx,
                                consolidatedWDoc, consolidateSettings));
                        ele.AddAfterSelf(contentToAddAfter);
                    }

                    consolidatedMainDocPartXDoc
                        .Root?
                        .Element(W.body)?
                        .Add(savedSectPr);

                    AddTableGridStyleToStylesPart(consolidatedWDoc.MainDocumentPart.StyleDefinitionsPart);
                    FixUpRevisionIds(consolidatedWDoc, consolidatedMainDocPartXDoc);
                    IgnorePt14NamespaceForFootnotesEndnotes(consolidatedWDoc);
                    FixUpDocPrIds(consolidatedWDoc);
                    FixUpShapeIds(consolidatedWDoc);
                    FixUpGroupIds(consolidatedWDoc);
                    FixUpShapeTypeIds(consolidatedWDoc);
                    IgnorePt14Namespace(consolidatedMainDocPartXDoc.Root);
                    consolidatedWDoc.MainDocumentPart.PutXDocument();
                    AddFootnotesEndnotesStyles(consolidatedWDoc);
                }

                var newConsolidatedDocument = new WmlDocument("consolidated.docx", consolidatedMs.ToArray());
                return newConsolidatedDocument;
            }
        }

        private static void MoveFootnotesEndnotesForConsolidatedRevisions(
            ConsolidationInfo ci,
            WordprocessingDocument wDocConsolidated)
        {
            var consolidatedFootnoteXDoc = wDocConsolidated.MainDocumentPart.FootnotesPart.GetXDocument();
            var consolidatedEndnoteXDoc = wDocConsolidated.MainDocumentPart.EndnotesPart.GetXDocument();

            var maxFootnoteId = 1;
            if (consolidatedFootnoteXDoc.Root?.Elements(W.footnote).Any() == true)
            {
                maxFootnoteId = consolidatedFootnoteXDoc.Root.Elements(W.footnote).Select(e => (int)e.Attribute(W.id)).Max();
            }

            var maxEndnoteId = 1;
            if (consolidatedEndnoteXDoc.Root?.Elements(W.endnote).Any() == true)
            {
                maxEndnoteId = consolidatedEndnoteXDoc.Root.Elements(W.endnote).Select(e => (int)e.Attribute(W.id)).Max();
            }

            // At this point, content might contain a footnote or endnote reference.
            // Need to add the footnote / endnote into the consolidated document (with the same guid id)
            // Because of preprocessing of the documents, all footnote and endnote references will be unique at this point

            if (ci.RevisionElement.Descendants(W.footnoteReference).Any())
            {
                var footnoteXDoc = wDocConsolidated.MainDocumentPart.FootnotesPart.GetXDocument();
                foreach (var footnoteReference in ci.RevisionElement.Descendants(W.footnoteReference))
                {
                    var id = (int)footnoteReference.Attribute(W.id);
                    var footnote = ci.Footnotes.FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                    if (footnote != null)
                    {
                        var newId = ++maxFootnoteId;
                        footnoteReference.SetAttributeValue(W.id, newId);

                        var clonedFootnote = new XElement(footnote);
                        clonedFootnote.SetAttributeValue(W.id, newId);
                        footnoteXDoc.Root?.Add(clonedFootnote);
                    }
                }

                wDocConsolidated.MainDocumentPart.FootnotesPart.PutXDocument();
            }

            if (ci.RevisionElement.Descendants(W.endnoteReference).Any())
            {
                var endnoteXDoc = wDocConsolidated.MainDocumentPart.EndnotesPart.GetXDocument();
                foreach (var endnoteReference in ci.RevisionElement.Descendants(W.endnoteReference))
                {
                    var id = (int)endnoteReference.Attribute(W.id);
                    var endnote = ci.Endnotes.FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                    if (endnote != null)
                    {
                        var newId = ++maxEndnoteId;
                        endnoteReference.SetAttributeValue(W.id, newId);

                        var clonedEndnote = new XElement(endnote);
                        clonedEndnote.SetAttributeValue(W.id, newId);
                        endnoteXDoc.Root?.Add(clonedEndnote);
                    }
                }

                wDocConsolidated.MainDocumentPart.EndnotesPart.PutXDocument();
            }
        }

        private static void FixUpGroupIds(WordprocessingDocument wDoc)
        {
            var elementToFind = VML.@group;
            var groupIdsToChange = wDoc
                .ContentParts()
                .Select(cp => cp.GetXDocument())
                .Select(xd => xd.Descendants().Where(d => d.Name == elementToFind))
                .SelectMany(m => m);
            var nextId = 1;
            foreach (var item in groupIdsToChange)
            {
                var thisId = nextId++;

                var idAtt = item.Attribute("id");
                if (idAtt != null)
                {
                    idAtt.Value = thisId.ToString();
                }
            }

            foreach (var cp in wDoc.ContentParts())
            {
                cp.PutXDocument();
            }
        }

        private static bool ContentContainsFootnoteEndnoteReferencesThatHaveRevisions(
            XElement element,
            WordprocessingDocument wDocDelta)
        {
            IEnumerable<XElement> footnoteEndnoteReferences = element
                .Descendants()
                .Where(d => d.Name == W.footnoteReference || d.Name == W.endnoteReference)
                .ToList();

            if (!footnoteEndnoteReferences.Any())
            {
                return false;
            }

            var footnoteXDoc = wDocDelta.MainDocumentPart.FootnotesPart.GetXDocument();
            var endnoteXDoc = wDocDelta.MainDocumentPart.EndnotesPart.GetXDocument();

            foreach (var note in footnoteEndnoteReferences)
            {
                XElement fnen;
                if (note.Name == W.footnoteReference)
                {
                    var id = (int)note.Attribute(W.id);
                    fnen = footnoteXDoc
                        .Root?
                        .Elements(W.footnote)
                        .FirstOrDefault(n => (int)n.Attribute(W.id) == id);

                    if (fnen?.Descendants().Any(d => d.Name == W.ins || d.Name == W.del) == true)
                    {
                        return true;
                    }
                }

                if (note.Name == W.endnoteReference)
                {
                    var id = (int)note.Attribute(W.id);
                    fnen = endnoteXDoc
                        .Root?
                        .Elements(W.endnote)
                        .FirstOrDefault(n => (int)n.Attribute(W.id) == id);

                    if (fnen?.Descendants().Any(d => d.Name == W.ins || d.Name == W.del) == true)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static string RevisionToLogFormTransform(XElement element, int depth, bool inserting)
        {
            if (element.Name == W.p)
            {
                return "Paragraph" + NewLine + element.Elements().Select(e => RevisionToLogFormTransform(e, depth + 2, false))
                           .StringConcatenate();
            }

            if (element.Name == W.pPr || element.Name == W.rPr)
            {
                return "";
            }

            if (element.Name == W.r)
            {
                return element.Elements().Select(e => RevisionToLogFormTransform(e, depth, inserting)).StringConcatenate();
            }

            if (element.Name == W.t)
            {
                if (inserting)
                {
                    return "".PadRight(depth) + "Inserted Text:" + QuoteIt((string)element) + NewLine;
                }

                return "".PadRight(depth) + "Text:" + QuoteIt((string)element) + NewLine;
            }

            if (element.Name == W.delText)
            {
                return "".PadRight(depth) + "Deleted Text:" + QuoteIt((string)element) + NewLine;
            }

            if (element.Name == W.ins)
            {
                return element.Elements().Select(e => RevisionToLogFormTransform(e, depth, true)).StringConcatenate();
            }

            if (element.Name == W.del)
            {
                return element.Elements().Select(e => RevisionToLogFormTransform(e, depth, false)).StringConcatenate();
            }

            return "";
        }

        private static string QuoteIt(string str)
        {
            var quoteString = "\"";
            if (str.Contains('\"'))
            {
                quoteString = "\'";
            }

            return quoteString + str + quoteString;
        }

        private static void IgnorePt14NamespaceForFootnotesEndnotes(WordprocessingDocument wDoc)
        {
            var footnotesPart = wDoc.MainDocumentPart.FootnotesPart;
            var endnotesPart = wDoc.MainDocumentPart.EndnotesPart;

            if (footnotesPart != null)
            {
                var footnotesPartXDoc = footnotesPart.GetXDocument();
                IgnorePt14Namespace(footnotesPartXDoc.Root);
            }

            if (endnotesPart != null)
            {
                var endnotesPartXDoc = endnotesPart.GetXDocument();
                IgnorePt14Namespace(endnotesPartXDoc.Root);
            }

            footnotesPart?.PutXDocument();
            endnotesPart?.PutXDocument();
        }

        private static XElement[] AssembledConjoinedRevisionContent(
            XElement emptyParagraph,
            IGrouping<string, ConsolidationInfo> groupedCi,
            int idx,
            WordprocessingDocument wDocConsolidated,
            WmlComparerConsolidateSettings consolidateSettings)
        {
            var consolidatedFootnoteXDoc = wDocConsolidated.MainDocumentPart.FootnotesPart.GetXDocument();
            var consolidatedEndnoteXDoc = wDocConsolidated.MainDocumentPart.EndnotesPart.GetXDocument();

            var maxFootnoteId = 1;
            if (consolidatedFootnoteXDoc.Root?.Elements(W.footnote).Any() == true)
            {
                maxFootnoteId = consolidatedFootnoteXDoc.Root.Elements(W.footnote).Select(e => (int)e.Attribute(W.id)).Max();
            }

            var maxEndnoteId = 1;
            if (consolidatedEndnoteXDoc.Root?.Elements(W.endnote).Any() == true)
            {
                maxEndnoteId = consolidatedEndnoteXDoc.Root.Elements(W.endnote).Select(e => (int)e.Attribute(W.id)).Max();
            }

            var revisor = groupedCi.First().Revisor;

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
                    new XElement(W.t, revisor)));

            var colorRgb = groupedCi.First().Color.ToArgb();
            var colorString = colorRgb.ToString("X");
            if (colorString.Length == 8)
            {
                colorString = colorString.Substring(2);
            }

            if (consolidateSettings.ConsolidateWithTable)
            {
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
                                {
                                    paraAfter = emptyParagraph;
                                }

                                XElement[] revisionInTable =
                                {
                                    ci.RevisionElement,
                                    paraAfter
                                };

                                // At this point, content might contain a footnote or endnote reference.
                                // Need to add the footnote / endnote into the consolidated document (with the same
                                // guid id). Because of preprocessing of the documents, all footnote and endnote
                                // references will be unique at this point

                                AddFootnotes(ci, wDocConsolidated, ref maxFootnoteId);
                                AddEndnotes(ci, wDocConsolidated, ref maxEndnoteId);

                                return revisionInTable;
                            }))));

                // if the last paragraph has a deleted paragraph mark, then remove the deletion from the paragraph mark.
                // This is to prevent Word from misbehaving. the last paragraph in a cell must not have a deleted
                // paragraph mark.
                var theCell = table.Descendants(W.tc).FirstOrDefault();
                var lastPara = theCell?.Elements(W.p).LastOrDefault();

                if (lastPara != null)
                {
                    var isDeleted = lastPara
                        .Elements(W.pPr)
                        .Elements(W.rPr)
                        .Elements(W.del)
                        .Any();

                    if (isDeleted)
                    {
                        lastPara
                            .Elements(W.pPr)
                            .Elements(W.rPr)
                            .Elements(W.del)
                            .Remove();
                    }
                }

                XElement[] content =
                {
                    idx == 0 ? emptyParagraph : null,
                    table,
                    emptyParagraph
                };

                return content;
            }
            else
            {
                var content = groupedCi.Select(ci =>
                {
                    XElement paraAfter = null;
                    if (ci.RevisionElement.Name == W.tbl)
                    {
                        paraAfter = emptyParagraph;
                    }

                    XElement[] revisionInTable =
                    {
                        ci.RevisionElement,
                        paraAfter
                    };

                    // At this point, content might contain a footnote or endnote reference.
                    // Need to add the footnote / endnote into the consolidated document (with the same
                    // guid id). Because of preprocessing of the documents, all footnote and endnote
                    // references will be unique at this point

                    AddFootnotes(ci, wDocConsolidated, ref maxFootnoteId);
                    AddEndnotes(ci, wDocConsolidated, ref maxEndnoteId);

                    return revisionInTable;
                });

                var dummyElement = new XElement("dummy", content.SelectMany(m => m));
                foreach (var rev in dummyElement.Descendants().Where(d => d.Attribute(W.author) != null))
                {
                    rev.SetAttributeValue(W.author, revisor);
                }

                return dummyElement.Elements().ToArray();
            }
        }

        private static void AddFootnotes(ConsolidationInfo ci, WordprocessingDocument wDocConsolidated, ref int maxFootnoteId)
        {
            if (ci.RevisionElement.Descendants(W.footnoteReference).Any())
            {
                var footnoteXDoc = wDocConsolidated.MainDocumentPart.FootnotesPart.GetXDocument();
                foreach (var footnoteReference in ci.RevisionElement.Descendants(W.footnoteReference))
                {
                    var id = (int)footnoteReference.Attribute(W.id);
                    var footnote = ci.Footnotes.FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                    if (footnote != null)
                    {
                        var newId = ++maxFootnoteId;
                        footnoteReference.SetAttributeValue(W.id, newId);

                        var clonedFootnote = new XElement(footnote);
                        clonedFootnote.SetAttributeValue(W.id, newId);
                        footnoteXDoc.Root?.Add(clonedFootnote);
                    }
                }

                wDocConsolidated.MainDocumentPart.FootnotesPart.PutXDocument();
            }
        }

        private static void AddEndnotes(ConsolidationInfo ci, WordprocessingDocument wDocConsolidated, ref int maxEndnoteId)
        {
            if (ci.RevisionElement.Descendants(W.endnoteReference).Any())
            {
                var endnoteXDoc = wDocConsolidated.MainDocumentPart.EndnotesPart.GetXDocument();
                foreach (var endnoteReference in ci.RevisionElement.Descendants(W.endnoteReference))
                {
                    var id = (int)endnoteReference.Attribute(W.id);
                    var endnote = ci.Endnotes.FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                    if (endnote != null)
                    {
                        var newId = ++maxEndnoteId;
                        endnoteReference.SetAttributeValue(W.id, newId);

                        var clonedEndnote = new XElement(endnote);
                        clonedEndnote.SetAttributeValue(W.id, newId);
                        endnoteXDoc.Root?.Add(clonedEndnote);
                    }
                }

                wDocConsolidated.MainDocumentPart.EndnotesPart.PutXDocument();
            }
        }

        private static void AddToAnnotation(
            WordprocessingDocument wDocDelta,
            WordprocessingDocument consolidatedWDoc,
            XElement elementToInsertAfter,
            ConsolidationInfo consolidationInfo,
            WmlComparerSettings settings)
        {
            var packageOfDeletedContent = wDocDelta.MainDocumentPart.OpenXmlPackage.Package;
            var packageOfNewContent = consolidatedWDoc.MainDocumentPart.OpenXmlPackage.Package;
            var partInDeletedDocument = packageOfDeletedContent.GetPart(wDocDelta.MainDocumentPart.Uri);
            var partInNewDocument = packageOfNewContent.GetPart(consolidatedWDoc.MainDocumentPart.Uri);
            consolidationInfo.RevisionElement =
                MoveRelatedPartsToDestination(partInDeletedDocument, partInNewDocument, consolidationInfo.RevisionElement);

            var clonedForHashing = (XElement)CloneBlockLevelContentForHashing(consolidatedWDoc.MainDocumentPart,
                consolidationInfo.RevisionElement, false, settings);
            clonedForHashing.Descendants().Where(d => d.Name == W.ins || d.Name == W.del).Attributes(W.id).Remove();
            var shaString = clonedForHashing.ToString(SaveOptions.DisableFormatting)
                .Replace(" xmlns=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", "");
            var sha1Hash = WmlComparerUtil.SHA1HashStringForUTF8String(shaString);
            consolidationInfo.RevisionString = shaString;
            consolidationInfo.RevisionHash = sha1Hash;

            var annotationList = elementToInsertAfter.Annotation<List<ConsolidationInfo>>();
            if (annotationList == null)
            {
                annotationList = new List<ConsolidationInfo>();
                elementToInsertAfter.AddAnnotation(annotationList);
            }

            annotationList.Add(consolidationInfo);
        }

        private static void AddTableGridStyleToStylesPart(StyleDefinitionsPart styleDefinitionsPart)
        {
            var sXDoc = styleDefinitionsPart.GetXDocument();
            var tableGridStyle = sXDoc
                .Root?
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
                    @"<w:style w:type=""table""
           w:default=""1""
           w:styleId=""TableNormal""
           xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
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
    }
}