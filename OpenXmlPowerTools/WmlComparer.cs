// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

// TODO Line 1202 there are inefficient calls to PutXDocument() for footnotes and endnotes
// TODO wDocConsolidated.MainDocumentPart.FootnotesPart.PutXDocument();
// TODO Take care of this after the conference

using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using System.Drawing;
using System.Security.Cryptography;
using OpenXmlPowerTools;

// It is possible to optimize DescendantContentAtoms

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/// Currently, the unid is set at the beginning of the algorithm.  It is used by the code that establishes correlation based on first rejecting
/// tracked revisions, then correlating paragraphs/tables.  It is requred for this algorithm - after finding a correlated sequence in the document with rejected
/// revisions, it uses the unid to find the same paragraph in the document without rejected revisions, then sets the correlated sha1 hash in that document.
/// 
/// But then when accepting tracked revisions, for certain paragraphs (where there are deleted paragraph marks) it is going to lose the unids.  But this isn't a
/// problem because when paragraph marks are deleted, the correlation is definitely no longer possible.  Any paragraphs that are in a range of paragraphs that
/// are coalesced can't be correlated to paragraphs in the other document via their hash.  At that point we no longer care what their unids are.
/// 
/// But after that it is only used to reconstruct the tree.  It is also used in the debugging code that
/// prints the various correlated sequences and comparison units - this is display for debugging purposes only.
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/// The key idea here is that a given paragraph will always have the same ancestors, and it doesn't matter whether the content was deleted from the old document,
/// inserted into the new document, or set as equal.  At this point, we identify a paragraph as a sequential list of content atoms, terminated by a paragraph mark.
/// This entire list will for a single paragraph, regardless of whether the paragraph is a child of the body, or if the paragraph is in a cell in a table, or if
/// the paragraph is in a text box.  The list of ancestors, from the paragraph to the root of the XML tree will be the same for all content atoms in the paragraph.
/// 
/// Therefore:
/// 
/// Iterate through the list of content atoms backwards.  When the loop sees a paragraph mark, it gets the ancestor unids from the paragraph mark to the top of the
/// tree, and sets this as the same for all content atoms in the paragraph.  For descendants of the paragraph mark, it doesn't really matter if content is put into
/// separate runs or what not.  We don't need to be concerned about what the unids are for descendants of the paragraph.
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


namespace OpenXmlPowerTools
{
    public class WmlComparerSettings
    {
        public char[] WordSeparators;
        public string AuthorForRevisions = "Open-Xml-PowerTools";
        public string DateTimeForRevisions = DateTime.Now.ToString("o");
        public double DetailThreshold = 0.15;
        public bool CaseInsensitive = false;
        public bool ConflateBreakingAndNonbreakingSpaces = true;
        public CultureInfo CultureInfo = null;
        public Action<string> LogCallback = null;
        public int StartingIdForFootnotesEndnotes = 1;

        public DirectoryInfo DebugTempFileDi;

        public WmlComparerSettings()
        {
            // note that , and . are processed explicitly to handle cases where they are in a number or word
            WordSeparators = new[] { ' ', '-', ')', '(', ';', ',', '（', '）', '，', '、', '、', '，', '；', '。', '：', '的', }; // todo need to fix this for complete list
        }
    }

    public class WmlComparerConsolidateSettings
    {
        public bool ConsolidateWithTable = true;
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
        public static bool s_SaveIntermediateFilesForDebugging = false;

        public static WmlDocument Compare(WmlDocument source1, WmlDocument source2, WmlComparerSettings settings)
        {
            return CompareInternal(source1, source2, settings, true);
        }

        private static WmlDocument CompareInternal(WmlDocument source1, WmlDocument source2, WmlComparerSettings settings,
            bool preProcessMarkupInOriginal)
        {
            if (preProcessMarkupInOriginal)
                source1 = PreProcessMarkup(source1, settings.StartingIdForFootnotesEndnotes + 1000);
            source2 = PreProcessMarkup(source2, settings.StartingIdForFootnotesEndnotes + 2000);

            if (s_SaveIntermediateFilesForDebugging && settings.DebugTempFileDi != null)
            {
                var name1 = "Source1-Step1-PreProcess.docx";
                var name2 = "Source2-Step1-PreProcess.docx";
                var preProcFi1 = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name1));
                source1.SaveAs(preProcFi1.FullName);
                var preProcFi2 = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name2));
                source2.SaveAs(preProcFi2.FullName);
            }

            // at this point, both source1 and source2 have unid on every element.  These are the values that will enable reassembly of the XML tree.
            // but we need other values.

            // In source1:
            // - accept tracked revisions
            // - determine hash code for every block-level element
            // - save as attribute on every element

            // - accept tracked revisions and reject tracked revisions leave the unids alone, where possible.
            // - after accepting and calculating the hash, then can use the unids to find the right block-level element in the unmodified source1, and install the hash

            // In source2:
            // - reject tracked revisions
            // - determine hash code for every block-level element
            // - save as an attribute on every element

            // - after rejecting and calculating the hash, then can use the unids to find the right block-level element in the unmodified source2, and install the hash

            // - sometimes after accepting or rejecting tracked revisions, several paragraphs will get coalesced into a single paragraph due to paragraph marks being inserted / deleted.
            // - in this case, some paragraphs will not get a hash injected onto them.
            // - if a paragraph doesn't have a hash, then it will never correspond to another paragraph, and such issues will need to be resolved in the normal execution of the LCS algorithm.
            // - note that when we do propagate the unid through for the first paragraph.

            // Establish correlation between the two.
            // Find the longest common sequence of block-level elements where hash codes are the same.
            // this sometimes will be every block level element in the document.  Or sometimes will be just a fair number of them.

            // at the start of doing the LCS algorithm, we will match up content, and put them in corresponding unknown correlated comparison units.  Those paragraphs will only ever be matched to their corresponding paragraph.
            // then the algorithm can proceed as usual.

            // need to call ChangeFootnoteEndnoteReferencesToUniqueRange before creating the wmlResult document, so that
            // the same GUID ids are used for footnote and endnote references in both the 'after' document, and in the
            // result document.

            var source1afterAccepting = RevisionProcessor.AcceptRevisions(source1);
            var source2afterRejecting = RevisionProcessor.RejectRevisions(source2);

            if (s_SaveIntermediateFilesForDebugging && settings.DebugTempFileDi != null)
            {
                var name1 = "Source1-Step2-AfterAccepting.docx";
                var name2 = "Source2-Step2-AfterRejecting.docx";
                var afterAcceptingFi1 = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name1));
                source1afterAccepting.SaveAs(afterAcceptingFi1.FullName);
                var afterRejectingFi2 = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name2));
                source2afterRejecting.SaveAs(afterRejectingFi2.FullName);
            }

            // this creates the correlated hash codes that enable us to match up ranges of paragraphs based on
            // accepting in source1, rejecting in source2
            source1 = HashBlockLevelContent(source1, source1afterAccepting, settings);
            source2 = HashBlockLevelContent(source2, source2afterRejecting, settings);

            if (s_SaveIntermediateFilesForDebugging && settings.DebugTempFileDi != null)
            {
                var name1 = "Source1-Step3-AfterHashing.docx";
                var name2 = "Source2-Step3-AfterHashing.docx";
                var afterHashingFi1 = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name1));
                source1.SaveAs(afterHashingFi1.FullName);
                var afterHashingFi2 = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name2));
                source2.SaveAs(afterHashingFi2.FullName);
            }

            // Accept revisions in before, and after
            source1 = RevisionProcessor.AcceptRevisions(source1);
            source2 = RevisionProcessor.AcceptRevisions(source2);

            if (s_SaveIntermediateFilesForDebugging && settings.DebugTempFileDi != null)
            {
                var name1 = "Source1-Step4-AfterAccepting.docx";
                var name2 = "Source2-Step4-AfterAccepting.docx";
                var afterAcceptingFi1 = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name1));
                source1.SaveAs(afterAcceptingFi1.FullName);
                var afterAcceptingFi2 = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name2));
                source2.SaveAs(afterAcceptingFi2.FullName);
            }

            // after accepting revisions, some unids may have been removed by revision accepter, along with the correlatedSHA1Hash codes,
            // this is as it should be.
            // but need to go back in and add guids to paragraphs that have had them removed.

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(source2.DocumentByteArray, 0, source2.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    AddUnidsToMarkupInContentParts(wDoc);
                }
            }

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
                }

                if (s_False && settings.DebugTempFileDi != null)
                {
                    var name1 = "Source1-Step5-AfterProducingDocWithRevTrk.docx";
                    var name2 = "Source2-Step5-AfterProducingDocWithRevTrk.docx";
                    var afterProducingFi1 = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name1));
                    var afterProducingWml1 = new WmlDocument("after1.docx", ms1.ToArray());
                    afterProducingWml1.SaveAs(afterProducingFi1.FullName);
                    var afterProducingFi2 = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name2));
                    var afterProducingWml2 = new WmlDocument("after2.docx", ms2.ToArray());
                    afterProducingWml2.SaveAs(afterProducingFi2.FullName);
                }

                if (s_False && settings.DebugTempFileDi != null)
                {
                    var cleanedSource = CleanPowerToolsAndRsid(source1);
                    var name1 = "Cleaned-Source.docx";
                    var cleanedSourceFi1 = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name1));
                    cleanedSource.SaveAs(cleanedSourceFi1.FullName);

                    var cleanedProduced = CleanPowerToolsAndRsid(producedDocument);
                    var name2 = "Cleaned-Produced.docx";
                    var cleanedProducedFi1 = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name2));
                    cleanedProduced.SaveAs(cleanedProducedFi1.FullName);
                }

                return producedDocument;
            }
        }

        private static WmlDocument CleanPowerToolsAndRsid(WmlDocument producedDocument)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(producedDocument.DocumentByteArray, 0, producedDocument.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    foreach (var cp in wDoc.ContentParts())
                    {
                        var xd = cp.GetXDocument();
                        var newRoot = CleanPartTransform(xd.Root);
                        xd.Root.ReplaceWith(newRoot);
                        cp.PutXDocument();
                    }
                }
                var cleaned = new WmlDocument("cleaned.docx", ms.ToArray());
                return cleaned;
            }
        }

        private static WmlDocument HashBlockLevelContent(WmlDocument source, WmlDocument source1afterProcessingRevTracking, WmlComparerSettings settings)
        {
            using (MemoryStream msSource = new MemoryStream())
            using (MemoryStream msAfterProc = new MemoryStream())
            {
                msSource.Write(source.DocumentByteArray, 0, source.DocumentByteArray.Length);
                msAfterProc.Write(source1afterProcessingRevTracking.DocumentByteArray, 0, source1afterProcessingRevTracking.DocumentByteArray.Length);
                using (WordprocessingDocument wDocSource = WordprocessingDocument.Open(msSource, true))
                using (WordprocessingDocument wDocAfterProc = WordprocessingDocument.Open(msAfterProc, true))
                {
                    // create Unid dictionary for source
                    var sourceMainXDoc = wDocSource
                        .MainDocumentPart
                        .GetXDocument();

                    var sourceUnidDict = sourceMainXDoc
                        .Root
                        .Descendants()
                        .Where(d => d.Name == W.p || d.Name == W.tbl || d.Name == W.tr)
                        .ToDictionary(d => (string)d.Attribute(PtOpenXml.Unid));

                    var afterProcMainXDoc = wDocAfterProc
                        .MainDocumentPart
                        .GetXDocument();

                    foreach (var blockLevelContent in afterProcMainXDoc.Root.Descendants().Where(d => d.Name == W.p || d.Name == W.tbl || d.Name == W.tr))
                    {
                        var cloneBlockLevelContentForHashing = (XElement)CloneBlockLevelContentForHashing(wDocAfterProc.MainDocumentPart, blockLevelContent, true, settings);
                        var shaString = cloneBlockLevelContentForHashing.ToString(SaveOptions.DisableFormatting)
                            .Replace(" xmlns=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", "");
                        var sha1Hash = PtUtils.SHA1HashStringForUTF8String(shaString);
                        var thisUnid = (string)blockLevelContent.Attribute(PtOpenXml.Unid);
                        if (thisUnid != null)
                        {
                            if (sourceUnidDict.ContainsKey(thisUnid))
                            {
                                var correlatedBlockLevelContent = sourceUnidDict[thisUnid];
                                correlatedBlockLevelContent.Add(new XAttribute(PtOpenXml.CorrelatedSHA1Hash, sha1Hash));
                            }
                        }
                    }

                    wDocSource.MainDocumentPart.PutXDocument();
                }
                WmlDocument sourceWithCorrelatedSHA1Hash = new WmlDocument(source.FileName, msSource.ToArray());
                return sourceWithCorrelatedSHA1Hash;
            }
        }

        private static WmlDocument PreProcessMarkup(WmlDocument source, int startingIdForFootnotesEndnotes)
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
                source = new WmlDocument(source.FileName, ms.ToArray());
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
                        AcceptRevisions = false,
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
                    ChangeFootnoteEndnoteReferencesToUniqueRange(wDoc, startingIdForFootnotesEndnotes);
                    AddUnidsToMarkupInContentParts(wDoc);
                    AddFootnotesEndnotesParts(wDoc);
                    FillInEmptyFootnotesEndnotes(wDoc);
                    DetachExternalData(wDoc);
                }
                return new WmlDocument(source.FileName, ms.ToArray());
            }
        }

        private static void DetachExternalData(WordprocessingDocument wDoc)
        {
            // External data for chart parts contains relationships to external links, which are not properly propagated to the destination document (There is little point to doing so.)
            // Therefore remove them.

            foreach (var chart in wDoc.MainDocumentPart.ChartParts)
            {
                var cxd = chart.GetXDocument();
                cxd.Descendants(C.externalData).Remove();
                chart.PutXDocument();
            }
        }

        // somehow, sometimes a footnote or endnote contains absolutely nothing - no paragraph - nothing.
        // This messes up the algorithm, so in this case, insert an empty paragraph.
        // This is pretty wacky markup to find, and I don't know how this markup comes into existence, but this is an innocuous fix.
        private static void FillInEmptyFootnotesEndnotes(WordprocessingDocument wDoc)
        {
            XElement emptyFootnote = XElement.Parse(
@"<w:p xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:pPr>
    <w:pStyle w:val='FootnoteText'/>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:rStyle w:val='FootnoteReference'/>
    </w:rPr>
    <w:footnoteRef/>
  </w:r>
</w:p>");

            XElement emptyEndnote = XElement.Parse(
@"<w:p xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:pPr>
    <w:pStyle w:val='EndnoteText'/>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:rStyle w:val='EndnoteReference'/>
    </w:rPr>
    <w:endnoteRef/>
  </w:r>
</w:p>");

            var footnotePart = wDoc.MainDocumentPart.FootnotesPart;
            if (footnotePart != null)
            {
                var fnXDoc = footnotePart.GetXDocument();
                foreach (var fn in fnXDoc.Root.Elements(W.footnote))
                {
                    if (!fn.HasElements)
                        fn.Add(emptyFootnote);
                }
                footnotePart.PutXDocument();
            }

            var endnotePart = wDoc.MainDocumentPart.EndnotesPart;
            if (endnotePart != null)
            {
                var fnXDoc = endnotePart.GetXDocument();
                foreach (var fn in fnXDoc.Root.Elements(W.endnote))
                {
                    if (!fn.HasElements)
                        fn.Add(emptyEndnote);
                }
                endnotePart.PutXDocument();
            }
        }

        private static bool ContentContainsFootnoteEndnoteReferencesThatHaveRevisions(XElement element, WordprocessingDocument wDocDelta)
        {
            var footnoteEndnoteReferences = element.Descendants().Where(d => d.Name == W.footnoteReference || d.Name == W.endnoteReference);
            if (!footnoteEndnoteReferences.Any())
                return false;
            var footnoteXDoc = wDocDelta.MainDocumentPart.FootnotesPart.GetXDocument();
            var endnoteXDoc = wDocDelta.MainDocumentPart.EndnotesPart.GetXDocument();
            foreach (var note in footnoteEndnoteReferences)
            {
                XElement fnen = null;
                if (note.Name == W.footnoteReference)
                {
                    var id = (int)note.Attribute(W.id);
                    fnen = footnoteXDoc
                        .Root
                        .Elements(W.footnote)
                        .FirstOrDefault(n => (int)n.Attribute(W.id) == id);
                    if (fnen.Descendants().Where(d => d.Name == W.ins || d.Name == W.del).Any())
                        return true;
                }
                if (note.Name == W.endnoteReference)
                {
                    var id = (int)note.Attribute(W.id);
                    fnen = endnoteXDoc
                        .Root
                        .Elements(W.endnote)
                        .FirstOrDefault(n => (int)n.Attribute(W.id) == id);
                    if (fnen.Descendants().Where(d => d.Name == W.ins || d.Name == W.del).Any())
                        return true;
                }
            }
            return false;
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
            public XElement[] Footnotes;
            public XElement[] Endnotes;
            public string RevisionString; // for debugging purposes only
        }

        private static string nl = Environment.NewLine;

        /*****************************************************************************************************************/
        // Consolidate processes footnotes and endnotes in a particular fashion - if the unmodified document has a footnote
        // reference, and a delta has a footnote reference, we end up with two footnotes - one is unmodified, and is refered to
        // from the unmodified content.  The footnote reference in the delta refers to the modified footnote.  This is as it
        // should be.
        /*****************************************************************************************************************/
        public static WmlDocument Consolidate(WmlDocument original,
            List<WmlRevisedDocumentInfo> revisedDocumentInfoList,
            WmlComparerSettings settings)
        {
            var consolidateSettings = new WmlComparerConsolidateSettings();
            return Consolidate(original, revisedDocumentInfoList, settings, consolidateSettings);
        }

        public static WmlDocument Consolidate(WmlDocument original,
            List<WmlRevisedDocumentInfo> revisedDocumentInfoList,
            WmlComparerSettings settings, WmlComparerConsolidateSettings consolidateSettings)
        {

#if false
            var now = DateTime.Now;
            var tempName = String.Format("{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", now.Year - 2000, now.Month, now.Day, now.Hour, now.Minute, now.Second);
            FileInfo fi = new FileInfo("./WmlComparer.Consolidate-" + tempName + "-Original.docx");
            File.WriteAllBytes(fi.FullName, original.DocumentByteArray);
            for (int i = 0; i < revisedDocumentInfoList.Count(); i++)
            {
                fi = new FileInfo("./WmlComparer.Consolidate-" + tempName + string.Format("-Revised-{0}", i) + ".docx");
                File.WriteAllBytes(fi.FullName, revisedDocumentInfoList.ElementAt(i).RevisedDocument.DocumentByteArray);
            }
            StringBuilder sbt = new StringBuilder();
            int count = 0;
            foreach (var rev in revisedDocumentInfoList)
            {
                sbt.Append("Revised #" + (count++).ToString() + Environment.NewLine);
                sbt.Append("Color:" + rev.Color.ToString() + Environment.NewLine);
                sbt.Append("Revisor:" + rev.Revisor + Environment.NewLine);
                sbt.Append("" + Environment.NewLine);
            }
            sbt.Append("settings.AuthorForRevisions:" + settings.AuthorForRevisions + Environment.NewLine);
            sbt.Append("settings.CaseInsensitive:" + settings.CaseInsensitive.ToString() + Environment.NewLine);
            sbt.Append("settings.CultureInfo:" + settings.CultureInfo.ToString() + Environment.NewLine);
            sbt.Append("settings.DateTimeForRevisions:" + settings.DateTimeForRevisions.ToString() + Environment.NewLine);
            sbt.Append("settings.DetailThreshold:" + settings.DetailThreshold.ToString() + Environment.NewLine);
            sbt.Append("settings.StartingIdForFootnotesEndnotes:" + settings.StartingIdForFootnotesEndnotes.ToString() + Environment.NewLine);
            sbt.Append("settings.WordSeparators:" + settings.WordSeparators.Select(ws => ws.ToString()).StringConcatenate() + Environment.NewLine);
            //sb.Append(":" + settings);
            fi = new FileInfo("./WmlComparer.Consolidate-" + tempName + "-Settings.txt");
            File.WriteAllText(fi.FullName, sbt.ToString());
#endif

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
            WmlDocument consolidated = new WmlDocument(originalWithUnids);

            if (s_SaveIntermediateFilesForDebugging && settings.DebugTempFileDi != null)
            {
                var name1 = "Original-with-Unids.docx";
                var preProcFi1 = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name1));
                originalWithUnids.SaveAs(preProcFi1.FullName);
            }

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

                    var consolidatedByUnid = consolidatedMainDocPartXDoc
                        .Descendants()
                        .Where(d => (d.Name == W.p || d.Name == W.tbl) && d.Attribute(PtOpenXml.Unid) != null)
                        .ToDictionary(d => (string)d.Attribute(PtOpenXml.Unid));

                    int deltaNbr = 1;
                    foreach (var revisedDocumentInfo in revisedDocumentInfoList)
                    {
                        settings.StartingIdForFootnotesEndnotes = (deltaNbr * 2000) + 3000;
                        var delta = WmlComparer.CompareInternal(originalWithUnids, revisedDocumentInfo.RevisedDocument, settings, false);

                        if (s_SaveIntermediateFilesForDebugging && settings.DebugTempFileDi != null)
                        {
                            var name1 = string.Format("Delta-{0}.docx", deltaNbr++);
                            var deltaFi = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name1));
                            delta.SaveAs(deltaFi.FullName);
                        }

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
                                            throw new OpenXmlPowerToolsException("Internal error");

                                        XElement elementToInsertAfter = null;
                                        if (consolidatedByUnid.ContainsKey(unid))
                                            elementToInsertAfter = consolidatedByUnid[unid];

                                        if (elementToInsertAfter != null)
                                        {
                                            ConsolidationInfo ci = new ConsolidationInfo();
                                            ci.Revisor = revisedDocumentInfo.Revisor;
                                            ci.Color = revisedDocumentInfo.Color;
                                            ci.RevisionElement = revision;
                                            ci.Footnotes = revision
                                                .Descendants(W.footnoteReference)
                                                .Select(fr =>
                                                {
                                                    var id = (int)fr.Attribute(W.id);
                                                    var fnXDoc = wDocDelta.MainDocumentPart.FootnotesPart.GetXDocument();
                                                    var footnote = fnXDoc.Root.Elements(W.footnote).FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                                                    if (footnote == null)
                                                        throw new OpenXmlPowerToolsException("Internal Error");
                                                    return footnote;
                                                })
                                                .ToArray();
                                            ci.Endnotes = revision
                                                .Descendants(W.endnoteReference)
                                                .Select(er =>
                                                {
                                                    var id = (int)er.Attribute(W.id);
                                                    var enXDoc = wDocDelta.MainDocumentPart.EndnotesPart.GetXDocument();
                                                    var endnote = enXDoc.Root.Elements(W.endnote).FirstOrDefault(en => (int)en.Attribute(W.id) == id);
                                                    if (endnote == null)
                                                        throw new OpenXmlPowerToolsException("Internal Error");
                                                    return endnote;
                                                })
                                                .ToArray();
                                            AddToAnnotation(
                                                wDocDelta,
                                                consolidatedWDoc,
                                                elementToInsertAfter,
                                                ci,
                                                settings);
                                            break;
                                        }
                                        else
                                        {
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

                                                ConsolidationInfo ci = new ConsolidationInfo();
                                                ci.Revisor = revisedDocumentInfo.Revisor;
                                                ci.Color = revisedDocumentInfo.Color;
                                                ci.RevisionElement = revision;
                                                ci.InsertBefore = true;
                                                ci.Footnotes = revision
                                                    .Descendants(W.footnoteReference)
                                                    .Select(fr =>
                                                    {
                                                        var id = (int)fr.Attribute(W.id);
                                                        var fnXDoc = wDocDelta.MainDocumentPart.FootnotesPart.GetXDocument();
                                                        var footnote = fnXDoc.Root.Elements(W.footnote).FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                                                        if (footnote == null)
                                                            throw new OpenXmlPowerToolsException("Internal Error");
                                                        return footnote;
                                                    })
                                                    .ToArray();
                                                ci.Endnotes = revision
                                                    .Descendants(W.endnoteReference)
                                                    .Select(er =>
                                                    {
                                                        var id = (int)er.Attribute(W.id);
                                                        var enXDoc = wDocDelta.MainDocumentPart.EndnotesPart.GetXDocument();
                                                        var endnote = enXDoc.Root.Elements(W.endnote).FirstOrDefault(en => (int)en.Attribute(W.id) == id);
                                                        if (endnote == null)
                                                            throw new OpenXmlPowerToolsException("Internal Error");
                                                        return endnote;
                                                    })
                                                    .ToArray();
                                                AddToAnnotation(
                                                    wDocDelta,
                                                    consolidatedWDoc,
                                                    firstElement,
                                                    ci,
                                                    settings);
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
                            .Where(ci => ci.InsertBefore == true)
                            .GroupAdjacent(ci => ci.Revisor + ci.Color.ToString())
                            .Select((groupedCi, idx) => AssembledConjoinedRevisionContent(emptyParagraph, groupedCi, idx, consolidatedWDoc, consolidateSettings));
                        ele.AddBeforeSelf(contentToAddBefore);

                        // process after
                        // if all revisions from all revisors are exactly the same, then instead of adding multiple tables after
                        // that contains the revisions, then simply replace the paragraph with the one with the revisions.
                        // RC004 documents contain the test data to exercise this.

                        var lciCount = lci.Where(ci => ci.InsertBefore == false).Count();

                        if (lciCount > 1 && lciCount == revisedDocumentInfoListCount)
                        {
                            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                            // This is the code that determines if revisions should be consolidated into one.

                            var uniqueRevisions = lci
                                .Where(ci => ci.InsertBefore == false)
                                .GroupBy(ci =>
                                {
                                    // Get a hash after first accepting revisions and compressing the text.
                                    var acceptedRevisionElement = RevisionProcessor.AcceptRevisionsForElement(ci.RevisionElement);
                                    var sha1Hash = PtUtils.SHA1HashStringForUTF8String(acceptedRevisionElement.Value.Replace(" ", "").Replace(" ", "").Replace(" ", "").Replace("\n", "").Replace(".", "").Replace(",", "").ToUpper());
                                    return sha1Hash;
                                })
                                .OrderByDescending(g => g.Count())
                                .ToList();
                            var uniqueRevisionCount = uniqueRevisions.Count();

                            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

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
                                StringBuilder sb = new StringBuilder();
                                sb.Append("====================================================================================================" + nl);
                                sb.Append("Non-Consolidated Revision" + nl);
                                sb.Append("====================================================================================================" + nl);
                                foreach (var urList in uniqueRevisions)
                                {
                                    var revisorList = urList.Select(ur => ur.Revisor + " : ").StringConcatenate().TrimEnd(' ', ':');
                                    sb.Append("Revisors: " + revisorList + nl);
                                    var str = RevisionToLogFormTransform(urList.First().RevisionElement, 0, false);
                                    sb.Append(str);
                                    sb.Append("=========================" + nl);
                                }
                                sb.Append(nl);
                                settings.LogCallback(sb.ToString());
                            }
                        }

                        // todo this is where it assembles the content to put into a single cell table
                        // the magic function is AssembledConjoinedRevisionContent

                        var contentToAddAfter = lci
                            .Where(ci => ci.InsertBefore == false)
                            .GroupAdjacent(ci => ci.Revisor + ci.Color.ToString())
                            .Select((groupedCi, idx) => AssembledConjoinedRevisionContent(emptyParagraph, groupedCi, idx, consolidatedWDoc, consolidateSettings));
                        ele.AddAfterSelf(contentToAddAfter);
                    }

#if false
                    // old code
                    foreach (var ele in elementsToProcess)
                    {
                        var lci = ele.Annotation<List<ConsolidationInfo>>();

                        // if all revisions from all revisors are exactly the same, then instead of adding multiple tables after
                        // that contains the revisions, then simply replace the paragraph with the one with the revisions.
                        // RC004 documents contain the test data to exercise this.

                        var lciCount = lci.Count();

                        if (lci.Count() > 1 && lciCount == revisedDocumentInfoListCount)
                        {
                            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                            // This is the code that determines if revisions should be consolidated into one.

                            var uniqueRevisions = lci
                                .GroupBy(ci =>
                                {
                                    // Get a hash after first accepting revisions and compressing the text.
                                    var ciz = ci;

                                    var acceptedRevisionElement = RevisionProcessor.AcceptRevisionsForElement(ci.RevisionElement);
                                    var text = acceptedRevisionElement.Value
                                        .Replace(" ", "")
                                        .Replace(" ", "")
                                        .Replace(" ", "")
                                        .Replace("\n", "");
                                    var sha1Hash = PtUtils.SHA1HashStringForUTF8String(text);
                                    return ci.InsertBefore.ToString() + sha1Hash;
                                })
                                .OrderByDescending(g => g.Count())
                                .ToList();
                            var uniqueRevisionCount = uniqueRevisions.Count();

                            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                            if (uniqueRevisionCount == 1)
                            {
                                MoveFootnotesEndnotesForConsolidatedRevisions(lci.First(), consolidatedWDoc);

                                var dummyElement = new XElement("dummy", lci.First().RevisionElement);

                                foreach(var rev in dummyElement.Descendants().Where(d => d.Attribute(W.author) != null))
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
                                StringBuilder sb = new StringBuilder();
                                sb.Append("====================================================================================================" + nl);
                                sb.Append("Non-Consolidated Revision" + nl);
                                sb.Append("====================================================================================================" + nl);
                                foreach (var urList in uniqueRevisions)
                                {
                                    var revisorList = urList.Select(ur => ur.Revisor + " : ").StringConcatenate().TrimEnd(' ', ':');
                                    sb.Append("Revisors: " + revisorList + nl);
                                    var str = RevisionToLogFormTransform(urList.First().RevisionElement, 0, false);
                                    sb.Append(str);
                                    sb.Append("=========================" + nl);
                                }
                                sb.Append(nl);
                                settings.LogCallback(sb.ToString());
                            }
                        }

                        var contentToAddBefore = lci
                            .Where(ci => ci.InsertBefore == true)
                            .GroupAdjacent(ci => ci.Revisor + ci.Color.ToString())
                            .Select((groupedCi, idx) => AssembledConjoinedRevisionContent(emptyParagraph, groupedCi, idx, consolidatedWDoc, consolidateSettings));
                        var contentToAddAfter = lci
                            .Where(ci => ci.InsertBefore == false)
                            .GroupAdjacent(ci => ci.Revisor + ci.Color.ToString())
                            .Select((groupedCi, idx) => AssembledConjoinedRevisionContent(emptyParagraph, groupedCi, idx, consolidatedWDoc, consolidateSettings));
                        ele.AddBeforeSelf(contentToAddBefore);
                        ele.AddAfterSelf(contentToAddAfter);
                    }
#endif

                    consolidatedMainDocPartXDoc
                        .Root
                        .Element(W.body)
                        .Add(savedSectPr);

                    AddTableGridStyleToStylesPart(consolidatedWDoc.MainDocumentPart.StyleDefinitionsPart);
                    FixUpRevisionIds(consolidatedWDoc, consolidatedMainDocPartXDoc);
                    IgnorePt14NamespaceForFootnotesEndnotes(consolidatedWDoc);
                    FixUpDocPrIds(consolidatedWDoc);
                    FixUpShapeIds(consolidatedWDoc);
                    FixUpGroupIds(consolidatedWDoc);
                    FixUpShapeTypeIds(consolidatedWDoc);
                    RemoveCustomMarkFollows(consolidatedWDoc);
                    WmlComparer.IgnorePt14Namespace(consolidatedMainDocPartXDoc.Root);
                    consolidatedWDoc.MainDocumentPart.PutXDocument();
                    AddFootnotesEndnotesStyles(consolidatedWDoc);
                }

                var newConsolidatedDocument = new WmlDocument("consolidated.docx", consolidatedMs.ToArray());
                return newConsolidatedDocument;
            }
        }

        private static void RemoveCustomMarkFollows(WordprocessingDocument consolidatedWDoc)
        {
            var mxDoc = consolidatedWDoc.MainDocumentPart.GetXDocument();
            mxDoc.Root.Descendants().Attributes(W.customMarkFollows).Remove();
            consolidatedWDoc.MainDocumentPart.PutXDocument();
        }

        private static void MoveFootnotesEndnotesForConsolidatedRevisions(ConsolidationInfo ci, WordprocessingDocument wDocConsolidated)
        {
            var consolidatedFootnoteXDoc = wDocConsolidated.MainDocumentPart.FootnotesPart.GetXDocument();
            var consolidatedEndnoteXDoc = wDocConsolidated.MainDocumentPart.EndnotesPart.GetXDocument();

            int maxFootnoteId = 1;
            if (consolidatedFootnoteXDoc.Root.Elements(W.footnote).Any())
                maxFootnoteId = consolidatedFootnoteXDoc.Root.Elements(W.footnote).Select(e => (int)e.Attribute(W.id)).Max();
            int maxEndnoteId = 1;
            if (consolidatedEndnoteXDoc.Root.Elements(W.endnote).Any())
                maxEndnoteId = consolidatedEndnoteXDoc.Root.Elements(W.endnote).Select(e => (int)e.Attribute(W.id)).Max(); ;

            /// At this point, content might contain a footnote or endnote reference.
            /// Need to add the footnote / endnote into the consolidated document (with the same guid id)
            /// Because of preprocessing of the documents, all footnote and endnote references will be unique at this point

            if (ci.RevisionElement.Descendants(W.footnoteReference).Any())
            {
                var footnoteXDoc = wDocConsolidated.MainDocumentPart.FootnotesPart.GetXDocument();
                foreach (var footnoteReference in ci.RevisionElement.Descendants(W.footnoteReference))
                {
                    var id = (int)footnoteReference.Attribute(W.id);
                    var footnote = ci.Footnotes.FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                    var newId = maxFootnoteId + 1;
                    maxFootnoteId++;
                    footnoteReference.Attribute(W.id).Value = newId.ToString();
                    var clonedFootnote = new XElement(footnote);
                    clonedFootnote.Attribute(W.id).Value = newId.ToString();
                    footnoteXDoc.Root.Add(clonedFootnote);
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
                    var newId = maxEndnoteId + 1;
                    maxEndnoteId++;
                    endnoteReference.Attribute(W.id).Value = newId.ToString();
                    var clonedEndnote = new XElement(endnote);
                    clonedEndnote.Attribute(W.id).Value = newId.ToString();
                    endnoteXDoc.Root.Add(clonedEndnote);
                }
                wDocConsolidated.MainDocumentPart.EndnotesPart.PutXDocument();
            }
        }

        private static object CleanPartTransform(XNode node)
        {
            var element = node as XElement;
            if (element != null)
            {
                return new XElement(element.Name,
                    element.Attributes().Where(a => a.Name.Namespace != PtOpenXml.pt &&
                        !a.Name.LocalName.ToLower().Contains("rsid")),
                    element.Nodes().Select(n => CleanPartTransform(n)));
            }
            return node;
        }

        private static string RevisionToLogFormTransform(XElement element, int depth, bool inserting)
        {
            if (element.Name == W.p)
                return "Paragraph" + nl + element.Elements().Select(e => RevisionToLogFormTransform(e, depth + 2, false)).StringConcatenate();
            if (element.Name == W.pPr || element.Name == W.rPr)
                return "";
            if (element.Name == W.r)
                return element.Elements().Select(e => RevisionToLogFormTransform(e, depth, inserting)).StringConcatenate();
            if (element.Name == W.t)
            {
                if (inserting)
                    return "".PadRight(depth) + "Inserted Text:" + QuoteIt((string)element) + nl;
                else
                    return "".PadRight(depth) + "Text:" + QuoteIt((string)element) + nl;
            }
            if (element.Name == W.delText)
                return "".PadRight(depth) + "Deleted Text:" + QuoteIt((string)element) + nl;
            if (element.Name == W.ins)
                return element.Elements().Select(e => RevisionToLogFormTransform(e, depth, true)).StringConcatenate();
            if (element.Name == W.del)
                return element.Elements().Select(e => RevisionToLogFormTransform(e, depth, false)).StringConcatenate();
            return "";
        }

        private static string QuoteIt(string str)
        {
            var quoteString = "\"";
            if (str.Contains('\"'))
                quoteString = "\'";
            return quoteString + str + quoteString;
        }

        private static void IgnorePt14NamespaceForFootnotesEndnotes(WordprocessingDocument wDoc)
        {
            var footnotesPart = wDoc.MainDocumentPart.FootnotesPart;
            var endnotesPart = wDoc.MainDocumentPart.EndnotesPart;

            XDocument footnotesPartXDoc = null;
            if (footnotesPart != null)
            {
                footnotesPartXDoc = footnotesPart.GetXDocument();
                WmlComparer.IgnorePt14Namespace(footnotesPartXDoc.Root);
            }

            XDocument endnotesPartXDoc = null;
            if (endnotesPart != null)
            {
                endnotesPartXDoc = endnotesPart.GetXDocument();
                WmlComparer.IgnorePt14Namespace(endnotesPartXDoc.Root);
            }

            if (footnotesPart != null)
                footnotesPart.PutXDocument();

            if (endnotesPart != null)
                endnotesPart.PutXDocument();
        }

        private static XElement[] AssembledConjoinedRevisionContent(XElement emptyParagraph, IGrouping<string, ConsolidationInfo> groupedCi, int idx, WordprocessingDocument wDocConsolidated,
            WmlComparerConsolidateSettings consolidateSettings)
        {
            var consolidatedFootnoteXDoc = wDocConsolidated.MainDocumentPart.FootnotesPart.GetXDocument();
            var consolidatedEndnoteXDoc = wDocConsolidated.MainDocumentPart.EndnotesPart.GetXDocument();

            int maxFootnoteId = 1;
            if (consolidatedFootnoteXDoc.Root.Elements(W.footnote).Any())
                maxFootnoteId = consolidatedFootnoteXDoc.Root.Elements(W.footnote).Select(e => (int)e.Attribute(W.id)).Max();
            int maxEndnoteId = 1;
            if (consolidatedEndnoteXDoc.Root.Elements(W.endnote).Any())
                maxEndnoteId = consolidatedEndnoteXDoc.Root.Elements(W.endnote).Select(e => (int)e.Attribute(W.id)).Max(); ;

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
                colorString = colorString.Substring(2);

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
                                /// At this point, content might contain a footnote or endnote reference.
                                /// Need to add the footnote / endnote into the consolidated document (with the same guid id)
                                /// Because of preprocessing of the documents, all footnote and endnote references will be unique at this point

                                if (ci.RevisionElement.Descendants(W.endnoteReference).Any())
                                {
                                    var endnoteXDoc = wDocConsolidated.MainDocumentPart.EndnotesPart.GetXDocument();
                                    foreach (var endnoteReference in ci.RevisionElement.Descendants(W.endnoteReference))
                                    {
                                        var id = (int)endnoteReference.Attribute(W.id);
                                        var endnote = ci.Endnotes.FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                                        var newId = maxEndnoteId + 1;
                                        maxEndnoteId++;
                                        endnoteReference.Attribute(W.id).Value = newId.ToString();
                                        var clonedEndnote = new XElement(endnote);
                                        clonedEndnote.Attribute(W.id).Value = newId.ToString();
                                        endnoteXDoc.Root.Add(clonedEndnote);
                                    }
                                    wDocConsolidated.MainDocumentPart.EndnotesPart.PutXDocument();
                                }

                                if (ci.RevisionElement.Descendants(W.footnoteReference).Any())
                                {
                                    var footnoteXDoc = wDocConsolidated.MainDocumentPart.FootnotesPart.GetXDocument();
                                    foreach (var footnoteReference in ci.RevisionElement.Descendants(W.footnoteReference))
                                    {
                                        var id = (int)footnoteReference.Attribute(W.id);
                                        var footnote = ci.Footnotes.FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                                        var newId = maxFootnoteId + 1;
                                        maxFootnoteId++;
                                        footnoteReference.Attribute(W.id).Value = newId.ToString();
                                        var clonedFootnote = new XElement(footnote);
                                        clonedFootnote.Attribute(W.id).Value = newId.ToString();
                                        footnoteXDoc.Root.Add(clonedFootnote);
                                    }
                                    wDocConsolidated.MainDocumentPart.FootnotesPart.PutXDocument();
                                }

                                // it is important that this code follows the code above, because the code above updates ci.RevisionElement (using DML)

                                XElement paraAfter = null;
                                if (ci.RevisionElement.Name == W.tbl)
                                    paraAfter = emptyParagraph;
                                var revisionInTable = new[] {
                                    ci.RevisionElement,
                                    paraAfter,
                                    };

                                return revisionInTable;
                            }))));

                // if the last paragraph has a deleted paragraph mark, then remove the deletion from the paragraph mark.  This is to prevent Word from misbehaving.
                // the last paragraph in a cell must not have a deleted paragraph mark.
                var theCell = table
                    .Descendants(W.tc)
                    .FirstOrDefault();
                var lastPara = theCell
                    .Elements(W.p)
                    .LastOrDefault();
                if (lastPara != null)
                {
                    var isDeleted = lastPara
                        .Elements(W.pPr)
                        .Elements(W.rPr)
                        .Elements(W.del)
                        .Any();
                    if (isDeleted)
                        lastPara
                            .Elements(W.pPr)
                            .Elements(W.rPr)
                            .Elements(W.del)
                            .Remove();
                }

                var content = new[] {
                                    idx == 0 ? emptyParagraph : null,
                                    table,
                                    emptyParagraph,
                                };
								
                var dummyElement = new XElement("dummy", content);

                foreach (var rev in dummyElement.Descendants().Where(d => d.Attribute(W.author) != null))
                {
                    var aut = rev.Attribute(W.author);
                    aut.Value = revisor;
                }

                return dummyElement.Elements().ToArray();
            }
            else
            {
                var content = groupedCi.Select(ci =>
                {
                    XElement paraAfter = null;
                    if (ci.RevisionElement.Name == W.tbl)
                        paraAfter = emptyParagraph;
                    var revisionInTable = new[] {
                                    ci.RevisionElement,
                                    paraAfter,
                                    };

                    /// At this point, content might contain a footnote or endnote reference.
                    /// Need to add the footnote / endnote into the consolidated document (with the same guid id)
                    /// Because of preprocessing of the documents, all footnote and endnote references will be unique at this point

                    if (ci.RevisionElement.Descendants(W.footnoteReference).Any())
                    {
                        var footnoteXDoc = wDocConsolidated.MainDocumentPart.FootnotesPart.GetXDocument();
                        foreach (var footnoteReference in ci.RevisionElement.Descendants(W.footnoteReference))
                        {
                            var id = (int)footnoteReference.Attribute(W.id);
                            var footnote = ci.Footnotes.FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                            var newId = maxFootnoteId + 1;
                            maxFootnoteId++;
                            footnoteReference.Attribute(W.id).Value = newId.ToString();
                            var clonedFootnote = new XElement(footnote);
                            clonedFootnote.Attribute(W.id).Value = newId.ToString();
                            footnoteXDoc.Root.Add(clonedFootnote);
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
                            var newId = maxEndnoteId + 1;
                            maxEndnoteId++;
                            endnoteReference.Attribute(W.id).Value = newId.ToString();
                            var clonedEndnote = new XElement(endnote);
                            clonedEndnote.Attribute(W.id).Value = newId.ToString();
                            endnoteXDoc.Root.Add(clonedEndnote);
                        }
                        wDocConsolidated.MainDocumentPart.EndnotesPart.PutXDocument();
                    }

                    return revisionInTable;
                });

                var dummyElement = new XElement("dummy",
                    content.SelectMany(m => m));

                foreach (var rev in dummyElement.Descendants().Where(d => d.Attribute(W.author) != null))
                {
                    var aut = rev.Attribute(W.author);
                    aut.Value = revisor;
                }

                return dummyElement.Elements().ToArray();
            }
        }

        private static void AddToAnnotation(
            WordprocessingDocument wDocDelta,
            WordprocessingDocument consolidatedWDoc,
            XElement elementToInsertAfter,
            ConsolidationInfo consolidationInfo,
            WmlComparerSettings settings)
        {
            Package packageOfDeletedContent = wDocDelta.MainDocumentPart.OpenXmlPackage.Package;
            Package packageOfNewContent = consolidatedWDoc.MainDocumentPart.OpenXmlPackage.Package;
            PackagePart partInDeletedDocument = packageOfDeletedContent.GetPart(wDocDelta.MainDocumentPart.Uri);
            PackagePart partInNewDocument = packageOfNewContent.GetPart(consolidatedWDoc.MainDocumentPart.Uri);
            consolidationInfo.RevisionElement = MoveRelatedPartsToDestination(partInDeletedDocument, partInNewDocument, consolidationInfo.RevisionElement);

            var clonedForHashing = (XElement)CloneBlockLevelContentForHashing(consolidatedWDoc.MainDocumentPart, consolidationInfo.RevisionElement, false, settings);
            clonedForHashing.Descendants().Where(d => d.Name == W.ins || d.Name == W.del).Attributes(W.id).Remove();
            var shaString = clonedForHashing.ToString(SaveOptions.DisableFormatting)
                .Replace(" xmlns=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", "");
            var sha1Hash = PtUtils.SHA1HashStringForUTF8String(shaString);
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

        private static void ChangeFootnoteEndnoteReferencesToUniqueRange(WordprocessingDocument wDoc, int startingIdForFootnotesEndnotes)
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

            var rnd = new Random();
            foreach (var r in references)
            {
                var oldId = (string)r.Attribute(W.id);
                var newId = startingIdForFootnotesEndnotes.ToString();
                startingIdForFootnotesEndnotes++;
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
            // save away sectPr so that can set in the newly produced document.
            var savedSectPr = wDoc1
                .MainDocumentPart
                .GetXDocument()
                .Root
                .Element(W.body)
                .Element(W.sectPr);

            var contentParent1 = wDoc1.MainDocumentPart.GetXDocument().Root.Element(W.body);
            AddSha1HashToBlockLevelContent(wDoc1.MainDocumentPart, contentParent1, settings);
            var contentParent2 = wDoc2.MainDocumentPart.GetXDocument().Root.Element(W.body);
            AddSha1HashToBlockLevelContent(wDoc2.MainDocumentPart, contentParent2, settings);

            var cal1 = WmlComparer.CreateComparisonUnitAtomList(wDoc1.MainDocumentPart, wDoc1.MainDocumentPart.GetXDocument().Root.Element(W.body), settings);

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in cal1)
                    sb.Append(item.ToString() + Environment.NewLine);
                var sbs = sb.ToString();
                DocxComparerUtil.NotePad(sbs);
            }

            var cus1 = GetComparisonUnitList(cal1, settings);

            if (s_False)
            {
                var sbs = ComparisonUnit.ComparisonUnitListToString(cus1);
                DocxComparerUtil.NotePad(sbs);
            }

            var cal2 = WmlComparer.CreateComparisonUnitAtomList(wDoc2.MainDocumentPart, wDoc2.MainDocumentPart.GetXDocument().Root.Element(W.body), settings);

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in cal2)
                    sb.Append(item.ToString() + Environment.NewLine);
                var sbs = sb.ToString();
                DocxComparerUtil.NotePad(sbs);
            }

            var cus2 = GetComparisonUnitList(cal2, settings);

            if (s_False)
            {
                var sbs = ComparisonUnit.ComparisonUnitListToString(cus2);
                DocxComparerUtil.NotePad(sbs);
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
                DocxComparerUtil.NotePad(sbs3);
            }

            // if cus1 and cus2 have completely different content, then just return the first document deleted, and the second document inserted.
            List<CorrelatedSequence> correlatedSequence = null;

            correlatedSequence = DetectUnrelatedSources(cus1, cus2, settings);

            if (correlatedSequence == null)
                correlatedSequence = Lcs(cus1, cus2, settings);

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in correlatedSequence)
                    sb.Append(item.ToString() + Environment.NewLine);
                var sbs = sb.ToString();
                DocxComparerUtil.NotePad(sbs);
            }

            // for any deleted or inserted rows, we go into the w:trPr properties, and add the appropriate w:ins or w:del element, and therefore
            // when generating the document, the appropriate row will be marked as deleted or inserted.
            MarkRowsAsDeletedOrInserted(settings, correlatedSequence);

            // the following gets a flattened list of ComparisonUnitAtoms, with status indicated in each ComparisonUnitAtom: Deleted, Inserted, or Equal
            var listOfComparisonUnitAtoms = FlattenToComparisonUnitAtomList(correlatedSequence, settings);

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in listOfComparisonUnitAtoms)
                    sb.Append(item.ToString() + Environment.NewLine);
                var sbs = sb.ToString();
                DocxComparerUtil.NotePad(sbs);
            }

            // note - we don't want to do the hack until after flattening all of the groups.  At the end of the flattening, we should simply
            // have a list of ComparisonUnitAtoms, appropriately marked as equal, inserted, or deleted.

            // the table id will be hacked in the normal course of events.
            // in the case where a row is deleted, not necessary to hack - the deleted row ID will do.
            // in the case where a row is inserted, not necessary to hack - the inserted row ID will do as well.
            AssembleAncestorUnidsInOrderToRebuildXmlTreeProperly(listOfComparisonUnitAtoms);

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in listOfComparisonUnitAtoms)
                    sb.Append(item.ToStringAncestorUnids() + Environment.NewLine);
                var sbs = sb.ToString();
                DocxComparerUtil.NotePad(sbs);
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

                    // move w:sectPr from source document into newly generated document.
                    if (savedSectPr != null)
                    {
                        var xd = wDocWithRevisions.MainDocumentPart.GetXDocument();
                        // add everything but headers/footers
                        var clonedSectPr = new XElement(W.sectPr,
                            savedSectPr.Attributes(),
                            savedSectPr.Element(W.type),
                            savedSectPr.Element(W.pgSz),
                            savedSectPr.Element(W.pgMar),
                            savedSectPr.Element(W.cols),
                            savedSectPr.Element(W.titlePg));
                        xd.Root.Element(W.body).Add(clonedSectPr);
                    }
                    /**********************************************************************************************/

                    wDocWithRevisions.MainDocumentPart.PutXDocument();
                    FixUpFootnotesEndnotesWithCustomMarkers(wDocWithRevisions);
                    FixUpRevMarkIds(wDocWithRevisions);
                    FixUpDocPrIds(wDocWithRevisions);
                    FixUpShapeIds(wDocWithRevisions);
                    FixUpShapeTypeIds(wDocWithRevisions);
                    AddFootnotesEndnotesStyles(wDocWithRevisions);
                    CopyMissingStylesFromOneDocToAnother(wDoc2, wDocWithRevisions);
                    DeleteFootnotePropertiesInSettings(wDocWithRevisions);
                }
                foreach (var part in wDoc1.ContentParts())
                    part.PutXDocument();
                foreach (var part in wDoc2.ContentParts())
                    part.PutXDocument();
                var updatedWmlResult = new WmlDocument("Dummy.docx", ms.ToArray());
                return updatedWmlResult;
            }
        }

        private static void DeleteFootnotePropertiesInSettings(WordprocessingDocument wDocWithRevisions)
        {
            var settingsPart = wDocWithRevisions.MainDocumentPart.DocumentSettingsPart;
            if (settingsPart != null)
            {
                var sxDoc = settingsPart.GetXDocument();
                sxDoc.Root.Elements().Where(e => e.Name == W.footnotePr || e.Name == W.endnotePr).Remove();
                settingsPart.PutXDocument();
            }
        }

        private static void FixUpFootnotesEndnotesWithCustomMarkers(WordprocessingDocument wDocWithRevisions)
        {
#if FALSE
      // this needs to change
      <w:del w:author="Open-Xml-PowerTools"
             w:id="7"
             w:date="2017-06-07T12:23:22.8601285-07:00">
        <w:r>
          <w:rPr pt14:Unid="ec75a71361c84562a757eee8b28fc229">
            <w:rFonts w:cs="Times New Roman Bold"
                      pt14:Unid="16bb355df5964ba09854f9152c97242b" />
            <w:b w:val="0"
                 pt14:Unid="9abcec54ad414791a5627cbb198e8aa9" />
            <w:bCs pt14:Unid="71ecd2eba85e4bfaa92b3d618e2f8829" />
            <w:position w:val="6"
                        pt14:Unid="61793f6a5f494700b7f2a3a753ce9055" />
            <w:sz w:val="16"
                  pt14:Unid="60b3cd020c214d0ea07e5a68ae0e4efe" />
            <w:szCs w:val="16"
                    pt14:Unid="9ae61a724de44a75868180aac44ea380" />
          </w:rPr>
          <w:footnoteReference w:customMarkFollows="1"
                               w:id="1"
                               pt14:Status="Deleted" />
        </w:r>
      </w:del>
      <w:del w:author="Open-Xml-PowerTools"
             w:id="8"
             w:date="2017-06-07T12:23:22.8601285-07:00">
        <w:r>
          <w:rPr pt14:Unid="445caef74a624e588e7adaa6d7775639">
            <w:rFonts w:cs="Times New Roman Bold"
                      pt14:Unid="5920885f8ec44c53bcaece2de7eafda2" />
            <w:b w:val="0"
                 pt14:Unid="023a29e2e6d44c3b8c5df47317ace4c6" />
            <w:bCs pt14:Unid="e96e37daf9174b268ef4731df831df7d" />
            <w:position w:val="6"
                        pt14:Unid="be3f8ff7ed0745ae9340bb2706b28b1f" />
            <w:sz w:val="16"
                  pt14:Unid="6fbbde024e7c46b9b72435ae50065459" />
            <w:szCs w:val="16"
                    pt14:Unid="cc82e7bd75f441f2b609eae0672fb285" />
          </w:rPr>
          <w:delText>1</w:delText>
        </w:r>
      </w:del>

      // to this
      <w:del w:author="Open-Xml-PowerTools"
             w:id="7"
             w:date="2017-06-07T12:23:22.8601285-07:00">
        <w:r>
          <w:rPr pt14:Unid="ec75a71361c84562a757eee8b28fc229">
            <w:rFonts w:cs="Times New Roman Bold"
                      pt14:Unid="16bb355df5964ba09854f9152c97242b" />
            <w:b w:val="0"
                 pt14:Unid="9abcec54ad414791a5627cbb198e8aa9" />
            <w:bCs pt14:Unid="71ecd2eba85e4bfaa92b3d618e2f8829" />
            <w:position w:val="6"
                        pt14:Unid="61793f6a5f494700b7f2a3a753ce9055" />
            <w:sz w:val="16"
                  pt14:Unid="60b3cd020c214d0ea07e5a68ae0e4efe" />
            <w:szCs w:val="16"
                    pt14:Unid="9ae61a724de44a75868180aac44ea380" />
          </w:rPr>
          <w:footnoteReference w:customMarkFollows="1"
                               w:id="1"
                               pt14:Status="Deleted" />
          <w:delText>1</w:delText>
        </w:r>
      </w:del>
#endif
            // this is pretty random - a bug in Word prevents display of a document if the delText element does not immediately follow the footnoteReference element, in the same run.
            var mainXDoc = wDocWithRevisions.MainDocumentPart.GetXDocument();
            var newRoot = (XElement)FootnoteEndnoteReferenceCleanupTransform(mainXDoc.Root);
            mainXDoc.Root.ReplaceWith(newRoot);
            wDocWithRevisions.MainDocumentPart.PutXDocument();
        }

        private static object FootnoteEndnoteReferenceCleanupTransform(XNode node)
        {
            var element = node as XElement;
            if (element != null)
            {
                // small optimization to eliminate the work for most elements
                if (element.Element(W.del) != null || element.Element(W.ins) != null)
                {
                    var hasFootnoteEndnoteReferencesThatNeedCleanedUp = element
                        .Elements()
                        .Where(e => e.Name == W.del || e.Name == W.ins)
                        .Elements(W.r)
                        .Elements()
                        .Where(e => e.Name == W.footnoteReference || e.Name == W.endnoteReference)
                        .Attributes(W.customMarkFollows)
                        .Any();

                    if (hasFootnoteEndnoteReferencesThatNeedCleanedUp)
                    {
                        var clone = new XElement(element.Name,
                            element.Attributes(),
                            element.Nodes().Select(n => FootnoteEndnoteReferenceCleanupTransform(n)));
                        var footnoteEndnoteReferencesToAdjust = clone
                            .Descendants()
                            .Where(d => d.Name == W.footnoteReference || d.Name == W.endnoteReference)
                            .Where(d => d.Attribute(W.customMarkFollows) != null);
                        foreach (var fnenr in footnoteEndnoteReferencesToAdjust)
                        {
                            var par = fnenr.Parent;
                            var gp = fnenr.Parent.Parent;
                            if (par.Name == W.r &&
                                gp.Name == W.del)
                            {
                                if (par.Element(W.delText) != null)
                                    continue;
                                var afterGp = gp.ElementsAfterSelf().FirstOrDefault();
                                if (afterGp == null)
                                    continue;
                                var afterGpDelText = afterGp.Elements(W.r).Elements(W.delText);
                                if (afterGpDelText.Any())
                                {
                                    par.Add(afterGpDelText);  // this will clone and add to run that contains the reference
                                    afterGpDelText.Remove();  // this leaves an empty run, does not matter.
                                }
                            }
                            if (par.Name == W.r &&
                                gp.Name == W.ins)
                            {
                                if (par.Element(W.t) != null)
                                    continue;
                                var afterGp = gp.ElementsAfterSelf().FirstOrDefault();
                                if (afterGp == null)
                                    continue;
                                var afterGpText = afterGp.Elements(W.r).Elements(W.t);
                                if (afterGpText.Any())
                                {
                                    par.Add(afterGpText);  // this will clone and add to run that contains the reference
                                    afterGpText.Remove();  // this leaves an empty run, does not matter.
                                }
                            }
                        }
                        return clone;
                    }
                }
                else
                {
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Nodes().Select(n => FootnoteEndnoteReferenceCleanupTransform(n)));
                }
            }
            return node;
        }

        private static void CopyMissingStylesFromOneDocToAnother(WordprocessingDocument wDocFrom, WordprocessingDocument wDocTo)
        {
            var revisionsStylesXDoc = wDocTo.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
            var afterStylesXDoc = wDocFrom.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
            foreach (var style in afterStylesXDoc.Root.Elements(W.style))
            {
                var type = (string)style.Attribute(W.type);
                var styleId = (string)style.Attribute(W.styleId);
                var styleInRevDoc = revisionsStylesXDoc
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(st => (string)st.Attribute(W.type) == type &&
                                          (string)st.Attribute(W.styleId) == styleId);
                if (styleInRevDoc != null)
                    continue;
                var cloned = new XElement(style);
                if (cloned.Attribute(W._default) != null)
                    cloned.Attribute(W._default).Remove();
                revisionsStylesXDoc.Root.Add(cloned);
            }
            wDocTo.MainDocumentPart.StyleDefinitionsPart.PutXDocument();
        }

        private static void AddFootnotesEndnotesStyles(WordprocessingDocument wDocWithRevisions)
        {
            var mainXDoc = wDocWithRevisions.MainDocumentPart.GetXDocument();
            var hasFootnotes = mainXDoc.Descendants(W.footnoteReference).Any();
            var hasEndnotes = mainXDoc.Descendants(W.endnoteReference).Any();
            var styleDefinitionsPart = wDocWithRevisions.MainDocumentPart.StyleDefinitionsPart;
            var sXDoc = styleDefinitionsPart.GetXDocument();
            if (hasFootnotes)
            {
                var footnoteTextStyle = sXDoc
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(s => (string)s.Attribute(W.styleId) == "FootnoteText");
                if (footnoteTextStyle == null)
                {
                    var footnoteTextStyleMarkup =
@"<w:style w:type=""paragraph""
           w:styleId=""FootnoteText""
           xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:name w:val=""footnote text""/>
    <w:basedOn w:val=""Normal""/>
    <w:link w:val=""FootnoteTextChar""/>
    <w:uiPriority w:val=""99""/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:pPr>
      <w:spacing w:after=""0""
                 w:line=""240""
                 w:lineRule=""auto""/>
    </w:pPr>
    <w:rPr>
      <w:sz w:val=""20""/>
      <w:szCs w:val=""20""/>
    </w:rPr>
  </w:style>";
                    var ftsElement = XElement.Parse(footnoteTextStyleMarkup);
                    sXDoc.Root.Add(ftsElement);
                }
                var footnoteTextCharStyle = sXDoc
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(s => (string)s.Attribute(W.styleId) == "FootnoteTextChar");
                if (footnoteTextCharStyle == null)
                {
                    var footnoteTextCharStyleMarkup =
@"<w:style w:type=""character""
           w:customStyle=""1""
           w:styleId=""FootnoteTextChar""
           xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:name w:val=""Footnote Text Char""/>
    <w:basedOn w:val=""DefaultParagraphFont""/>
    <w:link w:val=""FootnoteText""/>
    <w:uiPriority w:val=""99""/>
    <w:semiHidden/>
    <w:rPr>
      <w:sz w:val=""20""/>
      <w:szCs w:val=""20""/>
    </w:rPr>
  </w:style>";
                    var fntcsElement = XElement.Parse(footnoteTextCharStyleMarkup);
                    sXDoc.Root.Add(fntcsElement);
                }
                var footnoteReferenceStyle = sXDoc
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(s => (string)s.Attribute(W.styleId) == "FootnoteReference");
                if (footnoteReferenceStyle == null)
                {
                    var footnoteReferenceStyleMarkup =
@"<w:style w:type=""character""
           w:styleId=""FootnoteReference""
           xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:name w:val=""footnote reference""/>
    <w:basedOn w:val=""DefaultParagraphFont""/>
    <w:uiPriority w:val=""99""/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:rPr>
      <w:vertAlign w:val=""superscript""/>
    </w:rPr>
  </w:style>";
                    var fnrsElement = XElement.Parse(footnoteReferenceStyleMarkup);
                    sXDoc.Root.Add(fnrsElement);
                }
            }
            if (hasEndnotes)
            {
                var endnoteTextStyle = sXDoc
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(s => (string)s.Attribute(W.styleId) == "EndnoteText");
                if (endnoteTextStyle == null)
                {
                    var endnoteTextStyleMarkup =
@"<w:style w:type=""paragraph""
           w:styleId=""EndnoteText""
           xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:name w:val=""endnote text""/>
    <w:basedOn w:val=""Normal""/>
    <w:link w:val=""EndnoteTextChar""/>
    <w:uiPriority w:val=""99""/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:pPr>
      <w:spacing w:after=""0""
                 w:line=""240""
                 w:lineRule=""auto""/>
    </w:pPr>
    <w:rPr>
      <w:sz w:val=""20""/>
      <w:szCs w:val=""20""/>
    </w:rPr>
  </w:style>";
                    var etsElement = XElement.Parse(endnoteTextStyleMarkup);
                    sXDoc.Root.Add(etsElement);
                }
                var endnoteTextCharStyle = sXDoc
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(s => (string)s.Attribute(W.styleId) == "EndnoteTextChar");
                if (endnoteTextCharStyle == null)
                {
                    var endnoteTextCharStyleMarkup =
@"<w:style w:type=""character""
           w:customStyle=""1""
           w:styleId=""EndnoteTextChar""
           xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:name w:val=""Endnote Text Char""/>
    <w:basedOn w:val=""DefaultParagraphFont""/>
    <w:link w:val=""EndnoteText""/>
    <w:uiPriority w:val=""99""/>
    <w:semiHidden/>
    <w:rPr>
      <w:sz w:val=""20""/>
      <w:szCs w:val=""20""/>
    </w:rPr>
  </w:style>";
                    var entcsElement = XElement.Parse(endnoteTextCharStyleMarkup);
                    sXDoc.Root.Add(entcsElement);
                }
                var endnoteReferenceStyle = sXDoc
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(s => (string)s.Attribute(W.styleId) == "EndnoteReference");
                if (endnoteReferenceStyle == null)
                {
                    var endnoteReferenceStyleMarkup =
@"<w:style w:type=""character""
           w:styleId=""EndnoteReference""
           xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:name w:val=""endnote reference""/>
    <w:basedOn w:val=""DefaultParagraphFont""/>
    <w:uiPriority w:val=""99""/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:rPr>
      <w:vertAlign w:val=""superscript""/>
    </w:rPr>
  </w:style>";
                    var enrsElement = XElement.Parse(endnoteReferenceStyleMarkup);
                    sXDoc.Root.Add(enrsElement);
                }
            }
            if (hasFootnotes || hasEndnotes)
            {
                styleDefinitionsPart.PutXDocument();
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
                    AddSha1HashToBlockLevelContent(partToUseBefore, footnoteEndnoteBefore, settings);
                    AddSha1HashToBlockLevelContent(partToUseAfter, footnoteEndnoteAfter, settings);

                    var fncal1 = WmlComparer.CreateComparisonUnitAtomList(partToUseBefore, footnoteEndnoteBefore, settings);
                    var fncus1 = GetComparisonUnitList(fncal1, settings);

                    var fncal2 = WmlComparer.CreateComparisonUnitAtomList(partToUseAfter, footnoteEndnoteAfter, settings);
                    var fncus2 = GetComparisonUnitList(fncal2, settings);

                    if (!(fncus1.Length == 0 && fncus2.Length == 0))
                    {
                        var fnCorrelatedSequence = Lcs(fncus1, fncus2, settings);

                        if (s_False)
                        {
                            var sb = new StringBuilder();
                            foreach (var item in fnCorrelatedSequence)
                                sb.Append(item.ToString()).Append(Environment.NewLine);
                            var sbs = sb.ToString();
                            DocxComparerUtil.NotePad(sbs);
                        }

                        // for any deleted or inserted rows, we go into the w:trPr properties, and add the appropriate w:ins or w:del element, and therefore
                        // when generating the document, the appropriate row will be marked as deleted or inserted.
                        MarkRowsAsDeletedOrInserted(settings, fnCorrelatedSequence);

                        // the following gets a flattened list of ComparisonUnitAtoms, with status indicated in each ComparisonUnitAtom: Deleted, Inserted, or Equal
                        var fnListOfComparisonUnitAtoms = FlattenToComparisonUnitAtomList(fnCorrelatedSequence, settings);

                        if (s_False)
                        {
                            var sb = new StringBuilder();
                            foreach (var item in fnListOfComparisonUnitAtoms)
                                sb.Append(item.ToString() + Environment.NewLine);
                            var sbs = sb.ToString();
                            DocxComparerUtil.NotePad(sbs);
                        }

                        // hack = set the guid ID of the table, row, or cell from the 'before' document to be equal to the 'after' document.

                        // note - we don't want to do the hack until after flattening all of the groups.  At the end of the flattening, we should simply
                        // have a list of ComparisonUnitAtoms, appropriately marked as equal, inserted, or deleted.

                        // the table id will be hacked in the normal course of events.
                        // in the case where a row is deleted, not necessary to hack - the deleted row ID will do.
                        // in the case where a row is inserted, not necessary to hack - the inserted row ID will do as well.
                        AssembleAncestorUnidsInOrderToRebuildXmlTreeProperly(fnListOfComparisonUnitAtoms);

                        var newFootnoteEndnoteChildren = ProduceNewWmlMarkupFromCorrelatedSequence(partToUseAfter, fnListOfComparisonUnitAtoms, settings);
                        var tempElement = new XElement(W.body, newFootnoteEndnoteChildren);
                        var hasFootnoteReference = tempElement.Descendants(W.r).Any(r =>
                        {
                            var b = false;
                            if ((string)r.Elements(W.rPr).Elements(W.rStyle).Attributes(W.val).FirstOrDefault() == "FootnoteReference")
                                b = true;
                            if (r.Descendants(W.footnoteRef).Any())
                                b = true;
                            return b;
                        });
                        if (!hasFootnoteReference)
                        {
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
                        }
                        XElement newTempElement = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(tempElement);
                        var newContentElement = newTempElement.Descendants().FirstOrDefault(d => d.Name == W.footnote || d.Name == W.endnote);
                        if (newContentElement == null)
                            throw new OpenXmlPowerToolsException("Internal error");
                        footnoteEndnoteAfter.ReplaceNodes(newContentElement.Nodes());
                    }
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

                    AddSha1HashToBlockLevelContent(partToUseAfter, footnoteEndnoteAfter, settings);

                    var fncal2 = WmlComparer.CreateComparisonUnitAtomList(partToUseAfter, footnoteEndnoteAfter, settings);
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
                        DocxComparerUtil.NotePad(sbs);
                    }

                    MarkRowsAsDeletedOrInserted(settings, insertedCorrSequ);

                    var fnListOfComparisonUnitAtoms = FlattenToComparisonUnitAtomList(insertedCorrSequ, settings);

                    AssembleAncestorUnidsInOrderToRebuildXmlTreeProperly(fnListOfComparisonUnitAtoms);

                    var newFootnoteEndnoteChildren = ProduceNewWmlMarkupFromCorrelatedSequence(partToUseAfter,
                        fnListOfComparisonUnitAtoms, settings);
                    var tempElement = new XElement(W.body, newFootnoteEndnoteChildren);
                    var hasFootnoteReference = tempElement.Descendants(W.r).Any(r =>
                    {
                        var b = false;
                        if ((string)r.Elements(W.rPr).Elements(W.rStyle).Attributes(W.val).FirstOrDefault() == "FootnoteReference")
                            b = true;
                        if (r.Descendants(W.footnoteRef).Any())
                            b = true;
                        return b;
                    });
                    if (!hasFootnoteReference)
                    {
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
                    }
                    XElement newTempElement = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(tempElement);
                    var newContentElement = newTempElement
                        .Descendants()
                        .FirstOrDefault(d => d.Name == W.footnote || d.Name == W.endnote);
                    if (newContentElement != null)
                    {     //throw new OpenXmlPowerToolsException("Internal error");
                        footnoteEndnoteAfter.ReplaceNodes(newContentElement.Nodes());
                    }
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

                    AddSha1HashToBlockLevelContent(partToUseBefore, footnoteEndnoteBefore, settings);

                    var fncal2 = WmlComparer.CreateComparisonUnitAtomList(partToUseBefore, footnoteEndnoteBefore, settings);
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
                        DocxComparerUtil.NotePad(sbs);
                    }

                    MarkRowsAsDeletedOrInserted(settings, deletedCorrSequ);

                    var fnListOfComparisonUnitAtoms = FlattenToComparisonUnitAtomList(deletedCorrSequ, settings);

                    if (fnListOfComparisonUnitAtoms.Any())
                    {
                        AssembleAncestorUnidsInOrderToRebuildXmlTreeProperly(fnListOfComparisonUnitAtoms);

                        var newFootnoteEndnoteChildren = ProduceNewWmlMarkupFromCorrelatedSequence(partToUseBefore,
                            fnListOfComparisonUnitAtoms, settings);
                        var tempElement = new XElement(W.body, newFootnoteEndnoteChildren);
                        var hasFootnoteReference = tempElement.Descendants(W.r).Any(r =>
                        {
                            var b = false;
                            if ((string)r.Elements(W.rPr).Elements(W.rStyle).Attributes(W.val).FirstOrDefault() == "FootnoteReference")
                                b = true;
                            if (r.Descendants(W.footnoteRef).Any())
                                b = true;
                            return b;
                        });
                        if (!hasFootnoteReference)
                        {
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
                        }
                        XElement newTempElement = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(tempElement);
                        var newContentElement = newTempElement.Descendants().FirstOrDefault(d => d.Name == W.footnote || d.Name == W.endnote);
                        if (newContentElement == null)
                            throw new OpenXmlPowerToolsException("Internal error");
                        footnoteEndnoteBefore.ReplaceNodes(newContentElement.Nodes());
                    }
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
                var cloned = new XElement(footnote);
                cloned.Attribute(W.id).Value = newId;
                footnotesPartWithRevisionsXDoc
                    .Root
                    .Add(cloned);
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
                var cloned = new XElement(endnote);
                cloned.Attribute(W.id).Value = newId;
                endnotesPartWithRevisionsXDoc
                    .Root
                    .Add(cloned);
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

        /// Here is the crux of the fix to the algorithm.  After assembling the entire list of ComparisonUnitAtoms, we do the following:
        /// - First, figure out the maximum hierarchy depth, considering only paragraphs, txbx, txbxContent, tables, rows, cells, and content controls.
        /// - For documents that do not contain tables, nor text boxes, this maximum hierarchy depth will always be 1.
        /// - For atoms within a table, the depth will be 4.  The first level is the table, the second level is row, third is cell, fourth is paragraph.
        /// - For atoms within a nested table, the depth will be 7:  Table / Row / Cell / Table / Row / Cell / Paragraph
        /// - For atoms within a text box, the depth will be 3: Paragraph / txbxContent / Paragraph
        /// - For atoms within a table in a text box, the depth will be 5:  Paragraph / txbxContent / Table / Row / Cell / Paragraph
        /// In any case, we figure out the maximum depth.
        /// 
        /// Then we iterate through the list of content atoms backwards.  We do this n times, where n is the maximum depth.
        /// 
        /// At each level, we find a paragraph mark, and working backwards, we set the guids in the hierarchy so that the content will be assembled together correctly.
        /// 
        /// For each iteration, we only set unids at the level that we are working at.
        /// 
        /// So first we will set all unids at level 1.  When we find a paragraph mark, we get the unid for that level, and then working backwards, until we find another
        /// paragraph mark, we set all unids at level 1 to the same unid as level 1 of the paragraph mark.
        /// 
        /// Then we set all unids at level 2.  When we find a paragraph mark, we get the unid for that level, and then working backwards, until we find another paragraph
        /// mark, we set all unids at level 2 to the same unid as level 2 of the paragraph mark.  At some point, we will find a paragraph mark with no level 2.  This is
        /// not a problem.  We stop setting anything until we find another paragraph mark that has a level 2, at which point we resume setting values at level 2.
        /// 
        /// Same process for level 3, and so on, until we have processed to the maximum depth of the hierarchy.
        /// 
        /// At the end of this process, we will be able to do the coalsce recurse algorithm, and the content atom list will be put back together into a beautiful tree,
        /// where every element is correctly positioned in the hierarchy.
        /// 
        /// This should also properly assemble the test where just the paragraph marks have been deleted for a range of paragraphs.
        ///
        /// There is an interesting thought - it is possible that I have set two runs of text that were initially in the same paragraph, but then after
        /// processing, they match up to text in different paragraphs.  Therefore this will not work.  We need to actually keep a list of reconstructed ancestor
        /// Unids, because the same paragraph would get set to two different IDs - two ComparisonUnitAtoms need to be in separate paragraphs in the reconstructed
        /// document, but their ancestors actually point to the same paragraph.
        /// 
        /// Fix this in the algorithm, and also keep the appropriate list in ComparisonUnitAtom class.

        private static void AssembleAncestorUnidsInOrderToRebuildXmlTreeProperly(List<ComparisonUnitAtom> comparisonUnitAtomList)
        {
            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in comparisonUnitAtomList)
                    sb.Append(item.ToString()).Append(Environment.NewLine);
                var sbs = sb.ToString();
                DocxComparerUtil.NotePad(sbs);
            }

            // the following loop sets all ancestor unids in the after document to the unids in the before document for all pPr where the status is equal.
            // this should always be true.

            // one additional modification to make to this loop - where we find a pPr in a text box, we want to do this as well, regardless of whether the status is equal, inserted, or deleted.
            // reason being that this module does not support insertion / deletion of text boxes themselves.  If a text box is in the before or after document, it will be in the document that
            // contains deltas.  It may have inserted or deleted text, but regardless, it will be in the result document.
            foreach (var cua in comparisonUnitAtomList)
            {
                var doSet = false;
                if (cua.ContentElement.Name == W.pPr)
                {
                    if (cua.AncestorElements.Any(ae => ae.Name == W.txbxContent))
                        doSet = true;
                    if (cua.CorrelationStatus == CorrelationStatus.Equal)
                        doSet = true;
                }
                if (doSet)
                {
                    var cuaBefore = cua.ComparisonUnitAtomBefore;
                    var ancestorsAfter = cua.AncestorElements;
                    if (cuaBefore != null)
                    {
                        var ancestorsBefore = cuaBefore.AncestorElements;
                        if (ancestorsAfter.Length == ancestorsBefore.Length)
                        {
                            var zipped = ancestorsBefore.Zip(ancestorsAfter, (b, a) =>
                                new
                                {
                                    After = a,
                                    Before = b,
                                });

                            foreach (var z in zipped)
                            {
                                var afterUnidAtt = z.After.Attribute(PtOpenXml.Unid);
                                var beforeUnidAtt = z.Before.Attribute(PtOpenXml.Unid);
                                if (afterUnidAtt != null && beforeUnidAtt != null)
                                    afterUnidAtt.Value = beforeUnidAtt.Value;
                            }
                        }
                    }
                }
            }

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in comparisonUnitAtomList)
                    sb.Append(item.ToString()).Append(Environment.NewLine);
                var sbs = sb.ToString();
                DocxComparerUtil.NotePad(sbs);
            }

            var rComparisonUnitAtomList = ((IEnumerable<ComparisonUnitAtom>)comparisonUnitAtomList).Reverse().ToList();

            // the following should always succeed, because there will always be at least one element in rComparisonUnitAtomList, and there will always be at least one
            // ancestor in AncestorElements
            string deepestAncestorUnid = null;
            if (rComparisonUnitAtomList.Any())
            {
                var deepestAncestor = rComparisonUnitAtomList.First().AncestorElements.First();
                var deepestAncestorName = deepestAncestor.Name;

                if (deepestAncestorName == W.footnote || deepestAncestorName == W.endnote)
                {
                    deepestAncestorUnid = (string)deepestAncestor.Attribute(PtOpenXml.Unid);
                }
            }

            /// If the following loop finds a pPr that is in a text box, then continue on, processing the pPr and all of its contents as though it were
            /// content in the containing text box.  This is going to leave it after this loop where the AncestorUnids for the content in the text box will be
            /// incomplete.  We then will need to go through the rComparisonUnitAtomList a second time, processing all of the text boxes.

            /// Note that this makes the basic assumption that a text box can't be nested inside of a text box, which, as far as I know, is a good assumption.

            /// This also makes the basic assumption that an endnote / footnote can't contain a text box, which I believe is a good assumption.


            string[] currentAncestorUnids = null;
            foreach (var cua in rComparisonUnitAtomList)
            {
                if (cua.ContentElement.Name == W.pPr)
                {
                    var pPr_inTextBox = cua
                        .AncestorElements
                        .Any(ae => ae.Name == W.txbxContent);

                    if (!pPr_inTextBox)
                    {
                        // this will collect the ancestor unids for the paragraph.
                        // my hypothesis is that these ancestor unids should be the same for all content unit atoms within that paragraph.
                        currentAncestorUnids = cua
                            .AncestorElements
                            .Select(ae =>
                            {
                                var thisUnid = (string)ae.Attribute(PtOpenXml.Unid);
                                if (thisUnid == null)
                                    throw new OpenXmlPowerToolsException("Internal error");
                                return thisUnid;
                            })
                            .ToArray();
                        cua.AncestorUnids = currentAncestorUnids;
                        if (deepestAncestorUnid != null)
                            cua.AncestorUnids[0] = deepestAncestorUnid;
                        continue;
                    }
                }

                var thisDepth = cua.AncestorElements.Length;
                var additionalAncestorUnids = cua
                    .AncestorElements
                    .Skip(currentAncestorUnids.Length)
                    .Select(ae =>
                    {
                        var thisUnid = (string)ae.Attribute(PtOpenXml.Unid);
                        if (thisUnid == null)
                            Guid.NewGuid().ToString().Replace("-", "");
                        return thisUnid;
                    });
                var thisAncestorUnids = currentAncestorUnids
                    .Concat(additionalAncestorUnids)
                    .ToArray();
                cua.AncestorUnids = thisAncestorUnids;
                if (deepestAncestorUnid != null)
                    cua.AncestorUnids[0] = deepestAncestorUnid;
            }

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in comparisonUnitAtomList)
                    sb.Append(item.ToString()).Append(Environment.NewLine);
                var sbs = sb.ToString();
                DocxComparerUtil.NotePad(sbs);
            }

            // this is the second loop that processes all text boxes.
            currentAncestorUnids = null;
            bool skipUntilNextPpr = false;
            foreach (var cua in rComparisonUnitAtomList)
            {
                if (currentAncestorUnids != null && cua.AncestorElements.Length < currentAncestorUnids.Length)
                {
                    skipUntilNextPpr = true;
                    currentAncestorUnids = null;
                    continue;
                }
                if (cua.ContentElement.Name == W.pPr)
                {
                    //if (s_True)
                    //{
                    //    var sb = new StringBuilder();
                    //    foreach (var item in comparisonUnitAtomList)
                    //        sb.Append(item.ToString()).Append(Environment.NewLine);
                    //    var sbs = sb.ToString();
                    //    DocxComparerUtil.NotePad(sbs);
                    //}

                    var pPr_inTextBox = cua
                        .AncestorElements
                        .Any(ae => ae.Name == W.txbxContent);

                    if (!pPr_inTextBox)
                    {
                        skipUntilNextPpr = true;
                        currentAncestorUnids = null;
                        continue;
                    }
                    else
                    {
                        skipUntilNextPpr = false;

                        currentAncestorUnids = cua
                            .AncestorElements
                            .Select(ae =>
                            {
                                var thisUnid = (string)ae.Attribute(PtOpenXml.Unid);
                                if (thisUnid == null)
                                    throw new OpenXmlPowerToolsException("Internal error");
                                return thisUnid;
                            })
                            .ToArray();
                        cua.AncestorUnids = currentAncestorUnids;
                        continue;
                    }
                }

                if (skipUntilNextPpr)
                    continue;

                var thisDepth = cua.AncestorElements.Length;
                var additionalAncestorUnids = cua
                    .AncestorElements
                    .Skip(currentAncestorUnids.Length)
                    .Select(ae =>
                    {
                        var thisUnid = (string)ae.Attribute(PtOpenXml.Unid);
                        if (thisUnid == null)
                            Guid.NewGuid().ToString().Replace("-", "");
                        return thisUnid;
                    });
                var thisAncestorUnids = currentAncestorUnids
                    .Concat(additionalAncestorUnids)
                    .ToArray();
                cua.AncestorUnids = thisAncestorUnids;
            }

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in comparisonUnitAtomList)
                    sb.Append(item.ToStringAncestorUnids()).Append(Environment.NewLine);
                var sbs = sb.ToString();
                DocxComparerUtil.NotePad(sbs);
            }
        }

        // the following gets a flattened list of ComparisonUnitAtoms, with status indicated in each ComparisonUnitAtom: Deleted, Inserted, or Equal
        private static List<ComparisonUnitAtom> FlattenToComparisonUnitAtomList(List<CorrelatedSequence> correlatedSequence, WmlComparerSettings settings)
        {
            var listOfComparisonUnitAtoms = correlatedSequence
                .Select(cs =>
                {

                    // need to write some code here to find out if we are assembling a paragraph (or anything) that contains the following unid.
                    // why do are we dropping content???????
                    //string searchFor = "0ecb9184";











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
                                    return new ComparisonUnitAtom(after.ContentElement, after.AncestorElements, after.Part, settings)
                                    {
                                        CorrelationStatus = CorrelationStatus.Equal,
                                        ContentElementBefore = before.ContentElement,
                                        ComparisonUnitAtomBefore = before,
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
                                new ComparisonUnitAtom(ca.ContentElement, ca.AncestorElements, ca.Part, settings)
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
                                new ComparisonUnitAtom(ca.ContentElement, ca.AncestorElements, ca.Part, settings)
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
                DocxComparerUtil.NotePad(sbs);
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

        public static List<WmlComparerRevision> GetRevisions(WmlDocument source, WmlComparerSettings settings)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(source.DocumentByteArray, 0, source.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    TestForInvalidContent(wDoc);
                    RemoveExistingPowerToolsMarkup(wDoc);

                    var contentParent = wDoc.MainDocumentPart.GetXDocument().Root.Element(W.body);
                    var atomList = WmlComparer.CreateComparisonUnitAtomList(wDoc.MainDocumentPart, contentParent, settings).ToArray();

                    if (s_False)
                    {
                        var sb = new StringBuilder();
                        foreach (var item in atomList)
                            sb.Append(item.ToString() + Environment.NewLine);
                        var sbs = sb.ToString();
                        DocxComparerUtil.NotePad(sbs);
                    }

                    var grouped = atomList
                        .GroupAdjacent(a =>
                        {
                            var key = a.CorrelationStatus.ToString();
                            if (a.CorrelationStatus != CorrelationStatus.Equal)
                            {
                                var rt = new XElement(a.RevTrackElement.Name,
                                    new XAttribute(XNamespace.Xmlns + "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
                                    a.RevTrackElement.Attributes().Where(a2 => a2.Name != W.id && a2.Name != PtOpenXml.Unid));
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
                        DocxComparerUtil.NotePad(sbs);
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

                    var footnotesRevisionList = GetFootnoteEndnoteRevisionList(wDoc.MainDocumentPart.FootnotesPart, W.footnote, settings);
                    var endnotesRevisionList = GetFootnoteEndnoteRevisionList(wDoc.MainDocumentPart.EndnotesPart, W.endnote, settings);
                    var finalRevisionList = mainDocPartRevisionList.Concat(footnotesRevisionList).Concat(endnotesRevisionList).ToList();
                    return finalRevisionList;
                }
            }
        }

        private static IEnumerable<WmlComparerRevision> GetFootnoteEndnoteRevisionList(OpenXmlPart footnotesEndnotesPart,
            XName footnoteEndnoteElementName,
            WmlComparerSettings settings)
        {
            if (footnotesEndnotesPart == null)
                return Enumerable.Empty<WmlComparerRevision>();

            var xDoc = footnotesEndnotesPart.GetXDocument();
            var footnotesEndnotes = xDoc.Root.Elements(footnoteEndnoteElementName);
            List<WmlComparerRevision> revisionsForPart = new List<WmlComparerRevision>();
            foreach (var fn in footnotesEndnotes)
            {
                var atomList = WmlComparer.CreateComparisonUnitAtomList(footnotesEndnotesPart, fn, settings).ToArray();

                if (s_False)
                {
                    var sb = new StringBuilder();
                    foreach (var item in atomList)
                        sb.Append(item.ToString() + Environment.NewLine);
                    var sbs = sb.ToString();
                    DocxComparerUtil.NotePad(sbs);
                }

                var grouped = atomList
                    .GroupAdjacent(a =>
                    {
                        var key = a.CorrelationStatus.ToString();
                        if (a.CorrelationStatus != CorrelationStatus.Equal)
                        {
                            var rt = new XElement(a.RevTrackElement.Name,
                                new XAttribute(XNamespace.Xmlns + "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
                                a.RevTrackElement.Attributes().Where(a2 => a2.Name != W.id && a2.Name != PtOpenXml.Unid));
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
                .Where(a => a.Name != PtOpenXml.Unid)
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
                    .Where(a => a.Name != PtOpenXml.Unid)
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
                    .Where(a => a.Name != PtOpenXml.Unid)
                    .Remove();
                enPart.PutXDocument();
            }
        }

        private static void AddSha1HashToBlockLevelContent(OpenXmlPart part, XElement contentParent, WmlComparerSettings settings)
        {
            var blockLevelContentToAnnotate = contentParent
                .Descendants()
                .Where(d => ElementsToHaveSha1Hash.Contains(d.Name));

            foreach (var blockLevelContent in blockLevelContentToAnnotate)
            {
                var cloneBlockLevelContentForHashing = (XElement)CloneBlockLevelContentForHashing(part, blockLevelContent, true, settings);
                var shaString = cloneBlockLevelContentForHashing.ToString(SaveOptions.DisableFormatting)
                    .Replace(" xmlns=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", "");
                var sha1Hash = PtUtils.SHA1HashStringForUTF8String(shaString);
                blockLevelContent.Add(new XAttribute(PtOpenXml.SHA1Hash, sha1Hash));

                if (blockLevelContent.Name == W.tbl ||
                    blockLevelContent.Name == W.tr)
                {
                    var clonedForStructureHash = (XElement)CloneForStructureHash(cloneBlockLevelContentForHashing);

                    // this is a convenient place to look at why tables are being compared as different.

                    //if (blockLevelContent.Name == W.tbl)
                    //    Console.WriteLine();

                    var shaString2 = clonedForStructureHash.ToString(SaveOptions.DisableFormatting)
                        .Replace(" xmlns=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", "");
                    var sha1Hash2 = PtUtils.SHA1HashStringForUTF8String(shaString2);
                    blockLevelContent.Add(new XAttribute(PtOpenXml.StructureSHA1Hash, sha1Hash2));
                }
            }
        }

        // This strips all text nodes from the XML tree, thereby leaving only the structure.
        private static object CloneForStructureHash(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Elements().Select(e => CloneForStructureHash(e)));
            }
            return null;
        }

        static XName[] AttributesToTrimWhenCloning = new XName[] {
            WP14.anchorId,
            WP14.editId,
            "ObjectID",
            "ShapeID",
            "id",
            "type",
        };

        private static XElement CloneBlockLevelContentForHashing(OpenXmlPart mainDocumentPart, XNode node, bool includeRelatedParts, WmlComparerSettings settings)
        {
            var rValue = (XElement)CloneBlockLevelContentForHashingInternal(mainDocumentPart, node, includeRelatedParts, settings);
            rValue.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
            return rValue;
        }

        private static object CloneBlockLevelContentForHashingInternal(OpenXmlPart mainDocumentPart, XNode node, bool includeRelatedParts, WmlComparerSettings settings)
        {
            var element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.bookmarkStart ||
                    element.Name == W.bookmarkEnd ||
                    element.Name == W.pPr ||
                    element.Name == W.rPr)
                    return null;

                if (element.Name.Namespace == A14.a14)
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
                        element.Nodes().Select(n => CloneBlockLevelContentForHashingInternal(mainDocumentPart, n, includeRelatedParts, settings)));

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
                                var text = g.Select(t => t.Value).StringConcatenate();
                                if (settings.CaseInsensitive)
                                    text = text.ToUpper(settings.CultureInfo);
                                if (settings.ConflateBreakingAndNonbreakingSpaces)
                                    text = text.Replace(' ', '\x00a0');
                                var newRun = (object)new XElement(W.r,
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
                    var clonedRuns = element
                        .Elements()
                        .Where(e => e.Name != W.rPr)
                        .Select(rc => new XElement(W.r, CloneBlockLevelContentForHashingInternal(mainDocumentPart, rc, includeRelatedParts, settings)));
                    return clonedRuns;
                }

                if (element.Name == W.tbl)
                {
                    var clonedTable = new XElement(W.tbl,
                        element.Elements(W.tr).Select(n => CloneBlockLevelContentForHashingInternal(mainDocumentPart, n, includeRelatedParts, settings)));
                    return clonedTable;
                }

                if (element.Name == W.tr)
                {
                    var clonedRow = new XElement(W.tr,
                        element.Elements(W.tc).Select(n => CloneBlockLevelContentForHashingInternal(mainDocumentPart, n, includeRelatedParts, settings)));
                    return clonedRow;
                }

                if (element.Name == W.tc)
                {
                    var clonedCell = new XElement(W.tc,
                        element.Elements().Select(n => CloneBlockLevelContentForHashingInternal(mainDocumentPart, n, includeRelatedParts, settings)));
                    return clonedCell;
                }

                if (element.Name == W.tcPr)
                {
                    var clonedCellProps = new XElement(W.tcPr,
                        element.Elements(W.gridSpan).Select(n => CloneBlockLevelContentForHashingInternal(mainDocumentPart, n, includeRelatedParts, settings)));
                    return clonedCellProps;
                }

                if (element.Name == W.gridSpan)
                {
                    var clonedGridSpan = new XElement(W.gridSpan,
                        new XAttribute("val", (string)element.Attribute(W.val)));
                    return clonedGridSpan;
                }

                if (element.Name == W.txbxContent)
                {
                    var clonedTextbox = new XElement(W.txbxContent,
                        element.Elements().Select(n => CloneBlockLevelContentForHashingInternal(mainDocumentPart, n, includeRelatedParts, settings)));
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
                                            using (var str = oxp.GetStream())
                                            {
                                                byte[] ba;
                                                using (BinaryReader br = new BinaryReader(str))
                                                {
                                                    ba = br.ReadBytes((int)str.Length);
                                                }
                                                var sha1 = PtUtils.SHA1HashStringForByteArray(ba);
                                                oxp.AddAnnotation(new PartSHA1HashAnnotation(sha1));
                                                return new XAttribute(a.Name, sha1);
                                            }
                                        }
                                    }
                                    catch (ArgumentOutOfRangeException)
                                    {
                                        HyperlinkRelationship hr = mainDocumentPart.HyperlinkRelationships.FirstOrDefault(z => z.Id == rId);
                                        if (hr != null)
                                        {
                                            var str = hr.Uri.ToString();
                                            return new XAttribute(a.Name, str);
                                        }
                                        // could be an external relationship
                                        ExternalRelationship er = mainDocumentPart.ExternalRelationships.FirstOrDefault(z => z.Id == rId);
                                        if (er != null)
                                        {
                                            var str = er.Uri.ToString();
                                            return new XAttribute(a.Name, str);
                                        }
                                        return new XAttribute(a.Name, "NULL Relationship");
                                    }

                                    return null;
                                }),
                            element.Nodes().Select(n => CloneBlockLevelContentForHashingInternal(mainDocumentPart, n, includeRelatedParts, settings)));
                        return newElement;
                    }
                }

                if (element.Name == VML.shape)
                {
                    return new XElement(element.Name,
                        element.Attributes()
                            .Where(a => a.Name.Namespace != PtOpenXml.pt)
                            .Where(a => a.Name != "style" && a.Name != "id" && a.Name != "type"),
                        element.Nodes().Select(n => CloneBlockLevelContentForHashingInternal(mainDocumentPart, n, includeRelatedParts, settings)));
                }

                if (element.Name == O.OLEObject)
                {
                    var o = new XElement(element.Name,
                        element.Attributes()
                            .Where(a => a.Name.Namespace != PtOpenXml.pt)
                            .Where(a => a.Name != "ObjectID" && a.Name != R.id),
                        element.Nodes().Select(n => CloneBlockLevelContentForHashingInternal(mainDocumentPart, n, includeRelatedParts, settings)));
                    return o;
                }

                if (element.Name == W._object)
                {
                    var o = new XElement(element.Name,
                        element.Attributes()
                            .Where(a => a.Name.Namespace != PtOpenXml.pt),
                        element.Nodes().Select(n => CloneBlockLevelContentForHashingInternal(mainDocumentPart, n, includeRelatedParts, settings)));
                    return o;
                }

                if (element.Name == WP.docPr)
                {
                    return new XElement(element.Name,
                        element.Attributes()
                            .Where(a => a.Name.Namespace != PtOpenXml.pt && a.Name != "id"),
                        element.Nodes().Select(n => CloneBlockLevelContentForHashingInternal(mainDocumentPart, n, includeRelatedParts, settings)));
                }

                if (element.Name == W.footnoteReference || element.Name == W.endnoteReference)
                {
                    return new XElement(element.Name,
                        element.Attributes()
                            .Where(a => a.Name.Namespace != PtOpenXml.pt && a.Name != W.id),
                        element.Nodes().Select(n => CloneBlockLevelContentForHashingInternal(mainDocumentPart, n, includeRelatedParts, settings)));
                }

                return new XElement(element.Name,
                    element.Attributes()
                        .Where(a => a.Name.Namespace != PtOpenXml.pt)
                        .Where(a => !AttributesToTrimWhenCloning.Contains(a.Name)),
                    element.Nodes().Select(n => CloneBlockLevelContentForHashingInternal(mainDocumentPart, n, includeRelatedParts, settings)));
            }
            if (settings.CaseInsensitive || settings.ConflateBreakingAndNonbreakingSpaces)
            {
                var xt = node as XText;
                if (xt != null)
                {
                    var text = xt.Value;
                    if (settings.CaseInsensitive)
                        text = text.ToUpper(settings.CultureInfo);
                    if (settings.ConflateBreakingAndNonbreakingSpaces)
                        text = text.Replace(' ', '\x00a0');
                    return new XText(text);
                }
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

            if (countCommonAtBeginning != 0)
            {
                var newSequence = new List<CorrelatedSequence>();

                CorrelatedSequence csEqual = new CorrelatedSequence();
                csEqual.CorrelationStatus = CorrelationStatus.Equal;
                csEqual.ComparisonUnitArray1 = unknown
                    .ComparisonUnitArray1
                    .Take(countCommonAtBeginning)
                    .ToArray();
                csEqual.ComparisonUnitArray2 = unknown
                    .ComparisonUnitArray2
                    .Take(countCommonAtBeginning)
                    .ToArray();
                newSequence.Add(csEqual);

                var remainingLeft = unknown.ComparisonUnitArray1.Length - countCommonAtBeginning;
                var remainingRight = unknown.ComparisonUnitArray2.Length - countCommonAtBeginning;

                if (remainingLeft != 0 && remainingRight == 0)
                {
                    CorrelatedSequence csDeleted = new CorrelatedSequence();
                    csDeleted.CorrelationStatus = CorrelationStatus.Deleted;
                    csDeleted.ComparisonUnitArray1 = unknown.ComparisonUnitArray1.Skip(countCommonAtBeginning).ToArray();
                    csDeleted.ComparisonUnitArray2 = null;
                    newSequence.Add(csDeleted);
                }
                else if (remainingLeft == 0 && remainingRight != 0)
                {
                    CorrelatedSequence csInserted = new CorrelatedSequence();
                    csInserted.CorrelationStatus = CorrelationStatus.Inserted;
                    csInserted.ComparisonUnitArray1 = null;
                    csInserted.ComparisonUnitArray2 = unknown.ComparisonUnitArray2.Skip(countCommonAtBeginning).ToArray();
                    newSequence.Add(csInserted);
                }
                else if (remainingLeft != 0 && remainingRight != 0)
                {
                    var first1 = unknown.ComparisonUnitArray1[0] as ComparisonUnitWord;
                    var first2 = unknown.ComparisonUnitArray2[0] as ComparisonUnitWord;

                    if (first1 != null && first2 != null)
                    {
                        // if operating at the word level and
                        //   if the last word on the left != pPr && last word on right != pPr
                        //     then create an unknown for the rest of the paragraph, and create an unknown for the rest of the unknown
                        //   if the last word on the left != pPr and last word on right == pPr
                        //     then create deleted for the left, and create an unknown for the rest of the unknown
                        //   if the last word on the left == pPr and last word on right != pPr
                        //     then create inserted for the right, and create an unknown for the rest of the unknown
                        //   if the last word on the left == pPr and last word on right == pPr
                        //     then create an unknown for the rest of the unknown

                        var remainingInLeft = unknown
                            .ComparisonUnitArray1
                            .Skip(countCommonAtBeginning)
                            .ToArray();

                        var remainingInRight = unknown
                            .ComparisonUnitArray2
                            .Skip(countCommonAtBeginning)
                            .ToArray();

                        var lastContentAtomLeft = unknown.ComparisonUnitArray1[countCommonAtBeginning - 1].DescendantContentAtoms().FirstOrDefault();
                        var lastContentAtomRight = unknown.ComparisonUnitArray2[countCommonAtBeginning - 1].DescendantContentAtoms().FirstOrDefault();

                        if (lastContentAtomLeft.ContentElement.Name != W.pPr && lastContentAtomRight.ContentElement.Name != W.pPr)
                        {
                            var split1 = SplitAtParagraphMark(remainingInLeft);
                            var split2 = SplitAtParagraphMark(remainingInRight);
                            if (split1.Count() == 1 && split2.Count() == 1)
                            {
                                CorrelatedSequence csUnknown2 = new CorrelatedSequence();
                                csUnknown2.CorrelationStatus = CorrelationStatus.Unknown;
                                csUnknown2.ComparisonUnitArray1 = split1.First();
                                csUnknown2.ComparisonUnitArray2 = split2.First();
                                newSequence.Add(csUnknown2);
                                return newSequence;
                            }
                            else if (split1.Count == 2 && split2.Count == 2)
                            {
                                CorrelatedSequence csUnknown2 = new CorrelatedSequence();
                                csUnknown2.CorrelationStatus = CorrelationStatus.Unknown;
                                csUnknown2.ComparisonUnitArray1 = split1.First();
                                csUnknown2.ComparisonUnitArray2 = split2.First();
                                newSequence.Add(csUnknown2);

                                CorrelatedSequence csUnknown3 = new CorrelatedSequence();
                                csUnknown3.CorrelationStatus = CorrelationStatus.Unknown;
                                csUnknown3.ComparisonUnitArray1 = split1.Skip(1).First();
                                csUnknown3.ComparisonUnitArray2 = split2.Skip(1).First();
                                newSequence.Add(csUnknown3);

                                return newSequence;
                            }
                        }
                    }

                    CorrelatedSequence csUnknown = new CorrelatedSequence();
                    csUnknown.CorrelationStatus = CorrelationStatus.Unknown;
                    csUnknown.ComparisonUnitArray1 = unknown.ComparisonUnitArray1.Skip(countCommonAtBeginning).ToArray();
                    csUnknown.ComparisonUnitArray2 = unknown.ComparisonUnitArray2.Skip(countCommonAtBeginning).ToArray();
                    newSequence.Add(csUnknown);
                }
                else if (remainingLeft == 0 && remainingRight == 0)
                {
                    // nothing to do
                }
                return newSequence;
            }

            // if we get to here, then countCommonAtBeginning == 0

            var countCommonAtEnd = unknown
                .ComparisonUnitArray1
                .Reverse()
                .Take(lengthToCompare)
                .Zip(unknown
                    .ComparisonUnitArray2
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
            if (countCommonAtEnd == 2)
            {
                var firstCommon = unknown
                    .ComparisonUnitArray1
                    .Reverse()
                    .Take(countCommonAtEnd)
                    .LastOrDefault();

                var secondCommon = unknown
                    .ComparisonUnitArray1
                    .Reverse()
                    .Take(countCommonAtEnd)
                    .FirstOrDefault();

                var firstCommonWord = firstCommon as ComparisonUnitWord;
                var secondCommonWord = secondCommon as ComparisonUnitWord;
                if (firstCommonWord != null && secondCommonWord != null)
                {
                    // if the word contains more than one atom, then not a paragraph mark
                    if (firstCommonWord.Contents.Count() == 1 && secondCommonWord.Contents.Count() == 1)
                    {
                        var firstCommonAtom = firstCommonWord.Contents.First() as ComparisonUnitAtom;
                        var secondCommonAtom = secondCommonWord.Contents.First() as ComparisonUnitAtom;
                        if (firstCommonAtom != null && secondCommonAtom != null)
                        {
                            if (secondCommonAtom.ContentElement.Name == W.pPr)
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

            if (countCommonAtEnd == 0)
                return null;

            // if countCommonAtEnd != 0, and if it contains a paragraph mark, then if there are comparison units in the same paragraph before the common at end (in either version)
            // then we want to put all of those comparison units into a single unknown, where they must be resolved against each other.  We don't want those comparison units to go into the middle unknown comparison unit.

            if (countCommonAtEnd != 0)
            {
                int remainingInLeftParagraph = 0;
                int remainingInRightParagraph = 0;

                var commonEndSeq = unknown
                    .ComparisonUnitArray1
                    .Reverse()
                    .Take(countCommonAtEnd)
                    .Reverse()
                    .ToList();
                var firstOfCommonEndSeq = commonEndSeq.First();
                if (firstOfCommonEndSeq is ComparisonUnitWord)
                {
                    // are there any paragraph marks in the common seq at end?
                    //if (commonEndSeq.Any(cu => cu.Contents.OfType<ComparisonUnitAtom>().First().ContentElement.Name == W.pPr))
                    if (commonEndSeq.Any(cu =>
                    {
                        var firstComparisonUnitAtom = cu.Contents.OfType<ComparisonUnitAtom>().FirstOrDefault();
                        if (firstComparisonUnitAtom == null)
                            return false;
                        return firstComparisonUnitAtom.ContentElement.Name == W.pPr;
                    }))
                    {
                        remainingInLeftParagraph = unknown
                        .ComparisonUnitArray1
                        .Reverse()
                        .Skip(countCommonAtEnd)
                        .TakeWhile(cu =>
                        {
                            if (!(cu is ComparisonUnitWord))
                                return false;
                            var firstComparisonUnitAtom = cu.Contents.OfType<ComparisonUnitAtom>().FirstOrDefault();
                            if (firstComparisonUnitAtom == null)
                                return true;
                            return firstComparisonUnitAtom.ContentElement.Name != W.pPr;
                        })
                        .Count();
                        remainingInRightParagraph = unknown
                            .ComparisonUnitArray2
                            .Reverse()
                            .Skip(countCommonAtEnd)
                            .TakeWhile(cu =>
                            {
                                if (!(cu is ComparisonUnitWord))
                                    return false;
                                var firstComparisonUnitAtom = cu.Contents.OfType<ComparisonUnitAtom>().FirstOrDefault();
                                if (firstComparisonUnitAtom == null)
                                    return true;
                                return firstComparisonUnitAtom.ContentElement.Name != W.pPr;
                            })
                            .Count();
                    }
                }

                var newSequence = new List<CorrelatedSequence>();

                int beforeCommonParagraphLeft = unknown.ComparisonUnitArray1.Length - remainingInLeftParagraph - countCommonAtEnd;
                int beforeCommonParagraphRight = unknown.ComparisonUnitArray2.Length - remainingInRightParagraph - countCommonAtEnd;

                if (beforeCommonParagraphLeft != 0 && beforeCommonParagraphRight == 0)
                {
                    CorrelatedSequence csDeleted = new CorrelatedSequence();
                    csDeleted.CorrelationStatus = CorrelationStatus.Deleted;
                    csDeleted.ComparisonUnitArray1 = unknown.ComparisonUnitArray1.Take(beforeCommonParagraphLeft).ToArray();
                    csDeleted.ComparisonUnitArray2 = null;
                    newSequence.Add(csDeleted);
                }
                else if (beforeCommonParagraphLeft == 0 && beforeCommonParagraphRight != 0)
                {
                    CorrelatedSequence csInserted = new CorrelatedSequence();
                    csInserted.CorrelationStatus = CorrelationStatus.Inserted;
                    csInserted.ComparisonUnitArray1 = null;
                    csInserted.ComparisonUnitArray2 = unknown.ComparisonUnitArray2.Take(beforeCommonParagraphRight).ToArray();
                    newSequence.Add(csInserted);
                }
                else if (beforeCommonParagraphLeft != 0 && beforeCommonParagraphRight != 0)
                {
                    CorrelatedSequence csUnknown = new CorrelatedSequence();
                    csUnknown.CorrelationStatus = CorrelationStatus.Unknown;
                    csUnknown.ComparisonUnitArray1 = unknown.ComparisonUnitArray1.Take(beforeCommonParagraphLeft).ToArray();
                    csUnknown.ComparisonUnitArray2 = unknown.ComparisonUnitArray2.Take(beforeCommonParagraphRight).ToArray();
                    newSequence.Add(csUnknown);
                }
                else if (beforeCommonParagraphLeft == 0 && beforeCommonParagraphRight == 0)
                {
                    // nothing to do
                }

                if (remainingInLeftParagraph != 0 && remainingInRightParagraph == 0)
                {
                    CorrelatedSequence csDeleted = new CorrelatedSequence();
                    csDeleted.CorrelationStatus = CorrelationStatus.Deleted;
                    csDeleted.ComparisonUnitArray1 = unknown.ComparisonUnitArray1.Skip(beforeCommonParagraphLeft).Take(remainingInLeftParagraph).ToArray();
                    csDeleted.ComparisonUnitArray2 = null;
                    newSequence.Add(csDeleted);
                }
                else if (remainingInLeftParagraph == 0 && remainingInRightParagraph != 0)
                {
                    CorrelatedSequence csInserted = new CorrelatedSequence();
                    csInserted.CorrelationStatus = CorrelationStatus.Inserted;
                    csInserted.ComparisonUnitArray1 = null;
                    csInserted.ComparisonUnitArray2 = unknown.ComparisonUnitArray2.Skip(beforeCommonParagraphRight).Take(remainingInRightParagraph).ToArray();
                    newSequence.Add(csInserted);
                }
                else if (remainingInLeftParagraph != 0 && remainingInRightParagraph != 0)
                {
                    CorrelatedSequence csUnknown = new CorrelatedSequence();
                    csUnknown.CorrelationStatus = CorrelationStatus.Unknown;
                    csUnknown.ComparisonUnitArray1 = unknown.ComparisonUnitArray1.Skip(beforeCommonParagraphLeft).Take(remainingInLeftParagraph).ToArray();
                    csUnknown.ComparisonUnitArray2 = unknown.ComparisonUnitArray2.Skip(beforeCommonParagraphRight).Take(remainingInRightParagraph).ToArray();
                    newSequence.Add(csUnknown);
                }
                else if (remainingInLeftParagraph == 0 && remainingInRightParagraph == 0)
                {
                    // nothing to do
                }

                CorrelatedSequence csEqual = new CorrelatedSequence();
                csEqual.CorrelationStatus = CorrelationStatus.Equal;
                csEqual.ComparisonUnitArray1 = unknown.ComparisonUnitArray1.Skip(unknown.ComparisonUnitArray1.Length - countCommonAtEnd).ToArray();
                csEqual.ComparisonUnitArray2 = unknown.ComparisonUnitArray2.Skip(unknown.ComparisonUnitArray2.Length - countCommonAtEnd).ToArray();
                newSequence.Add(csEqual);

                return newSequence;
            }

            return null;
#if false
            var middleLeft = unknown
                .ComparisonUnitArray1
                .Skip(countCommonAtBeginning)
                .SkipLast(remainingInLeftParagraph)
                .SkipLast(countCommonAtEnd)
                .ToArray();

            var middleRight = unknown
                .ComparisonUnitArray2
                .Skip(countCommonAtBeginning)
                .SkipLast(remainingInRightParagraph)
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

            var remainingInParaLeft = unknown
                .ComparisonUnitArray1
                .Skip(countCommonAtBeginning)
                .Skip(middleLeft.Length)
                .Take(remainingInLeftParagraph)
                .ToArray();

            var remainingInParaRight = unknown
                .ComparisonUnitArray2
                .Skip(countCommonAtBeginning)
                .Skip(middleRight.Length)
                .Take(remainingInRightParagraph)
                .ToArray();

            if (remainingInParaLeft.Length > 0 && remainingInParaRight.Length == 0)
            {
                CorrelatedSequence cs = new CorrelatedSequence();
                cs.CorrelationStatus = CorrelationStatus.Deleted;
                cs.ComparisonUnitArray1 = remainingInParaLeft;
                cs.ComparisonUnitArray2 = null;
                newSequence.Add(cs);
            }
            else if (remainingInParaLeft.Length == 0 && remainingInParaRight.Length > 0)
            {
                CorrelatedSequence cs = new CorrelatedSequence();
                cs.CorrelationStatus = CorrelationStatus.Inserted;
                cs.ComparisonUnitArray1 = null;
                cs.ComparisonUnitArray2 = remainingInParaRight;
                newSequence.Add(cs);
            }
            else if (remainingInParaLeft.Length > 0 && remainingInParaRight.Length > 0)
            {
                CorrelatedSequence cs = new CorrelatedSequence();
                cs.CorrelationStatus = CorrelationStatus.Unknown;
                cs.ComparisonUnitArray1 = remainingInParaLeft;
                cs.ComparisonUnitArray2 = remainingInParaRight;
                newSequence.Add(cs);
            }

            if (countCommonAtEnd != 0)
            {
                CorrelatedSequence cs = new CorrelatedSequence();
                cs.CorrelationStatus = CorrelationStatus.Equal;

                cs.ComparisonUnitArray1 = unknown
                    .ComparisonUnitArray1
                    .Skip(countCommonAtBeginning + middleLeft.Length + remainingInParaLeft.Length)
                    .ToArray();

                cs.ComparisonUnitArray2 = unknown
                    .ComparisonUnitArray2
                    .Skip(countCommonAtBeginning + middleRight.Length + remainingInParaRight.Length)
                    .ToArray();

                if (cs.ComparisonUnitArray1.Length != cs.ComparisonUnitArray2.Length)
                    throw new OpenXmlPowerToolsException("Internal error");

                newSequence.Add(cs);
            }
            return newSequence;
#endif
        }

        private static List<ComparisonUnit[]> SplitAtParagraphMark(ComparisonUnit[] cua)
        {
            int i;
            for (i = 0; i < cua.Length; i++)
            {
                var atom = cua[i].DescendantContentAtoms().FirstOrDefault();
                if (atom != null && atom.ContentElement.Name == W.pPr)
                    break;
            }
            if (i == cua.Length)
            {
                return new List<ComparisonUnit[]>()
                {
                    cua
                };
            }
            return new List<ComparisonUnit[]>()
            {
                cua.Take(i).ToArray(),
                cua.Skip(i).ToArray(),
            };
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
                var thisId = nextId++;

                var idAtt = item.Attribute("id");
                if (idAtt != null)
                    idAtt.Value = thisId.ToString();

                var oleObject = item.Parent.Element(O.OLEObject);
                if (oleObject != null)
                {
                    var shapeIdAtt = oleObject.Attribute("ShapeID");
                    if (shapeIdAtt != null)
                        shapeIdAtt.Value = thisId.ToString();
                }
            }
            foreach (var cp in wDoc.ContentParts())
                cp.PutXDocument();
        }

        private static void FixUpGroupIds(WordprocessingDocument wDoc)
        {
            var elementToFind = VML.group;
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
                    idAtt.Value = thisId.ToString();
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
                var thisId = nextId++;

                var idAtt = item.Attribute("id");
                if (idAtt != null)
                    idAtt.Value = thisId.ToString();

                var shape = item.Parent.Element(VML.shape);
                if (shape != null)
                {
                    var typeAtt = shape.Attribute("type");
                    if (typeAtt != null)
                        typeAtt.Value = thisId.ToString();
                }
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
                return ca.AncestorUnids[level];
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
                DocxComparerUtil.NotePad(sbs);
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
                                key = gc.AncestorUnids[level + 1];
                            }
                            if (gc.AncestorElements.Skip(level).Any(ae => ae.Name == W.txbxContent))
                                key += "|" + CorrelationStatus.Equal.ToString();
                            else
                                key += "|" + gc.CorrelationStatus.ToString();
                            return key;
                        })
                        .ToList();

                    if (ancestorBeingConstructed.Name == W.p)
                    {
                        var newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                var spl = gc.Key.Split('|');
                                if (spl[0] == "")
                                    return (object)gc.Select(gcc =>
                                    {
                                        var dup = new XElement(gcc.ContentElement);
                                        if (spl[1] == "Deleted")
                                            dup.Add(new XAttribute(PtOpenXml.Status, "Deleted"));
                                        else if (spl[1] == "Inserted")
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
                            ancestorBeingConstructed.Attributes().Where(a => a.Name.Namespace != PtOpenXml.pt),
                            new XAttribute(PtOpenXml.Unid, g.Key),
                            newChildElements);

                        return newPara;
                    }

                    if (ancestorBeingConstructed.Name == W.r)
                    {
                        var newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                var spl = gc.Key.Split('|');
                                if (spl[0] == "")
                                    return (object)gc.Select(gcc =>
                                    {
                                        var dup = new XElement(gcc.ContentElement);
                                        if (spl[1] == "Deleted")
                                            dup.Add(new XAttribute(PtOpenXml.Status, "Deleted"));
                                        else if (spl[1] == "Inserted")
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
                            ancestorBeingConstructed.Attributes().Where(a => a.Name.Namespace != PtOpenXml.pt),
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
                                            ancestorBeingConstructed.Attributes().Where(a => a.Name.Namespace != PtOpenXml.pt),
                                            new XAttribute(PtOpenXml.Status, "Deleted"));
                                        return dup;
                                    });
                                }
                                else if (ins)
                                {
                                    return gc.Select(gcc =>
                                    {
                                        var dup = new XElement(ancestorBeingConstructed.Name,
                                            ancestorBeingConstructed.Attributes().Where(a => a.Name.Namespace != PtOpenXml.pt),
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
                        return ReconstructElement(part, g, ancestorBeingConstructed, W.tblPr, W.tblGrid, null, level, settings);
                    if (ancestorBeingConstructed.Name == W.tr)
                        return ReconstructElement(part, g, ancestorBeingConstructed, W.trPr, null, null, level, settings);
                    if (ancestorBeingConstructed.Name == W.tc)
                        return ReconstructElement(part, g, ancestorBeingConstructed, W.tcPr, null, null, level, settings);
                    if (ancestorBeingConstructed.Name == W.sdt)
                        return ReconstructElement(part, g, ancestorBeingConstructed, W.sdtPr, W.sdtEndPr, null, level, settings);
                    if (ancestorBeingConstructed.Name == W.pict)
                        return ReconstructElement(part, g, ancestorBeingConstructed, VML.shapetype, null, null, level, settings);
                    if (ancestorBeingConstructed.Name == VML.shape)
                        return ReconstructElement(part, g, ancestorBeingConstructed, W10.wrap, null, null, level, settings);
                    if (ancestorBeingConstructed.Name == W._object)
                        return ReconstructElement(part, g, ancestorBeingConstructed, VML.shapetype, VML.shape, O.OLEObject, level, settings);
                    if (ancestorBeingConstructed.Name == W.ruby)
                        return ReconstructElement(part, g, ancestorBeingConstructed, W.rubyPr, null, null, level, settings);
                    return (object)ReconstructElement(part, g, ancestorBeingConstructed, null, null, null, level, settings);
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
                .Where(d => d.Name != C.externalData)
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

                    var tartString = relationshipForDeletedPart.TargetUri.ToString();

                    Uri targetUri;
                    try
                    {
                        targetUri = PackUriHelper
                            .ResolvePartUri(
                                new Uri(partOfDeletedContent.Uri.ToString(), UriKind.RelativeOrAbsolute),
                                    new Uri(tartString, UriKind.RelativeOrAbsolute));
                    }
                    catch (System.ArgumentException)
                    {
                        targetUri = null;
                    }

                    if (targetUri != null)
                    {

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
            XName props2XName, XName props3XName, int level, WmlComparerSettings settings)
        {
            var newChildElements = CoalesceRecurse(part, g, level + 1, settings);

            object props1 = null;
            if (props1XName != null)
                props1 = ancestorBeingConstructed.Elements(props1XName);

            object props2 = null;
            if (props2XName != null)
                props2 = ancestorBeingConstructed.Elements(props2XName);

            object props3 = null;
            if (props3XName != null)
                props3 = ancestorBeingConstructed.Elements(props3XName);

            var reconstructedElement = new XElement(ancestorBeingConstructed.Name,
                ancestorBeingConstructed.Attributes(),
                props1, props2, props3, newChildElements);

            return reconstructedElement;
        }

        private static List<CorrelatedSequence> DetectUnrelatedSources(ComparisonUnit[] cu1, ComparisonUnit[] cu2, WmlComparerSettings settings)
        {
            if (cu1.OfType<ComparisonUnitGroup>().Take(4).Count() > 3 &&
                cu2.OfType<ComparisonUnitGroup>().Take(4).Count() > 3)
            {
                var list1 = cu1.OfType<ComparisonUnitGroup>().Select(g => g.SHA1Hash).ToList();
                var list2 = cu2.OfType<ComparisonUnitGroup>().Select(g => g.SHA1Hash).ToList();
                var intersect = list1.Intersect(list2).ToList();

                if (intersect.Count() == 0)
                {
                    var newListOfCorrelatedSequence = new List<CorrelatedSequence>();

                    var cul1 = cu1;
                    var cul2 = cu2;

                    var deletedCorrelatedSequence = new CorrelatedSequence();
                    deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                    deletedCorrelatedSequence.ComparisonUnitArray1 = cul1;
                    deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                    newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);

                    var insertedCorrelatedSequence = new CorrelatedSequence();
                    insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                    insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                    insertedCorrelatedSequence.ComparisonUnitArray2 = cul2;
                    newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);

                    return newListOfCorrelatedSequence;
                }
            }
            return null;
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
                    DocxComparerUtil.NotePad(sbs);
                }

                var unknown = csList
                    .FirstOrDefault(z => z.CorrelationStatus == CorrelationStatus.Unknown);

                if (unknown != null)
                {
                    // if unknown consists of a single group of the same type in each side, then can set some Unids in the 'after' document.
                    // if the unknown is a pair of single tables, then can set table Unid.
                    // if the unknown is a pair of single rows, then can set table and rows Unids.
                    // if the unknown is a pair of single cells, then can set table, row, and cell Unids.
                    // if the unknown is a pair of paragraphs, then can set paragraph (and all ancestor) Unids.
                    SetAfterUnids(unknown);

                    if (s_False)
                    {
                        var sb = new StringBuilder();
                        sb.Append(unknown.ToString());
                        var sbs = sb.ToString();
                        DocxComparerUtil.NotePad(sbs);
                    }

                    List<CorrelatedSequence> newSequence = ProcessCorrelatedHashes(unknown, settings);
                    if (newSequence == null)
                    {
                        newSequence = FindCommonAtBeginningAndEnd(unknown, settings);
                        if (newSequence == null)
                        {
                            newSequence = DoLcsAlgorithm(unknown, settings);
                        }
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

        private static void SetAfterUnids(CorrelatedSequence unknown)
        {
            if (unknown.ComparisonUnitArray1.Length == 1 && unknown.ComparisonUnitArray2.Length == 1)
            {
                var cua1 = unknown.ComparisonUnitArray1[0] as ComparisonUnitGroup;
                var cua2 = unknown.ComparisonUnitArray2[0] as ComparisonUnitGroup;
                if (cua1 != null &&
                    cua2 != null &&
                    cua1.ComparisonUnitGroupType == cua2.ComparisonUnitGroupType)
                {
                    var groupType = cua1.ComparisonUnitGroupType;
                    var da1 = cua1.DescendantContentAtoms();
                    var da2 = cua2.DescendantContentAtoms();
                    XName takeThruName = null;
                    switch (groupType)
                    {
                        case ComparisonUnitGroupType.Paragraph:
                            takeThruName = W.p;
                            break;
                        case ComparisonUnitGroupType.Table:
                            takeThruName = W.tbl;
                            break;
                        case ComparisonUnitGroupType.Row:
                            takeThruName = W.tr;
                            break;
                        case ComparisonUnitGroupType.Cell:
                            takeThruName = W.tc;
                            break;
                        case ComparisonUnitGroupType.Textbox:
                            takeThruName = W.txbxContent;
                            break;
                    }
                    if (takeThruName == null)
                        throw new OpenXmlPowerToolsException("Internal error");
                    var relevantAncestors = new List<XElement>();
                    foreach (var ae in da1.First().AncestorElements)
                    {
                        if (ae.Name != takeThruName)
                        {
                            relevantAncestors.Add(ae);
                            continue;
                        }
                        relevantAncestors.Add(ae);
                        break;
                    }
                    var unidList = relevantAncestors
                        .Select(a =>
                        {
                            var unid = (string)a.Attribute(PtOpenXml.Unid);
                            if (unid == null)
                                throw new OpenXmlPowerToolsException("Internal error");
                            return unid;
                        })
                        .ToArray();
                    foreach (var da in da2)
                    {
                        var ancestorsToSet = da.AncestorElements.Take(unidList.Length);
                        var zipped = ancestorsToSet.Zip(unidList, (a, u) =>
                            new
                            {
                                Ancestor = a,
                                Unid = u,
                            });

                        foreach (var z in zipped)
                        {
                            var unid = z.Ancestor.Attribute(PtOpenXml.Unid);

                            if (z.Ancestor.Name == W.footnotes || z.Ancestor.Name == W.endnotes)
                                continue;

                            if (unid == null)
                                throw new OpenXmlPowerToolsException("Internal error");
                            unid.Value = z.Unid;
                        }
                    }
                }
            }
        }

        private static List<CorrelatedSequence> ProcessCorrelatedHashes(CorrelatedSequence unknown, WmlComparerSettings settings)
        {
            // never attempt this optimization if there are less than 3 groups
            var maxd = Math.Min(unknown.ComparisonUnitArray1.Length, unknown.ComparisonUnitArray2.Length);
            if (maxd < 3)
                return null;

            var firstInCu1 = unknown.ComparisonUnitArray1.FirstOrDefault() as ComparisonUnitGroup;
            var firstInCu2 = unknown.ComparisonUnitArray2.FirstOrDefault() as ComparisonUnitGroup;
            if (firstInCu1 != null && firstInCu2 != null)
            {
                if ((firstInCu1.ComparisonUnitGroupType == ComparisonUnitGroupType.Paragraph ||
                    firstInCu1.ComparisonUnitGroupType == ComparisonUnitGroupType.Table ||
                    firstInCu1.ComparisonUnitGroupType == ComparisonUnitGroupType.Row) &&
                    (firstInCu2.ComparisonUnitGroupType == ComparisonUnitGroupType.Paragraph ||
                    firstInCu2.ComparisonUnitGroupType == ComparisonUnitGroupType.Table ||
                    firstInCu2.ComparisonUnitGroupType == ComparisonUnitGroupType.Row))
                {
                    var groupType = firstInCu1.ComparisonUnitGroupType;

                    // Next want to do the lcs algorithm on this.
                    // potentially, we will find all paragraphs are correlated, but they may not be for two reasons-
                    // - if there were changes that were not tracked
                    // - if the anomolies in the change tracking cause there to be a mismatch in the number of paragraphs
                    // therefore we are going to do the whole LCS algorithm thing
                    // and at the end of the process, we set up the correlated sequence list where correlated paragraphs are together in their
                    // own unknown correlated sequence.

                    var cul1 = unknown.ComparisonUnitArray1;
                    var cul2 = unknown.ComparisonUnitArray2;
                    int currentLongestCommonSequenceLength = 0;
                    int currentLongestCommonSequenceAtomCount = 0;
                    int currentI1 = -1;
                    int currentI2 = -1;
                    for (int i1 = 0; i1 < cul1.Length; i1++)
                    {
                        for (int i2 = 0; i2 < cul2.Length; i2++)
                        {
                            var thisSequenceLength = 0;
                            var thisSequenceAtomCount = 0;
                            var thisI1 = i1;
                            var thisI2 = i2;
                            while (true)
                            {
                                var group1 = cul1[thisI1] as ComparisonUnitGroup;
                                var group2 = cul2[thisI2] as ComparisonUnitGroup;
                                bool match = group1 != null &&
                                    group2 != null &&
                                    group1.ComparisonUnitGroupType == group2.ComparisonUnitGroupType &&
                                    group1.CorrelatedSHA1Hash != null &&
                                    group2.CorrelatedSHA1Hash != null &&
                                    group1.CorrelatedSHA1Hash == group2.CorrelatedSHA1Hash;

                                if (match)
                                {
                                    thisSequenceAtomCount += cul1[thisI1].DescendantContentAtomsCount;
                                    thisI1++;
                                    thisI2++;
                                    thisSequenceLength++;
                                    if (thisI1 == cul1.Length || thisI2 == cul2.Length)
                                    {
                                        if (thisSequenceAtomCount > currentLongestCommonSequenceAtomCount)
                                        {
                                            currentLongestCommonSequenceLength = thisSequenceLength;
                                            currentLongestCommonSequenceAtomCount = thisSequenceAtomCount;
                                            currentI1 = i1;
                                            currentI2 = i2;
                                        }
                                        break;
                                    }
                                    continue;
                                }
                                else
                                {
                                    if (thisSequenceAtomCount > currentLongestCommonSequenceAtomCount)
                                    {
                                        currentLongestCommonSequenceLength = thisSequenceLength;
                                        currentLongestCommonSequenceAtomCount = thisSequenceAtomCount;
                                        currentI1 = i1;
                                        currentI2 = i2;
                                    }
                                    break;
                                }
                            }
                        }
                    }

                    // here we want to have some sort of threshold, and if the currentLongestCommonSequenceLength is not longer than the threshold, then don't do anything
                    bool doCorrelation = false;
                    if (currentLongestCommonSequenceLength == 1)
                    {
                        var numberOfAtoms1 = unknown.ComparisonUnitArray1[currentI1].DescendantContentAtoms().Count();
                        var numberOfAtoms2 = unknown.ComparisonUnitArray2[currentI2].DescendantContentAtoms().Count();
                        if (numberOfAtoms1 > 16 && numberOfAtoms2 > 16)
                            doCorrelation = true;
                    }
                    else if (currentLongestCommonSequenceLength > 1 && currentLongestCommonSequenceLength <= 3)
                    {
                        var numberOfAtoms1 = unknown.ComparisonUnitArray1.Skip(currentI1).Take(currentLongestCommonSequenceLength).Select(z => z.DescendantContentAtoms().Count()).Sum();
                        var numberOfAtoms2 = unknown.ComparisonUnitArray2.Skip(currentI2).Take(currentLongestCommonSequenceLength).Select(z => z.DescendantContentAtoms().Count()).Sum();
                        if (numberOfAtoms1 > 32 && numberOfAtoms2 > 32)
                            doCorrelation = true;
                    }
                    else if (currentLongestCommonSequenceLength > 3)
                        doCorrelation = true;
                    if (doCorrelation)
                    {
                        var newListOfCorrelatedSequence = new List<CorrelatedSequence>();

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

                        for (int i = 0; i < currentLongestCommonSequenceLength; i++)
                        {
                            var unknownCorrelatedSequence = new CorrelatedSequence();
                            unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                            unknownCorrelatedSequence.ComparisonUnitArray1 = cul1
                                .Skip(currentI1)
                                .Skip(i)
                                .Take(1)
                                .ToArray();
                            unknownCorrelatedSequence.ComparisonUnitArray2 = cul2
                                .Skip(currentI2)
                                .Skip(i)
                                .Take(1)
                                .ToArray();
                            newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
                        }

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
                    return null;
                }
            }
            return null;
        }

        private static List<CorrelatedSequence> DoLcsAlgorithm(CorrelatedSequence unknown, WmlComparerSettings settings)
        {
            var newListOfCorrelatedSequence = new List<CorrelatedSequence>();

            var cul1 = unknown.ComparisonUnitArray1;
            var cul2 = unknown.ComparisonUnitArray2;

            // first thing to do - if we have an unknown with zero length on left or right side, create appropriate 
            // this is a code optimization that enables easier processing of cases elsewhere.
            if (cul1.Length > 0 && cul2.Length == 0)
            {
                var deletedCorrelatedSequence = new CorrelatedSequence();
                deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                deletedCorrelatedSequence.ComparisonUnitArray1 = cul1;
                deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                return newListOfCorrelatedSequence;
            }
            else if (cul1.Length == 0 && cul2.Length > 0)
            {
                var insertedCorrelatedSequence = new CorrelatedSequence();
                insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                insertedCorrelatedSequence.ComparisonUnitArray2 = cul2;
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                return newListOfCorrelatedSequence;
            }
            else if (cul1.Length == 0 && cul2.Length == 0)
            {
                return newListOfCorrelatedSequence; // this will effectively remove the unknown with no data on either side from the current data model.
            }

            int currentLongestCommonSequenceLength = 0;
            int currentI1 = -1;
            int currentI2 = -1;
            for (int i1 = 0; i1 < cul1.Length - currentLongestCommonSequenceLength; i1++)
            {
                for (int i2 = 0; i2 < cul2.Length - currentLongestCommonSequenceLength; i2++)
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

            // don't match just a single character
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

            // don't match only word break characters
            if (currentLongestCommonSequenceLength > 0 && currentLongestCommonSequenceLength <= 3)
            {
                var commonSequence = cul1.Skip(currentI1).Take(currentLongestCommonSequenceLength).ToArray();
                // if they are all ComparisonUnitWord objects
                var oneIsNotWord = commonSequence.Any(cs => (cs as ComparisonUnitWord) == null);
                var allAreWords = !oneIsNotWord;
                if (allAreWords)
                {
                    var contentOtherThanWordSplitChars = commonSequence
                        .Cast<ComparisonUnitWord>()
                        .Any(cs =>
                        {
                            var otherThanText = cs.DescendantContentAtoms().Any(dca => dca.ContentElement.Name != W.t);
                            if (otherThanText)
                                return true;
                            var otherThanWordSplit = cs
                                .DescendantContentAtoms()
                                .Any(dca =>
                                {
                                    var charValue = dca.ContentElement.Value;
                                    var isWordSplit = ((int)charValue[0] >= 0x4e00 && (int)charValue[0] <= 0x9fff);
                                    if (! isWordSplit)
                                        isWordSplit = settings.WordSeparators.Contains(charValue[0]);
                                    if (isWordSplit)
                                        return false;
                                    return true;
                                });
                            return otherThanWordSplit;
                        });
                    if (!contentOtherThanWordSplitChars)
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

                        // have to decide which of the following two branches to do first based on whether the left contains a paragraph mark
                        // i.e. cant insert a string of deleted text right before a table.

                        else if (leftGrouped[iLeft].Key == "Word" &&
                            leftGrouped[iLeft].Select(lg => lg.DescendantContentAtoms()).SelectMany(m => m).Last().ContentElement.Name != W.pPr &&
                            rightGrouped[iRight].Key == "Row")
                        {
                            var insertedCorrelatedSequence = new CorrelatedSequence();
                            insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                            insertedCorrelatedSequence.ComparisonUnitArray2 = rightGrouped[iRight].ToArray();
                            insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                            ++iRight;
                        }
                        else if (rightGrouped[iRight].Key == "Word" &&
                            rightGrouped[iRight].Select(lg => lg.DescendantContentAtoms()).SelectMany(m => m).Last().ContentElement.Name != W.pPr &&
                            leftGrouped[iLeft].Key == "Row")
                        {
                            var insertedCorrelatedSequence = new CorrelatedSequence();
                            insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                            insertedCorrelatedSequence.ComparisonUnitArray2 = leftGrouped[iLeft].ToArray();
                            insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                            ++iLeft;
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
                    var result = DoLcsAlgorithmForTable(unknown, settings);
                    if (result != null)
                        return result;
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
                    .FirstOrDefault() as ComparisonUnitGroup;

                var firstRight = unknown
                    .ComparisonUnitArray2
                    .FirstOrDefault() as ComparisonUnitGroup;

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
                                var unknownCorrelatedSequence = new CorrelatedSequence();
                                unknownCorrelatedSequence.ComparisonUnitArray1 = new[] { l };
                                unknownCorrelatedSequence.ComparisonUnitArray2 = new[] { r };
                                unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                                return new[] { unknownCorrelatedSequence };
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
                            DocxComparerUtil.NotePad(sbs);
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

                if (unknown.ComparisonUnitArray1.Any() && unknown.ComparisonUnitArray2.Any())
                {
                    var left = unknown.ComparisonUnitArray1.First() as ComparisonUnitWord;
                    var right = unknown.ComparisonUnitArray2.First() as ComparisonUnitGroup;
                    if (left != null &&
                        right != null &&
                        right.ComparisonUnitGroupType == ComparisonUnitGroupType.Row)
                    {
                        var insertedCorrelatedSequence3 = new CorrelatedSequence();
                        insertedCorrelatedSequence3.CorrelationStatus = CorrelationStatus.Inserted;
                        insertedCorrelatedSequence3.ComparisonUnitArray1 = null;
                        insertedCorrelatedSequence3.ComparisonUnitArray2 = unknown.ComparisonUnitArray2;
                        newListOfCorrelatedSequence.Add(insertedCorrelatedSequence3);

                        var deletedCorrelatedSequence3 = new CorrelatedSequence();
                        deletedCorrelatedSequence3.CorrelationStatus = CorrelationStatus.Deleted;
                        deletedCorrelatedSequence3.ComparisonUnitArray1 = unknown.ComparisonUnitArray1;
                        deletedCorrelatedSequence3.ComparisonUnitArray2 = null;
                        newListOfCorrelatedSequence.Add(deletedCorrelatedSequence3);

                        return newListOfCorrelatedSequence;
                    }

                    var left2 = unknown.ComparisonUnitArray1.First() as ComparisonUnitGroup;
                    var right2 = unknown.ComparisonUnitArray2.First() as ComparisonUnitWord;
                    if (right2 != null &&
                        left2 != null &&
                        left2.ComparisonUnitGroupType == ComparisonUnitGroupType.Row)
                    {
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

                    var lastContentAtomLeft = unknown.ComparisonUnitArray1.Select(cu => cu.DescendantContentAtoms().Last()).LastOrDefault();
                    var lastContentAtomRight = unknown.ComparisonUnitArray2.Select(cu => cu.DescendantContentAtoms().Last()).LastOrDefault();
                    if (lastContentAtomLeft != null && lastContentAtomRight != null)
                    {
                        if (lastContentAtomLeft.ContentElement.Name == W.pPr &&
                            lastContentAtomRight.ContentElement.Name != W.pPr)
                        {
                            var insertedCorrelatedSequence5 = new CorrelatedSequence();
                            insertedCorrelatedSequence5.CorrelationStatus = CorrelationStatus.Inserted;
                            insertedCorrelatedSequence5.ComparisonUnitArray1 = null;
                            insertedCorrelatedSequence5.ComparisonUnitArray2 = unknown.ComparisonUnitArray2;
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence5);

                            var deletedCorrelatedSequence5 = new CorrelatedSequence();
                            deletedCorrelatedSequence5.CorrelationStatus = CorrelationStatus.Deleted;
                            deletedCorrelatedSequence5.ComparisonUnitArray1 = unknown.ComparisonUnitArray1;
                            deletedCorrelatedSequence5.ComparisonUnitArray2 = null;
                            newListOfCorrelatedSequence.Add(deletedCorrelatedSequence5);

                            return newListOfCorrelatedSequence;
                        }
                        else if (lastContentAtomLeft.ContentElement.Name != W.pPr &&
                            lastContentAtomRight.ContentElement.Name == W.pPr)
                        {
                            var deletedCorrelatedSequence5 = new CorrelatedSequence();
                            deletedCorrelatedSequence5.CorrelationStatus = CorrelationStatus.Deleted;
                            deletedCorrelatedSequence5.ComparisonUnitArray1 = unknown.ComparisonUnitArray1;
                            deletedCorrelatedSequence5.ComparisonUnitArray2 = null;
                            newListOfCorrelatedSequence.Add(deletedCorrelatedSequence5);

                            var insertedCorrelatedSequence5 = new CorrelatedSequence();
                            insertedCorrelatedSequence5.CorrelationStatus = CorrelationStatus.Inserted;
                            insertedCorrelatedSequence5.ComparisonUnitArray1 = null;
                            insertedCorrelatedSequence5.ComparisonUnitArray2 = unknown.ComparisonUnitArray2;
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence5);

                            return newListOfCorrelatedSequence;
                        }
                    }
                }

                var deletedCorrelatedSequence4 = new CorrelatedSequence();
                deletedCorrelatedSequence4.CorrelationStatus = CorrelationStatus.Deleted;
                deletedCorrelatedSequence4.ComparisonUnitArray1 = unknown.ComparisonUnitArray1;
                deletedCorrelatedSequence4.ComparisonUnitArray2 = null;
                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence4);

                var insertedCorrelatedSequence4 = new CorrelatedSequence();
                insertedCorrelatedSequence4.CorrelationStatus = CorrelationStatus.Inserted;
                insertedCorrelatedSequence4.ComparisonUnitArray1 = null;
                insertedCorrelatedSequence4.ComparisonUnitArray2 = unknown.ComparisonUnitArray2;
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence4);

                return newListOfCorrelatedSequence;
            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // here we have the longest common subsequence.
            // but it may start in the middle of a paragraph.
            // therefore need to dispose of the content from the beginning of the longest common subsequence to the beginning of the paragraph.
            // this should be in a separate unknown region
            // if countCommonAtEnd != 0, and if it contains a paragraph mark, then if there are comparison units in the same paragraph before the common at end (in either version)
            // then we want to put all of those comparison units into a single unknown, where they must be resolved against each other.  We don't want those
            // comparison units to go into the middle unknown comparison unit.

            int remainingInLeftParagraph = 0;
            int remainingInRightParagraph = 0;
            if (currentLongestCommonSequenceLength != 0)
            {
                var commonSeq = unknown
                    .ComparisonUnitArray1
                    .Skip(currentI1)
                    .Take(currentLongestCommonSequenceLength)
                    .ToList();
                var firstOfCommonSeq = commonSeq.First();
                if (firstOfCommonSeq is ComparisonUnitWord)
                {
                    // are there any paragraph marks in the common seq at end?
                    if (commonSeq.Any(cu =>
                    {
                        var firstComparisonUnitAtom = cu.Contents.OfType<ComparisonUnitAtom>().FirstOrDefault();
                        if (firstComparisonUnitAtom == null)
                            return false;
                        return firstComparisonUnitAtom.ContentElement.Name == W.pPr;
                    }))
                    {
                        remainingInLeftParagraph = unknown
                            .ComparisonUnitArray1
                            .Take(currentI1)
                            .Reverse()
                            .TakeWhile(cu =>
                            {
                                if (!(cu is ComparisonUnitWord))
                                    return false;
                                var firstComparisonUnitAtom = cu.Contents.OfType<ComparisonUnitAtom>().FirstOrDefault();
                                if (firstComparisonUnitAtom == null)
                                    return true;
                                return firstComparisonUnitAtom.ContentElement.Name != W.pPr;
                            })
                            .Count();
                        remainingInRightParagraph = unknown
                            .ComparisonUnitArray2
                            .Take(currentI2)
                            .Reverse()
                            .TakeWhile(cu =>
                            {
                                if (!(cu is ComparisonUnitWord))
                                    return false;
                                var firstComparisonUnitAtom = cu.Contents.OfType<ComparisonUnitAtom>().FirstOrDefault();
                                if (firstComparisonUnitAtom == null)
                                    return true;
                                return firstComparisonUnitAtom.ContentElement.Name != W.pPr;
                            })
                            .Count();
                    }
                }
            }

            var countBeforeCurrentParagraphLeft = currentI1 - remainingInLeftParagraph;
            var countBeforeCurrentParagraphRight = currentI2 - remainingInRightParagraph;

            if (countBeforeCurrentParagraphLeft > 0 && countBeforeCurrentParagraphRight == 0)
            {
                var deletedCorrelatedSequence = new CorrelatedSequence();
                deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                deletedCorrelatedSequence.ComparisonUnitArray1 = cul1
                    .Take(countBeforeCurrentParagraphLeft)
                    .ToArray();
                deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
            }
            else if (countBeforeCurrentParagraphLeft == 0 && countBeforeCurrentParagraphRight > 0)
            {
                var insertedCorrelatedSequence = new CorrelatedSequence();
                insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                insertedCorrelatedSequence.ComparisonUnitArray2 = cul2
                    .Take(countBeforeCurrentParagraphRight)
                    .ToArray();
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
            }
            else if (countBeforeCurrentParagraphLeft > 0 && countBeforeCurrentParagraphRight > 0)
            {
                var unknownCorrelatedSequence = new CorrelatedSequence();
                unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                unknownCorrelatedSequence.ComparisonUnitArray1 = cul1
                    .Take(countBeforeCurrentParagraphLeft)
                    .ToArray();
                unknownCorrelatedSequence.ComparisonUnitArray2 = cul2
                    .Take(countBeforeCurrentParagraphRight)
                    .ToArray();

                newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
            }
            else if (countBeforeCurrentParagraphLeft == 0 && countBeforeCurrentParagraphRight == 0)
            {
                // nothing to do
            }

            if (remainingInLeftParagraph > 0 && remainingInRightParagraph == 0)
            {
                var deletedCorrelatedSequence = new CorrelatedSequence();
                deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                deletedCorrelatedSequence.ComparisonUnitArray1 = cul1
                    .Skip(countBeforeCurrentParagraphLeft)
                    .Take(remainingInLeftParagraph)
                    .ToArray();
                deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
            }
            else if (remainingInLeftParagraph == 0 && remainingInRightParagraph > 0)
            {
                var insertedCorrelatedSequence = new CorrelatedSequence();
                insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                insertedCorrelatedSequence.ComparisonUnitArray2 = cul2
                    .Skip(countBeforeCurrentParagraphRight)
                    .Take(remainingInRightParagraph)
                    .ToArray();
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
            }
            else if (remainingInLeftParagraph > 0 && remainingInRightParagraph > 0)
            {
                var unknownCorrelatedSequence = new CorrelatedSequence();
                unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                unknownCorrelatedSequence.ComparisonUnitArray1 = cul1
                    .Skip(countBeforeCurrentParagraphLeft)
                    .Take(remainingInLeftParagraph)
                    .ToArray();
                unknownCorrelatedSequence.ComparisonUnitArray2 = cul2
                    .Skip(countBeforeCurrentParagraphRight)
                    .Take(remainingInRightParagraph)
                    .ToArray();
                newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
            }
            else if (remainingInLeftParagraph == 0 && remainingInRightParagraph == 0)
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

            var remaining1 = cul1
                .Skip(endI1)
                .ToArray();

            var remaining2 = cul2
                .Skip(endI2)
                .ToArray();

            // here is the point that we want to make a new unknown from this point to the end of the paragraph that contains the equal parts.
            // this will never hurt anything, and will in many cases result in a better difference.

            var leftCuw = middleEqual.ComparisonUnitArray1[middleEqual.ComparisonUnitArray1.Length - 1] as ComparisonUnitWord;
            if (leftCuw != null)
            {
                var lastContentAtom = leftCuw.DescendantContentAtoms().LastOrDefault();
                // if the middleEqual did not end with a paragraph mark
                if (lastContentAtom != null && lastContentAtom.ContentElement.Name != W.pPr)
                {
                    int idx1 = FindIndexOfNextParaMark(remaining1);
                    int idx2 = FindIndexOfNextParaMark(remaining2);

                    var unknownCorrelatedSequenceRemaining = new CorrelatedSequence();
                    unknownCorrelatedSequenceRemaining.CorrelationStatus = CorrelationStatus.Unknown;
                    unknownCorrelatedSequenceRemaining.ComparisonUnitArray1 = remaining1.Take(idx1).ToArray();
                    unknownCorrelatedSequenceRemaining.ComparisonUnitArray2 = remaining2.Take(idx2).ToArray();
                    newListOfCorrelatedSequence.Add(unknownCorrelatedSequenceRemaining);

                    var unknownCorrelatedSequenceAfter = new CorrelatedSequence();
                    unknownCorrelatedSequenceAfter.CorrelationStatus = CorrelationStatus.Unknown;
                    unknownCorrelatedSequenceAfter.ComparisonUnitArray1 = remaining1.Skip(idx1).ToArray();
                    unknownCorrelatedSequenceAfter.ComparisonUnitArray2 = remaining2.Skip(idx2).ToArray();
                    newListOfCorrelatedSequence.Add(unknownCorrelatedSequenceAfter);

                    return newListOfCorrelatedSequence;
                }
            }

            var unknownCorrelatedSequence20 = new CorrelatedSequence();
            unknownCorrelatedSequence20.CorrelationStatus = CorrelationStatus.Unknown;
            unknownCorrelatedSequence20.ComparisonUnitArray1 = remaining1;
            unknownCorrelatedSequence20.ComparisonUnitArray2 = remaining2;
            newListOfCorrelatedSequence.Add(unknownCorrelatedSequence20);

            return newListOfCorrelatedSequence;
        }

        private static int FindIndexOfNextParaMark(ComparisonUnit[] cul)
        {
            for (int i = 0; i < cul.Length; i++)
            {
                var cuw = cul[i] as ComparisonUnitWord;
                var lastAtom = cuw.DescendantContentAtoms().LastOrDefault();
                if (lastAtom.ContentElement.Name == W.pPr)
                    return i;
            }
            return cul.Length;
        }

        private static List<CorrelatedSequence> DoLcsAlgorithmForTable(CorrelatedSequence unknown, WmlComparerSettings settings)
        {
            List<CorrelatedSequence> newListOfCorrelatedSequence = new List<CorrelatedSequence>();

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // if we have a table with the same number of rows, and all rows have equal CorrelatedSHA1Hash, then we can flatten and compare every corresponding row.
            // This is true regardless of whether there are horizontally or vertically merged cells, since that characteristic is incorporated into the CorrespondingSHA1Hash.
            // This is probably not very common, but it will never do any harm.
            var tblGroup1 = unknown.ComparisonUnitArray1.First() as ComparisonUnitGroup;
            var tblGroup2 = unknown.ComparisonUnitArray2.First() as ComparisonUnitGroup;
            if (tblGroup1.Contents.Count() == tblGroup2.Contents.Count()) // if there are the same number of rows
            {
                var zipped = tblGroup1.Contents.Zip(tblGroup2.Contents, (r1, r2) => new
                {
                    Row1 = r1 as ComparisonUnitGroup,
                    Row2 = r2 as ComparisonUnitGroup,
                });
                var canCollapse = true;
                if (zipped.Any(z => z.Row1.CorrelatedSHA1Hash != z.Row2.CorrelatedSHA1Hash))
                    canCollapse = false;
                if (canCollapse)
                {
                    newListOfCorrelatedSequence = zipped
                        .Select(z =>
                        {
                            var unknownCorrelatedSequence = new CorrelatedSequence();
                            unknownCorrelatedSequence.ComparisonUnitArray1 = new[] { z.Row1 };
                            unknownCorrelatedSequence.ComparisonUnitArray2 = new[] { z.Row2 };
                            unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                            return unknownCorrelatedSequence;
                        })
                        .ToList();
                    return newListOfCorrelatedSequence;
                }
            }

            var firstContentAtom1 = tblGroup1.DescendantContentAtoms().FirstOrDefault();
            if (firstContentAtom1 == null)
                throw new OpenXmlPowerToolsException("Internal error");
            var tblElement1 = firstContentAtom1
                .AncestorElements
                .Reverse()
                .FirstOrDefault(a => a.Name == W.tbl);

            var firstContentAtom2 = tblGroup2.DescendantContentAtoms().FirstOrDefault();
            if (firstContentAtom2 == null)
                throw new OpenXmlPowerToolsException("Internal error");
            var tblElement2 = firstContentAtom2
                .AncestorElements
                .Reverse()
                .FirstOrDefault(a => a.Name == W.tbl);

            var leftContainsMerged = tblElement1
                .Descendants()
                .Any(d => d.Name == W.vMerge || d.Name == W.gridSpan);

            var rightContainsMerged = tblElement2
                .Descendants()
                .Any(d => d.Name == W.vMerge || d.Name == W.gridSpan);

            if (leftContainsMerged || rightContainsMerged)
            {
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                // If StructureSha1Hash is the same for both tables, then we know that the structure of the tables is identical, so we can break into correlated sequences for rows.
                if (tblGroup1.StructureSHA1Hash != null &&
                    tblGroup2.StructureSHA1Hash != null &&
                    tblGroup1.StructureSHA1Hash == tblGroup2.StructureSHA1Hash)
                {
                    var zipped = tblGroup1.Contents.Zip(tblGroup2.Contents, (r1, r2) => new
                    {
                        Row1 = r1 as ComparisonUnitGroup,
                        Row2 = r2 as ComparisonUnitGroup,
                    });
                    newListOfCorrelatedSequence = zipped
                        .Select(z =>
                        {
                            var unknownCorrelatedSequence = new CorrelatedSequence();
                            unknownCorrelatedSequence.ComparisonUnitArray1 = new[] { z.Row1 };
                            unknownCorrelatedSequence.ComparisonUnitArray2 = new[] { z.Row2 };
                            unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                            return unknownCorrelatedSequence;
                        })
                        .ToList();
                    return newListOfCorrelatedSequence;
                }

                // otherwise flatten to rows
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
            return null;
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
                        else if (((int)ch >= 0x4e00 && (int)ch <= 0x9fff) || settings.WordSeparators.Contains(ch))
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
                DocxComparerUtil.NotePad(sbs);
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
                DocxComparerUtil.NotePad(sbs);
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

                   return new WithHierarchicalGroupingKey()
                   {
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
                DocxComparerUtil.NotePad(sbs);
            }

            var cul = GetHierarchicalComparisonUnits(withHierarchicalGroupingKey, 0).ToArray();

            if (s_False)
            {
                var str = ComparisonUnit.ComparisonUnitListToString(cul);
                DocxComparerUtil.NotePad(str);
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
                        var childHierarchicalComparisonUnits = GetHierarchicalComparisonUnits(gc, level + 1);
                        var newCompUnitGroup = new ComparisonUnitGroup(childHierarchicalComparisonUnits, (ComparisonUnitGroupType)group, level);
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
                ElementName = VML.group,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = VML.shape,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = VML.rect,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = VML.textbox,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = O._lock,
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
                ElementName = VML.shapetype,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = W.smartTag,
                ChildElementPropertyNames = new[] { W.smartTagPr },
            },
            new RecursionInfo()
            {
                ElementName = W.ruby,
                ChildElementPropertyNames = new[] { W.rubyPr },
            },
        };

        internal static ComparisonUnitAtom[] CreateComparisonUnitAtomList(OpenXmlPart part, XElement contentParent, WmlComparerSettings settings)
        {
            VerifyNoInvalidContent(contentParent);
            AssignUnidToAllElements(contentParent);  // add the Guid id to every element
            MoveLastSectPrIntoLastParagraph(contentParent);
            var cal = CreateComparisonUnitAtomListInternal(part, contentParent, settings).ToArray();

            if (s_False)
            {
                var sb = new StringBuilder();
                foreach (var item in cal)
                    sb.Append(item.ToString() + Environment.NewLine);
                var sbs = sb.ToString();
                DocxComparerUtil.NotePad(sbs);
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
                    lastParagraph = contentParent.Descendants(W.p).LastOrDefault();
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

        private static List<ComparisonUnitAtom> CreateComparisonUnitAtomListInternal(OpenXmlPart part, XElement contentParent, WmlComparerSettings settings)
        {
            var comparisonUnitAtomList = new List<ComparisonUnitAtom>();
            CreateComparisonUnitAtomListRecurse(part, contentParent, comparisonUnitAtomList, settings);
            return comparisonUnitAtomList;
        }

        private static XName[] ComparisonGroupingElements = new[] {
            W.p,
            W.tbl,
            W.tr,
            W.tc,
            W.txbxContent,
        };

        private static void CreateComparisonUnitAtomListRecurse(OpenXmlPart part, XElement element, List<ComparisonUnitAtom> comparisonUnitAtomList, WmlComparerSettings settings)
        {
            if (element.Name == W.body || element.Name == W.footnote || element.Name == W.endnote)
            {
                foreach (var item in element.Elements())
                    CreateComparisonUnitAtomListRecurse(part, item, comparisonUnitAtomList, settings);
                return;
            }

            if (element.Name == W.p)
            {
                var paraChildrenToProcess = element
                    .Elements()
                    .Where(e => e.Name != W.pPr);
                foreach (var item in paraChildrenToProcess)
                    CreateComparisonUnitAtomListRecurse(part, item, comparisonUnitAtomList, settings);
                var paraProps = element.Element(W.pPr);
                if (paraProps == null)
                {
                    ComparisonUnitAtom pPrComparisonUnitAtom = new ComparisonUnitAtom(
                        new XElement(W.pPr),
                        element.AncestorsAndSelf().TakeWhile(a => a.Name != W.body && a.Name != W.footnotes && a.Name != W.endnotes).Reverse().ToArray(),
                        part,
                        settings);
                    comparisonUnitAtomList.Add(pPrComparisonUnitAtom);
                }
                else
                {
                    ComparisonUnitAtom pPrComparisonUnitAtom = new ComparisonUnitAtom(
                        paraProps,
                        element.AncestorsAndSelf().TakeWhile(a => a.Name != W.body && a.Name != W.footnotes && a.Name != W.endnotes).Reverse().ToArray(),
                        part,
                        settings);
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
                    CreateComparisonUnitAtomListRecurse(part, item, comparisonUnitAtomList, settings);
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
                        part,
                        settings);
                    comparisonUnitAtomList.Add(sr);
                }
                return;
            }

            if (AllowableRunChildren.Contains(element.Name) || element.Name == W._object)
            {
                ComparisonUnitAtom sr3 = new ComparisonUnitAtom(
                    element,
                    element.AncestorsAndSelf().TakeWhile(a => a.Name != W.body && a.Name != W.footnotes && a.Name != W.endnotes).Reverse().ToArray(),
                    part,
                    settings);
                comparisonUnitAtomList.Add(sr3);
                return;
            }

            var re = RecursionElements.FirstOrDefault(z => z.ElementName == element.Name);
            if (re != null)
            {
                AnnotateElementWithProps(part, element, comparisonUnitAtomList, re.ChildElementPropertyNames, settings);
                return;
            }

            if (ElementsToThrowAway.Contains(element.Name))
                return;

            AnnotateElementWithProps(part, element, comparisonUnitAtomList, null, settings);
        }

        private static void AnnotateElementWithProps(OpenXmlPart part, XElement element, List<ComparisonUnitAtom> comparisonUnitAtomList, XName[] childElementPropertyNames, WmlComparerSettings settings)
        {
            IEnumerable<XElement> runChildrenToProcess = null;
            if (childElementPropertyNames == null)
                runChildrenToProcess = element.Elements();
            else
                runChildrenToProcess = element
                    .Elements()
                    .Where(e => !childElementPropertyNames.Contains(e.Name));

            foreach (var item in runChildrenToProcess)
                CreateComparisonUnitAtomListRecurse(part, item, comparisonUnitAtomList, settings);
        }



        private static void AssignUnidToAllElements(XElement contentParent)
        {
            var content = contentParent.Descendants();
            foreach (var d in content)
            {
                if (d.Attribute(PtOpenXml.Unid) == null)
                {
                    string unid = Guid.NewGuid().ToString().Replace("-", "");
                    var newAtt = new XAttribute(PtOpenXml.Unid, unid);
                    d.Add(newAtt);
                }
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

        private int? m_DescendantContentAtomsCount = null;

        public int DescendantContentAtomsCount
        {
            get
            {
                if (m_DescendantContentAtomsCount != null)
                    return (int)m_DescendantContentAtomsCount;
                m_DescendantContentAtomsCount = this.DescendantContentAtoms().Count();
                return (int)m_DescendantContentAtomsCount;
            }
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
            SHA1Hash = PtUtils.SHA1HashStringForUTF8String(sha1String);
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

    public class ComparisonUnitAtom : ComparisonUnit
    {
        // AncestorElements are kept in order from the body to the leaf, because this is the order in which we need to access in order
        // to reassemble the document.  However, in many places in the code, it is necessary to find the nearest ancestor, i.e. cell
        // so it is necessary to reverse the order when looking for it, i.e. look from the leaf back to the body element.

        public XElement[] AncestorElements;
        public string[] AncestorUnids;
        public XElement ContentElement;
        public XElement ContentElementBefore;
        public ComparisonUnitAtom ComparisonUnitAtomBefore;
        public OpenXmlPart Part;
        public XElement RevTrackElement;

        public ComparisonUnitAtom(XElement contentElement, XElement[] ancestorElements, OpenXmlPart part, WmlComparerSettings settings)
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
                var shaHashString = GetSha1HashStringForElement(ContentElement, settings);
                SHA1Hash = PtUtils.SHA1HashStringForUTF8String(shaHashString);
            }
        }

        private string GetSha1HashStringForElement(XElement contentElement, WmlComparerSettings settings)
        {
            var text = contentElement.Value;
            if (settings.CaseInsensitive)
                text = text.ToUpper(settings.CultureInfo);
            if (settings.ConflateBreakingAndNonbreakingSpaces)
                text = text.Replace(' ', '\x00a0');
            return contentElement.Name.LocalName + text;
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

        public string ToStringAncestorUnids(int indent)
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
                AppendAncestorsUnidsDump(sb, this);
            }
            else
            {
                sb.AppendFormat("Atom {0}:   {1} SHA1:{2} ", PadLocalName(xNamePad, this), correlationStatus, this.SHA1Hash.Substring(0, 8));
                AppendAncestorsUnidsDump(sb, this);
            }
            return sb.ToString();
        }

        public override string ToString()
        {
            return ToString(0);
        }

        public string ToStringAncestorUnids()
        {
            return ToStringAncestorUnids(0);
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

        private void AppendAncestorsUnidsDump(StringBuilder sb, ComparisonUnitAtom sr)
        {
            var zipped = sr.AncestorElements.Zip(sr.AncestorUnids, (a, u) => new
            {
                AncestorElement = a,
                AncestorUnid = u,
            });
            var s = zipped.Select(p => p.AncestorElement.Name.LocalName + "[" + p.AncestorUnid.Substring(0, 8) + "]/").StringConcatenate().TrimEnd('/');
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
        public string CorrelatedSHA1Hash;
        public string StructureSHA1Hash;

        public ComparisonUnitGroup(IEnumerable<ComparisonUnit> comparisonUnitList, ComparisonUnitGroupType groupType, int level)
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

            var ancestorsToLookAt = comparisonUnitAtom.AncestorElements.Where(ae => ae.Name == W.tbl || ae.Name == W.tr || ae.Name == W.tc || ae.Name == W.p || ae.Name == W.txbxContent).ToArray(); ;
            var ancestor = ancestorsToLookAt[level];

            if (ancestor == null)
                throw new OpenXmlPowerToolsException("Internal error: ComparisonUnitGroup");
            SHA1Hash = (string)ancestor.Attribute(PtOpenXml.SHA1Hash);
            CorrelatedSHA1Hash = (string)ancestor.Attribute(PtOpenXml.CorrelatedSHA1Hash);
            StructureSHA1Hash = (string)ancestor.Attribute(PtOpenXml.StructureSHA1Hash);
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

    internal class DocxComparerUtil
    {
        public static void NotePad(string str)
        {
            var tempPath = Path.GetTempPath();
            var guidName = Guid.NewGuid().ToString().Replace("-", "") + ".txt";
            var fi = new FileInfo(Path.Combine(tempPath, guidName));
            File.WriteAllText(fi.FullName, str);
            var notepadExe = new FileInfo(@"C:\Program Files (x86)\Notepad++\notepad++.exe");
            if (!notepadExe.Exists)
                notepadExe = new FileInfo(@"C:\Program Files\Notepad++\notepad++.exe");
            if (!notepadExe.Exists)
                notepadExe = new FileInfo(@"C:\Windows\System32\notepad.exe");
            ExecutableRunner.RunExecutable(notepadExe.FullName, fi.FullName, tempPath);
        }
    }

#if false
    public class PtpSHA1Util
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

    public class Base64Util
    {
        private class Bs64Tupple
        {
            public char Bs64Character;
            public int Bs64Chunk;
        }

        public static string Convert76CharLineLength(byte[] byteArray)
        {
            string base64String = (System.Convert.ToBase64String(byteArray))
                .Select
                (
                    (c, i) => new Bs64Tupple()
                    {
                        Bs64Character = c,
                        Bs64Chunk = i / 76
                    }
                )
                .GroupBy(c => c.Bs64Chunk)
                .Aggregate(
                    new StringBuilder(),
                    (s, i) =>
                        s.Append(
                            i.Aggregate(
                                new StringBuilder(),
                                (seed, it) => seed.Append(it.Bs64Character),
                                sb => sb.ToString()
                            )
                        )
                        .Append(Environment.NewLine),
                    s => s.ToString()
                );
            return base64String;
        }
    }
#endif
}
