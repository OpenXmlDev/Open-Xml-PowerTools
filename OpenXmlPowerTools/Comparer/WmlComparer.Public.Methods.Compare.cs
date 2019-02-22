// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

// It is possible to optimize DescendantContentAtoms

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Currently, the unid is set at the beginning of the algorithm.  It is used by the code that establishes correlation
// based on first rejecting// tracked revisions, then correlating paragraphs/tables.  It is requred for this algorithm
// - after finding a correlated sequence in the document with rejected revisions, it uses the unid to find the same
// paragraph in the document without rejected revisions, then sets the correlated sha1 hash in that document.
//
// But then when accepting tracked revisions, for certain paragraphs (where there are deleted paragraph marks) it is
// going to lose the unids.  But this isn't a problem because when paragraph marks are deleted, the correlation is
// definitely no longer possible.  Any paragraphs that are in a range of paragraphs that are coalesced can't be
// correlated to paragraphs in the other document via their hash.  At that point we no longer care what their unids
// are.
//
// But after that it is only used to reconstruct the tree.  It is also used in the debugging code that
// prints the various correlated sequences and comparison units - this is display for debugging purposes only.
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// The key idea here is that a given paragraph will always have the same ancestors, and it doesn't matter whether the
// content was deleted from the old document, inserted into the new document, or set as equal.  At this point, we
// identify a paragraph as a sequential list of content atoms, terminated by a paragraph mark.  This entire list will
// for a single paragraph, regardless of whether the paragraph is a child of the body, or if the paragraph is in a cell
// in a table, or if the paragraph is in a text box.  The list of ancestors, from the paragraph to the root of the XML
// tree will be the same for all content atoms in the paragraph.
//
// Therefore:
//
// Iterate through the list of content atoms backwards.  When the loop sees a paragraph mark, it gets the ancestor
// unids from the paragraph mark to the top of the tree, and sets this as the same for all content atoms in the
// paragraph.  For descendants of the paragraph mark, it doesn't really matter if content is put into separate runs
// or what not.  We don't need to be concerned about what the unids are for descendants of the paragraph.
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

namespace OpenXmlPowerTools
{
    public static partial class WmlComparer
    {
        public static WmlDocument Compare(WmlDocument source1, WmlDocument source2, WmlComparerSettings settings)
        {
            return CompareInternal(source1, source2, settings, true);
        }

        private static WmlDocument CompareInternal(
            WmlDocument source1,
            WmlDocument source2,
            WmlComparerSettings settings,
            bool preProcessMarkupInOriginal)
        {
            if (preProcessMarkupInOriginal)
            {
                source1 = PreProcessMarkup(source1, settings.StartingIdForFootnotesEndnotes + 1000);
            }

            source2 = PreProcessMarkup(source2, settings.StartingIdForFootnotesEndnotes + 2000);

            SaveDocumentIfDesired(source1, "Source1-Step1-PreProcess.docx", settings);
            SaveDocumentIfDesired(source2, "Source2-Step1-PreProcess.docx", settings);

            // at this point, both source1 and source2 have unid on every element.  These are the values that will
            // enable reassembly of the XML tree.  But we need other values.

            // In source1:
            // - accept tracked revisions
            // - determine hash code for every block-level element
            // - save as attribute on every element

            // - accept tracked revisions and reject tracked revisions leave the unids alone, where possible.
            // - after accepting and calculating the hash, then can use the unids to find the right block-level
            //   element in the unmodified source1, and install the hash

            // In source2:
            // - reject tracked revisions
            // - determine hash code for every block-level element
            // - save as an attribute on every element

            // - after rejecting and calculating the hash, then can use the unids to find the right block-level element
            //   in the unmodified source2, and install the hash

            // - sometimes after accepting or rejecting tracked revisions, several paragraphs will get coalesced into a
            //   single paragraph due to paragraph marks being inserted / deleted.
            // - in this case, some paragraphs will not get a hash injected onto them.
            // - if a paragraph doesn't have a hash, then it will never correspond to another paragraph, and such
            //   issues will need to be resolved in the normal execution of the LCS algorithm.
            // - note that when we do propagate the unid through for the first paragraph.

            // Establish correlation between the two.
            // Find the longest common sequence of block-level elements where hash codes are the same.
            // this sometimes will be every block level element in the document.  Or sometimes will be just a fair
            // number of them.

            // at the start of doing the LCS algorithm, we will match up content, and put them in corresponding unknown
            // correlated comparison units.  Those paragraphs will only ever be matched to their corresponding paragraph.
            // then the algorithm can proceed as usual.

            // need to call ChangeFootnoteEndnoteReferencesToUniqueRange before creating the wmlResult document, so that
            // the same GUID ids are used for footnote and endnote references in both the 'after' document, and in the
            // result document.

            WmlDocument source1AfterAccepting = RevisionProcessor.AcceptRevisions(source1);
            WmlDocument source2AfterRejecting = RevisionProcessor.RejectRevisions(source2);

            SaveDocumentIfDesired(source1AfterAccepting, "Source1-Step2-AfterAccepting.docx", settings);
            SaveDocumentIfDesired(source2AfterRejecting, "Source2-Step2-AfterRejecting.docx", settings);

            // this creates the correlated hash codes that enable us to match up ranges of paragraphs based on
            // accepting in source1, rejecting in source2
            source1 = HashBlockLevelContent(source1, source1AfterAccepting, settings);
            source2 = HashBlockLevelContent(source2, source2AfterRejecting, settings);

            SaveDocumentIfDesired(source1, "Source1-Step3-AfterHashing.docx", settings);
            SaveDocumentIfDesired(source2, "Source2-Step3-AfterHashing.docx", settings);

            // Accept revisions in before, and after
            source1 = RevisionProcessor.AcceptRevisions(source1);
            source2 = RevisionProcessor.AcceptRevisions(source2);

            SaveDocumentIfDesired(source1, "Source1-Step4-AfterAccepting.docx", settings);
            SaveDocumentIfDesired(source2, "Source2-Step4-AfterAccepting.docx", settings);

            // after accepting revisions, some unids may have been removed by revision accepter, along with the
            // correlatedSHA1Hash codes, this is as it should be.
            // but need to go back in and add guids to paragraphs that have had them removed.

            using (var ms = new MemoryStream())
            {
                ms.Write(source2.DocumentByteArray, 0, source2.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    AddUnidsToMarkupInContentParts(wDoc);
                }
            }

            var wmlResult = new WmlDocument(source1);
            using (var ms1 = new MemoryStream())
            using (var ms2 = new MemoryStream())
            {
                ms1.Write(source1.DocumentByteArray, 0, source1.DocumentByteArray.Length);
                ms2.Write(source2.DocumentByteArray, 0, source2.DocumentByteArray.Length);
                WmlDocument producedDocument;

                using (WordprocessingDocument wDoc1 = WordprocessingDocument.Open(ms1, true))
                using (WordprocessingDocument wDoc2 = WordprocessingDocument.Open(ms2, true))
                {
                    producedDocument = ProduceDocumentWithTrackedRevisions(settings, wmlResult, wDoc1, wDoc2);
                }

                SaveDocumentsAfterProducingDocument(ms1, ms2, settings);
                SaveCleanedDocuments(source1, producedDocument, settings);

                return producedDocument;
            }
        }

        private static void SaveDocumentIfDesired(WmlDocument source, string name, WmlComparerSettings settings)
        {
            if (SaveIntermediateFilesForDebugging && settings.DebugTempFileDi != null)
            {
                var fileInfo = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name));
                source.SaveAs(fileInfo.FullName);
            }
        }

        private static void SaveDocumentsAfterProducingDocument(MemoryStream ms1, MemoryStream ms2, WmlComparerSettings settings)
        {
            if (SaveIntermediateFilesForDebugging && settings.DebugTempFileDi != null)
            {
                SaveDocumentIfDesired(new WmlDocument("after1.docx", ms1), "Source1-Step5-AfterProducingDocument.docx", settings);
                SaveDocumentIfDesired(new WmlDocument("after2.docx", ms2), "Source2-Step5-AfterProducingDocument.docx", settings);
            }
        }

        private static void SaveCleanedDocuments(WmlDocument source1, WmlDocument producedDocument, WmlComparerSettings settings)
        {
            if (SaveIntermediateFilesForDebugging && settings.DebugTempFileDi != null)
            {
                WmlDocument cleanedSource = CleanPowerToolsAndRsid(source1);
                SaveDocumentIfDesired(cleanedSource, "Cleaned-Source.docx", settings);

                WmlDocument cleanedProduced = CleanPowerToolsAndRsid(producedDocument);
                SaveDocumentIfDesired(cleanedProduced, "Cleaned-Produced.docx", settings);
            }
        }

        private static WmlDocument CleanPowerToolsAndRsid(WmlDocument producedDocument)
        {
            using (var ms = new MemoryStream())
            {
                ms.Write(producedDocument.DocumentByteArray, 0, producedDocument.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    foreach (OpenXmlPart cp in wDoc.ContentParts())
                    {
                        XDocument xd = cp.GetXDocument();
                        object newRoot = CleanPartTransform(xd.Root);
                        xd.Root?.ReplaceWith(newRoot);
                        cp.PutXDocument();
                    }
                }

                var cleaned = new WmlDocument("cleaned.docx", ms.ToArray());
                return cleaned;
            }
        }

        private static object CleanPartTransform(XNode node)
        {
            if (node is XElement element)
            {
                return new XElement(element.Name,
                    element.Attributes().Where(a => a.Name.Namespace != PtOpenXml.pt &&
                                                    !a.Name.LocalName.ToLower().Contains("rsid")),
                    element.Nodes().Select(CleanPartTransform));
            }

            return node;
        }
    }
}
