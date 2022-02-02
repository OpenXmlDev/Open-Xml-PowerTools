// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public static partial class WmlComparer
    {
        private static WmlDocument ProduceDocumentWithTrackedRevisions(
            WmlComparerSettings settings,
            WmlDocument wmlResult,
            WordprocessingDocument wDoc1,
            WordprocessingDocument wDoc2)
        {
            // save away sectPr so that can set in the newly produced document.
            XElement savedSectPr = wDoc1
                .MainDocumentPart
                .GetXDocument()
                .Root?
                .Element(W.body)?
                .Element(W.sectPr);

            XElement contentParent1 = wDoc1.MainDocumentPart.GetXDocument().Root?.Element(W.body);
            AddSha1HashToBlockLevelContent(wDoc1.MainDocumentPart, contentParent1, settings);

            XElement contentParent2 = wDoc2.MainDocumentPart.GetXDocument().Root?.Element(W.body);
            AddSha1HashToBlockLevelContent(wDoc2.MainDocumentPart, contentParent2, settings);

            ComparisonUnitAtom[] cal1 = CreateComparisonUnitAtomList(
                wDoc1.MainDocumentPart,
                wDoc1.MainDocumentPart.GetXDocument().Root?.Element(W.body),
                settings);

            if (False)
            {
                var sb = new StringBuilder();
                foreach (ComparisonUnitAtom item in cal1)
                    sb.Append(item + Environment.NewLine);

                string sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            ComparisonUnit[] cus1 = GetComparisonUnitList(cal1, settings);

            if (False)
            {
                string sbs = ComparisonUnit.ComparisonUnitListToString(cus1);
                TestUtil.NotePad(sbs);
            }

            ComparisonUnitAtom[] cal2 = CreateComparisonUnitAtomList(
                wDoc2.MainDocumentPart,
                wDoc2.MainDocumentPart.GetXDocument().Root?.Element(W.body),
                settings);

            if (False)
            {
                var sb = new StringBuilder();
                foreach (ComparisonUnitAtom item in cal2)
                    sb.Append(item + Environment.NewLine);

                string sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            ComparisonUnit[] cus2 = GetComparisonUnitList(cal2, settings);

            if (False)
            {
                string sbs = ComparisonUnit.ComparisonUnitListToString(cus2);
                TestUtil.NotePad(sbs);
            }

            if (False)
            {
                var sb3 = new StringBuilder();
                sb3.Append("ComparisonUnitList 1 =====" + Environment.NewLine + Environment.NewLine);
                sb3.Append(ComparisonUnit.ComparisonUnitListToString(cus1));
                sb3.Append(Environment.NewLine);
                sb3.Append("ComparisonUnitList 2 =====" + Environment.NewLine + Environment.NewLine);
                sb3.Append(ComparisonUnit.ComparisonUnitListToString(cus2));
                string sbs3 = sb3.ToString();
                TestUtil.NotePad(sbs3);
            }

            List<CorrelatedSequence> correlatedSequence = Lcs(cus1, cus2, settings);

            if (False)
            {
                var sb = new StringBuilder();
                foreach (CorrelatedSequence item in correlatedSequence)
                {
                    sb.Append(item + Environment.NewLine);
                }

                string sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            // for any deleted or inserted rows, we go into the w:trPr properties, and add the appropriate w:ins or
            // w:del element, and therefore when generating the document, the appropriate row will be marked as deleted
            // or inserted.
            MarkRowsAsDeletedOrInserted(settings, correlatedSequence);

            // the following gets a flattened list of ComparisonUnitAtoms, with status indicated in each
            // ComparisonUnitAtom: Deleted, Inserted, or Equal
            List<ComparisonUnitAtom> listOfComparisonUnitAtoms = FlattenToComparisonUnitAtomList(correlatedSequence, settings);

            if (False)
            {
                var sb = new StringBuilder();
                foreach (ComparisonUnitAtom item in listOfComparisonUnitAtoms)
                {
                    sb.Append(item + Environment.NewLine);
                }

                string sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            // note - we don't want to do the hack until after flattening all of the groups.  At the end of the
            // flattening, we should simply have a list of ComparisonUnitAtoms, appropriately marked as equal,
            // inserted, or deleted.

            // the table id will be hacked in the normal course of events.
            // in the case where a row is deleted, not necessary to hack - the deleted row ID will do.
            // in the case where a row is inserted, not necessary to hack - the inserted row ID will do as well.
            AssembleAncestorUnidsInOrderToRebuildXmlTreeProperly(listOfComparisonUnitAtoms);

            if (False)
            {
                var sb = new StringBuilder();
                foreach (ComparisonUnitAtom item in listOfComparisonUnitAtoms)
                    sb.Append(item.ToStringAncestorUnids() + Environment.NewLine);

                string sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            // and then finally can generate the document with revisions
            using (var ms = new MemoryStream())
            {
                ms.Write(wmlResult.DocumentByteArray, 0, wmlResult.DocumentByteArray.Length);
                using (WordprocessingDocument wDocWithRevisions = WordprocessingDocument.Open(ms, true))
                {
                    XDocument xDoc = wDocWithRevisions.MainDocumentPart.GetXDocument();
                    List<XAttribute> rootNamespaceAttributes = xDoc
                        .Root?
                        .Attributes()
                        .Where(a => a.IsNamespaceDeclaration || a.Name.Namespace == MC.mc)
                        .ToList();

                    // ======================================
                    // The following produces a new valid WordprocessingML document from the listOfComparisonUnitAtoms
                    object newBodyChildren = ProduceNewWmlMarkupFromCorrelatedSequence(
                        wDocWithRevisions.MainDocumentPart,
                        listOfComparisonUnitAtoms,
                        settings);

                    var newXDoc = new XDocument();
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
                    var newXDoc2Root = (XElement) WordprocessingMLUtil.WmlOrderElementsPerStandard(newXDoc.Root);
                    xDoc.Root?.ReplaceWith(newXDoc2Root);

                    /**********************************************************************************************/
                    // temporary code to remove sections.  When remove this code, get validation errors for some ITU documents.
                    // Note: This is a no-go for use cases in which documents have multiple sections, e.g., for title pages,
                    // front matter, and body matter. Another example is where you have to switch between portrait and
                    // landscape orientation, which requires sections.
                    // TODO: Revisit
                    xDoc.Root?.Descendants(W.sectPr).Remove();

                    // move w:sectPr from source document into newly generated document.
                    if (savedSectPr != null)
                    {
                        XDocument xd = wDocWithRevisions.MainDocumentPart.GetXDocument();

                        // add everything but headers/footers
                        var clonedSectPr = new XElement(W.sectPr,
                            savedSectPr.Attributes(),
                            savedSectPr.Element(W.type),
                            savedSectPr.Element(W.pgSz),
                            savedSectPr.Element(W.pgMar),
                            savedSectPr.Element(W.cols),
                            savedSectPr.Element(W.titlePg));
                        xd.Root?.Element(W.body)?.Add(clonedSectPr);
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

                foreach (OpenXmlPart part in wDoc1.ContentParts())
                {
                    part.PutXDocument();
                }

                foreach (OpenXmlPart part in wDoc2.ContentParts())
                {
                    part.PutXDocument();
                }

                var updatedWmlResult = new WmlDocument("Dummy.docx", ms.ToArray());
                return updatedWmlResult;
            }
        }

        private static void AddSha1HashToBlockLevelContent(OpenXmlPart part, XElement contentParent, WmlComparerSettings settings)
        {
            IEnumerable<XElement> blockLevelContentToAnnotate = contentParent
                .Descendants()
                .Where(d => ElementsToHaveSha1Hash.Contains(d.Name));

            foreach (XElement blockLevelContent in blockLevelContentToAnnotate)
            {
                var cloneBlockLevelContentForHashing =
                    (XElement) CloneBlockLevelContentForHashing(part, blockLevelContent, true, settings);
                string shaString = cloneBlockLevelContentForHashing.ToString(SaveOptions.DisableFormatting)
                    .Replace(" xmlns=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", "");
                string sha1Hash = WmlComparerUtil.SHA1HashStringForUTF8String(shaString);
                blockLevelContent.Add(new XAttribute(PtOpenXml.SHA1Hash, sha1Hash));

                if (blockLevelContent.Name == W.tbl ||
                    blockLevelContent.Name == W.tr)
                {
                    var clonedForStructureHash = (XElement) CloneForStructureHash(cloneBlockLevelContentForHashing);

                    // this is a convenient place to look at why tables are being compared as different.

                    //if (blockLevelContent.Name == W.tbl)
                    //    Console.WriteLine();

                    string shaString2 = clonedForStructureHash.ToString(SaveOptions.DisableFormatting)
                        .Replace(" xmlns=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", "");
                    string sha1Hash2 = WmlComparerUtil.SHA1HashStringForUTF8String(shaString2);
                    blockLevelContent.Add(new XAttribute(PtOpenXml.StructureSHA1Hash, sha1Hash2));
                }
            }
        }

        private static List<CorrelatedSequence> Lcs(ComparisonUnit[] cu1, ComparisonUnit[] cu2, WmlComparerSettings settings)
        {
            // set up initial state - one CorrelatedSequence, UnKnown, contents == entire sequences (both)
            var cs = new CorrelatedSequence
            {
                CorrelationStatus = CorrelationStatus.Unknown,
                ComparisonUnitArray1 = cu1,
                ComparisonUnitArray2 = cu2
            };
            var csList = new List<CorrelatedSequence>
            {
                cs
            };

            while (true)
            {
                if (False)
                {
                    var sb = new StringBuilder();
                    foreach (CorrelatedSequence item in csList)
                        sb.Append(item).Append(Environment.NewLine);
                    string sbs = sb.ToString();
                    TestUtil.NotePad(sbs);
                }

                CorrelatedSequence unknown = csList
                    .FirstOrDefault(z => z.CorrelationStatus == CorrelationStatus.Unknown);

                if (unknown != null)
                {
                    // if unknown consists of a single group of the same type in each side, then can set some Unids in the 'after' document.
                    // if the unknown is a pair of single tables, then can set table Unid.
                    // if the unknown is a pair of single rows, then can set table and rows Unids.
                    // if the unknown is a pair of single cells, then can set table, row, and cell Unids.
                    // if the unknown is a pair of paragraphs, then can set paragraph (and all ancestor) Unids.
                    SetAfterUnids(unknown);

                    if (False)
                    {
                        var sb = new StringBuilder();
                        sb.Append(unknown);
                        string sbs = sb.ToString();
                        TestUtil.NotePad(sbs);
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

                    int indexOfUnknown = csList.IndexOf(unknown);
                    csList.Remove(unknown);

                    newSequence.Reverse();
                    foreach (CorrelatedSequence item in newSequence)
                        csList.Insert(indexOfUnknown, item);

                    continue;
                }

                return csList;
            }
        }

        private static void MarkRowsAsDeletedOrInserted(WmlComparerSettings settings, List<CorrelatedSequence> correlatedSequence)
        {
            foreach (CorrelatedSequence dcs in correlatedSequence.Where(cs =>
                cs.CorrelationStatus == CorrelationStatus.Deleted || cs.CorrelationStatus == CorrelationStatus.Inserted))
            {
                // iterate through all deleted/inserted items in dcs.ComparisonUnitArray1/ComparisonUnitArray2
                ComparisonUnit[] toIterateThrough = dcs.ComparisonUnitArray1;
                if (dcs.CorrelationStatus == CorrelationStatus.Inserted)
                    toIterateThrough = dcs.ComparisonUnitArray2;

                foreach (ComparisonUnit ca in toIterateThrough)
                {
                    var cug = ca as ComparisonUnitGroup;

                    // this works because we will never see a table in this list, only rows.  If tables were in this list, would need to recursively
                    // go into children, but tables are always flattened in the LCS process.

                    // when we have a row, it is only necessary to find the first content atom of the row, then find the row ancestor, and then tweak
                    // the w:trPr

                    if (cug != null && cug.ComparisonUnitGroupType == ComparisonUnitGroupType.Row)
                    {
                        ComparisonUnitAtom firstContentAtom = cug.DescendantContentAtoms().FirstOrDefault();
                        if (firstContentAtom == null)
                            throw new OpenXmlPowerToolsException("Internal error");

                        XElement tr = firstContentAtom
                            .AncestorElements
                            .Reverse()
                            .FirstOrDefault(a => a.Name == W.tr);

                        if (tr == null)
                            throw new OpenXmlPowerToolsException("Internal error");

                        XElement trPr = tr.Element(W.trPr);
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
                            new XAttribute(W.id, _maxId++),
                            new XAttribute(W.date, settings.DateTimeForRevisions)));
                    }
                }
            }
        }

        private static List<ComparisonUnitAtom> FlattenToComparisonUnitAtomList(
            List<CorrelatedSequence> correlatedSequence,
            WmlComparerSettings settings)
        {
            List<ComparisonUnitAtom> listOfComparisonUnitAtoms = correlatedSequence
                .Select(cs =>
                {
                    // need to write some code here to find out if we are assembling a paragraph (or anything) that contains the following unid.
                    // why do are we dropping content???????
                    //string searchFor = "0ecb9184";


                    if (cs.CorrelationStatus == CorrelationStatus.Equal)
                    {
                        IEnumerable<ComparisonUnitAtom> contentAtomsBefore = cs
                            .ComparisonUnitArray1
                            .Select(ca => ca.DescendantContentAtoms())
                            .SelectMany(m => m);

                        IEnumerable<ComparisonUnitAtom> contentAtomsAfter = cs
                            .ComparisonUnitArray2
                            .Select(ca => ca.DescendantContentAtoms())
                            .SelectMany(m => m);

                        List<ComparisonUnitAtom> comparisonUnitAtomList = contentAtomsBefore
                            .Zip(contentAtomsAfter,
                                (before, after) => new ComparisonUnitAtom(
                                    after.ContentElement,
                                    after.AncestorElements,
                                    after.Part,
                                    settings)
                                {
                                    CorrelationStatus = CorrelationStatus.Equal,
                                    ContentElementBefore = before.ContentElement,
                                    ComparisonUnitAtomBefore = before
                                })
                            .ToList();

                        return comparisonUnitAtomList;
                    }

                    if (cs.CorrelationStatus == CorrelationStatus.Deleted)
                    {
                        IEnumerable<ComparisonUnitAtom> comparisonUnitAtomList = cs
                            .ComparisonUnitArray1
                            .Select(ca => ca.DescendantContentAtoms())
                            .SelectMany(m => m)
                            .Select(ca =>
                                new ComparisonUnitAtom(ca.ContentElement, ca.AncestorElements, ca.Part, settings)
                                {
                                    CorrelationStatus = CorrelationStatus.Deleted
                                });

                        return comparisonUnitAtomList;
                    }

                    if (cs.CorrelationStatus == CorrelationStatus.Inserted)
                    {
                        IEnumerable<ComparisonUnitAtom> comparisonUnitAtomList = cs
                            .ComparisonUnitArray2
                            .Select(ca => ca.DescendantContentAtoms())
                            .SelectMany(m => m)
                            .Select(ca =>
                                new ComparisonUnitAtom(ca.ContentElement, ca.AncestorElements, ca.Part, settings)
                                {
                                    CorrelationStatus = CorrelationStatus.Inserted
                                });
                        return comparisonUnitAtomList;
                    }

                    throw new OpenXmlPowerToolsException("Internal error");
                })
                .SelectMany(m => m)
                .ToList();

            if (False)
            {
                var sb = new StringBuilder();
                foreach (ComparisonUnitAtom item in listOfComparisonUnitAtoms)
                    sb.Append(item).Append(Environment.NewLine);
                string sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            return listOfComparisonUnitAtoms;
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
            if (False)
            {
                var sb = new StringBuilder();
                foreach (ComparisonUnitAtom item in comparisonUnitAtomList)
                    sb.Append(item).Append(Environment.NewLine);
                string sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            // the following loop sets all ancestor unids in the after document to the unids in the before document for all pPr where the status is equal.
            // this should always be true.

            // one additional modification to make to this loop - where we find a pPr in a text box, we want to do this as well, regardless of whether the status is equal, inserted, or deleted.
            // reason being that this module does not support insertion / deletion of text boxes themselves.  If a text box is in the before or after document, it will be in the document that
            // contains deltas.  It may have inserted or deleted text, but regardless, it will be in the result document.
            foreach (ComparisonUnitAtom cua in comparisonUnitAtomList)
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
                    ComparisonUnitAtom cuaBefore = cua.ComparisonUnitAtomBefore;
                    XElement[] ancestorsAfter = cua.AncestorElements;
                    if (cuaBefore != null)
                    {
                        XElement[] ancestorsBefore = cuaBefore.AncestorElements;
                        if (ancestorsAfter.Length == ancestorsBefore.Length)
                        {
                            var zipped = ancestorsBefore.Zip(ancestorsAfter, (b, a) =>
                                new
                                {
                                    After = a,
                                    Before = b
                                });

                            foreach (var z in zipped)
                            {
                                XAttribute afterUnidAtt = z.After.Attribute(PtOpenXml.Unid);
                                XAttribute beforeUnidAtt = z.Before.Attribute(PtOpenXml.Unid);
                                if (afterUnidAtt != null && beforeUnidAtt != null)
                                    afterUnidAtt.Value = beforeUnidAtt.Value;
                            }
                        }
                    }
                }
            }

            if (False)
            {
                var sb = new StringBuilder();
                foreach (ComparisonUnitAtom item in comparisonUnitAtomList)
                    sb.Append(item).Append(Environment.NewLine);
                string sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            List<ComparisonUnitAtom> rComparisonUnitAtomList =
                ((IEnumerable<ComparisonUnitAtom>) comparisonUnitAtomList).Reverse().ToList();

            // the following should always succeed, because there will always be at least one element in
            // rComparisonUnitAtomList, and there will always be at least one ancestor in AncestorElements
            XElement deepestAncestor = rComparisonUnitAtomList.First().AncestorElements.First();
            XName deepestAncestorName = deepestAncestor.Name;
            string deepestAncestorUnid = null;
            if (deepestAncestorName == W.footnote || deepestAncestorName == W.endnote)
            {
                deepestAncestorUnid = (string) deepestAncestor.Attribute(PtOpenXml.Unid);
            }

            // If the following loop finds a pPr that is in a text box, then continue on, processing the pPr and all of its contents as though it were
            // content in the containing text box.  This is going to leave it after this loop where the AncestorUnids for the content in the text box will be
            // incomplete.  We then will need to go through the rComparisonUnitAtomList a second time, processing all of the text boxes.

            // Note that this makes the basic assumption that a text box can't be nested inside of a text box, which, as far as I know, is a good assumption.

            // This also makes the basic assumption that an endnote / footnote can't contain a text box, which I believe is a good assumption.


            string[] currentAncestorUnids = null;
            foreach (ComparisonUnitAtom cua in rComparisonUnitAtomList)
            {
                if (cua.ContentElement.Name == W.pPr)
                {
                    bool pPr_inTextBox = cua
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
                                var thisUnid = (string) ae.Attribute(PtOpenXml.Unid);
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

                int thisDepth = cua.AncestorElements.Length;
                IEnumerable<string> additionalAncestorUnids = cua
                    .AncestorElements
                    .Skip(currentAncestorUnids.Length)
                    .Select(ae =>
                    {
                        var thisUnid = (string) ae.Attribute(PtOpenXml.Unid);
                        if (thisUnid == null)
                            Guid.NewGuid().ToString().Replace("-", "");
                        return thisUnid;
                    });
                string[] thisAncestorUnids = currentAncestorUnids
                    .Concat(additionalAncestorUnids)
                    .ToArray();
                cua.AncestorUnids = thisAncestorUnids;
                if (deepestAncestorUnid != null)
                    cua.AncestorUnids[0] = deepestAncestorUnid;
            }

            if (False)
            {
                var sb = new StringBuilder();
                foreach (ComparisonUnitAtom item in comparisonUnitAtomList)
                    sb.Append(item).Append(Environment.NewLine);
                string sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            // this is the second loop that processes all text boxes.
            currentAncestorUnids = null;
            var skipUntilNextPpr = false;
            foreach (ComparisonUnitAtom cua in rComparisonUnitAtomList)
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
                    //    TestUtil.NotePad(sbs);
                    //}

                    bool pPr_inTextBox = cua
                        .AncestorElements
                        .Any(ae => ae.Name == W.txbxContent);

                    if (!pPr_inTextBox)
                    {
                        skipUntilNextPpr = true;
                        currentAncestorUnids = null;
                        continue;
                    }

                    skipUntilNextPpr = false;

                    currentAncestorUnids = cua
                        .AncestorElements
                        .Select(ae =>
                        {
                            var thisUnid = (string) ae.Attribute(PtOpenXml.Unid);
                            if (thisUnid == null)
                                throw new OpenXmlPowerToolsException("Internal error");

                            return thisUnid;
                        })
                        .ToArray();
                    cua.AncestorUnids = currentAncestorUnids;
                    continue;
                }

                if (skipUntilNextPpr)
                    continue;

                int thisDepth = cua.AncestorElements.Length;
                IEnumerable<string> additionalAncestorUnids = cua
                    .AncestorElements
                    .Skip(currentAncestorUnids.Length)
                    .Select(ae =>
                    {
                        var thisUnid = (string) ae.Attribute(PtOpenXml.Unid);
                        if (thisUnid == null)
                            Guid.NewGuid().ToString().Replace("-", "");
                        return thisUnid;
                    });
                string[] thisAncestorUnids = currentAncestorUnids
                    .Concat(additionalAncestorUnids)
                    .ToArray();
                cua.AncestorUnids = thisAncestorUnids;
            }

            if (False)
            {
                var sb = new StringBuilder();
                foreach (ComparisonUnitAtom item in comparisonUnitAtomList)
                    sb.Append(item.ToStringAncestorUnids()).Append(Environment.NewLine);
                string sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }
        }

        private static object ProduceNewWmlMarkupFromCorrelatedSequence(
            OpenXmlPart part,
            IEnumerable<ComparisonUnitAtom> comparisonUnitAtomList,
            WmlComparerSettings settings)
        {
            // fabricate new MainDocumentPart from correlatedSequence
            _maxId = 0;
            object newBodyChildren = CoalesceRecurse(part, comparisonUnitAtomList, 0, settings);
            return newBodyChildren;
        }

        private static void MarkContentAsDeletedOrInserted(XDocument newXDoc, WmlComparerSettings settings)
        {
            object newRoot = MarkContentAsDeletedOrInsertedTransform(newXDoc.Root, settings);
            newXDoc.Root?.ReplaceWith(newRoot);
        }

        private static object MarkContentAsDeletedOrInsertedTransform(XNode node, WmlComparerSettings settings)
        {
            if (node is XElement element)
            {
                if (element.Name == W.r)
                {
                    List<string> statusList = element
                        .DescendantsTrimmed(W.txbxContent)
                        .Where(d => d.Name == W.t || d.Name == W.delText || AllowableRunChildren.Contains(d.Name))
                        .Attributes(PtOpenXml.Status)
                        .Select(a => (string) a)
                        .Distinct()
                        .ToList();

                    if (statusList.Count() > 1)
                    {
                        throw new OpenXmlPowerToolsException(
                            "Internal error - have both deleted and inserted text elements in the same run.");
                    }

                    if (statusList.Count == 0)
                    {
                        return new XElement(W.r,
                            element.Attributes(),
                            element.Nodes().Select(n => MarkContentAsDeletedOrInsertedTransform(n, settings)));
                    }

                    if (statusList.First() == "Deleted")
                    {
                        return new XElement(W.del,
                            new XAttribute(W.author, settings.AuthorForRevisions),
                            new XAttribute(W.id, _maxId++),
                            new XAttribute(W.date, settings.DateTimeForRevisions),
                            new XElement(W.r,
                                element.Attributes(),
                                element.Nodes().Select(n => MarkContentAsDeletedOrInsertedTransform(n, settings))));
                    }

                    if (statusList.First() == "Inserted")
                    {
                        return new XElement(W.ins,
                            new XAttribute(W.author, settings.AuthorForRevisions),
                            new XAttribute(W.id, _maxId++),
                            new XAttribute(W.date, settings.DateTimeForRevisions),
                            new XElement(W.r,
                                element.Attributes(),
                                element.Nodes().Select(n => MarkContentAsDeletedOrInsertedTransform(n, settings))));
                    }
                }

                if (element.Name == W.pPr)
                {
                    var status = (string) element.Attribute(PtOpenXml.Status);
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
                            new XAttribute(W.id, _maxId++),
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
                            new XAttribute(W.id, _maxId++),
                            new XAttribute(W.date, settings.DateTimeForRevisions)));
                        if (pPr.Element(W.rPr) != null)
                            pPr.Element(W.rPr).ReplaceWith(rPr);
                        else
                            pPr.AddFirst(rPr);
                    }
                    else
                    {
                        throw new OpenXmlPowerToolsException("Internal error");
                    }

                    return pPr;
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => MarkContentAsDeletedOrInsertedTransform(n, settings)));
            }

            return node;
        }

        private static void CoalesceAdjacentRunsWithIdenticalFormatting(XDocument xDoc)
        {
            IEnumerable<XElement> paras = xDoc.Root.DescendantsTrimmed(W.txbxContent).Where(d => d.Name == W.p);
            foreach (XElement para in paras)
            {
                XElement newPara = WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(para);
                para.ReplaceNodes(newPara.Nodes());
            }
        }

        private static void IgnorePt14Namespace(XElement root)
        {
            if (root.Attribute(XNamespace.Xmlns + "pt14") == null)
            {
                root.Add(new XAttribute(XNamespace.Xmlns + "pt14", PtOpenXml.pt.NamespaceName));
            }

            var ignorable = (string) root.Attribute(MC.Ignorable);
            if (ignorable != null)
            {
                string[] list = ignorable.Split(' ');
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

        private static void ProcessFootnoteEndnote(
            WmlComparerSettings settings,
            List<ComparisonUnitAtom> listOfComparisonUnitAtoms,
            MainDocumentPart mainDocumentPartBefore,
            MainDocumentPart mainDocumentPartAfter,
            XDocument mainDocumentXDoc)
        {
            FootnotesPart footnotesPartBefore = mainDocumentPartBefore.FootnotesPart;
            EndnotesPart endnotesPartBefore = mainDocumentPartBefore.EndnotesPart;
            FootnotesPart footnotesPartAfter = mainDocumentPartAfter.FootnotesPart;
            EndnotesPart endnotesPartAfter = mainDocumentPartAfter.EndnotesPart;

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

            List<ComparisonUnitAtom> possiblyModifiedFootnotesEndNotes = listOfComparisonUnitAtoms
                .Where(cua =>
                    cua.ContentElement.Name == W.footnoteReference ||
                    cua.ContentElement.Name == W.endnoteReference)
                .ToList();

            foreach (ComparisonUnitAtom fn in possiblyModifiedFootnotesEndNotes)
            {
                string beforeId = null;
                if (fn.ContentElementBefore != null)
                    beforeId = (string) fn.ContentElementBefore.Attribute(W.id);
                var afterId = (string) fn.ContentElement.Attribute(W.id);

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
                            .FirstOrDefault(fnn => (string) fnn.Attribute(W.id) == beforeId);
                        footnoteEndnoteAfter = footnotesPartAfterXDoc
                            .Root
                            .Elements()
                            .FirstOrDefault(fnn => (string) fnn.Attribute(W.id) == afterId);
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
                            .FirstOrDefault(fnn => (string) fnn.Attribute(W.id) == beforeId);
                        footnoteEndnoteAfter = endnotesPartAfterXDoc
                            .Root
                            .Elements()
                            .FirstOrDefault(fnn => (string) fnn.Attribute(W.id) == afterId);
                        partToUseBefore = endnotesPartBefore;
                        partToUseAfter = endnotesPartAfter;
                        partToUseBeforeXDoc = endnotesPartBeforeXDoc;
                        partToUseAfterXDoc = endnotesPartAfterXDoc;
                    }

                    AddSha1HashToBlockLevelContent(partToUseBefore, footnoteEndnoteBefore, settings);
                    AddSha1HashToBlockLevelContent(partToUseAfter, footnoteEndnoteAfter, settings);

                    ComparisonUnitAtom[] fncal1 = CreateComparisonUnitAtomList(partToUseBefore, footnoteEndnoteBefore, settings);
                    ComparisonUnit[] fncus1 = GetComparisonUnitList(fncal1, settings);

                    ComparisonUnitAtom[] fncal2 = CreateComparisonUnitAtomList(partToUseAfter, footnoteEndnoteAfter, settings);
                    ComparisonUnit[] fncus2 = GetComparisonUnitList(fncal2, settings);

                    if (!(fncus1.Length == 0 && fncus2.Length == 0))
                    {
                        List<CorrelatedSequence> fnCorrelatedSequence = Lcs(fncus1, fncus2, settings);

                        if (False)
                        {
                            var sb = new StringBuilder();
                            foreach (CorrelatedSequence item in fnCorrelatedSequence)
                                sb.Append(item).Append(Environment.NewLine);
                            string sbs = sb.ToString();
                            TestUtil.NotePad(sbs);
                        }

                        // for any deleted or inserted rows, we go into the w:trPr properties, and add the appropriate w:ins or w:del element, and therefore
                        // when generating the document, the appropriate row will be marked as deleted or inserted.
                        MarkRowsAsDeletedOrInserted(settings, fnCorrelatedSequence);

                        // the following gets a flattened list of ComparisonUnitAtoms, with status indicated in each ComparisonUnitAtom: Deleted, Inserted, or Equal
                        List<ComparisonUnitAtom> fnListOfComparisonUnitAtoms =
                            FlattenToComparisonUnitAtomList(fnCorrelatedSequence, settings);

                        if (False)
                        {
                            var sb = new StringBuilder();
                            foreach (ComparisonUnitAtom item in fnListOfComparisonUnitAtoms)
                                sb.Append(item + Environment.NewLine);
                            string sbs = sb.ToString();
                            TestUtil.NotePad(sbs);
                        }

                        // hack = set the guid ID of the table, row, or cell from the 'before' document to be equal to the 'after' document.

                        // note - we don't want to do the hack until after flattening all of the groups.  At the end of the flattening, we should simply
                        // have a list of ComparisonUnitAtoms, appropriately marked as equal, inserted, or deleted.

                        // the table id will be hacked in the normal course of events.
                        // in the case where a row is deleted, not necessary to hack - the deleted row ID will do.
                        // in the case where a row is inserted, not necessary to hack - the inserted row ID will do as well.
                        AssembleAncestorUnidsInOrderToRebuildXmlTreeProperly(fnListOfComparisonUnitAtoms);

                        object newFootnoteEndnoteChildren =
                            ProduceNewWmlMarkupFromCorrelatedSequence(partToUseAfter, fnListOfComparisonUnitAtoms, settings);
                        var tempElement = new XElement(W.body, newFootnoteEndnoteChildren);
                        bool hasFootnoteReference = tempElement.Descendants(W.r).Any(r =>
                        {
                            var b = false;
                            if ((string) r.Elements(W.rPr).Elements(W.rStyle).Attributes(W.val).FirstOrDefault() ==
                                "FootnoteReference")
                                b = true;
                            if (r.Descendants(W.footnoteRef).Any())
                                b = true;
                            return b;
                        });
                        if (!hasFootnoteReference)
                        {
                            XElement firstPara = tempElement.Descendants(W.p).FirstOrDefault();
                            if (firstPara != null)
                            {
                                XElement firstRun = firstPara.Element(W.r);
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

                        var newTempElement = (XElement) WordprocessingMLUtil.WmlOrderElementsPerStandard(tempElement);
                        XElement newContentElement = newTempElement.Descendants()
                            .FirstOrDefault(d => d.Name == W.footnote || d.Name == W.endnote);
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
                            .FirstOrDefault(fnn => (string) fnn.Attribute(W.id) == afterId);
                        partToUseAfter = footnotesPartAfter;
                        partToUseAfterXDoc = footnotesPartAfterXDoc;
                    }
                    else
                    {
                        footnoteEndnoteAfter = endnotesPartAfterXDoc
                            .Root
                            .Elements()
                            .FirstOrDefault(fnn => (string) fnn.Attribute(W.id) == afterId);
                        partToUseAfter = endnotesPartAfter;
                        partToUseAfterXDoc = endnotesPartAfterXDoc;
                    }

                    AddSha1HashToBlockLevelContent(partToUseAfter, footnoteEndnoteAfter, settings);

                    ComparisonUnitAtom[] fncal2 = CreateComparisonUnitAtomList(partToUseAfter, footnoteEndnoteAfter, settings);
                    ComparisonUnit[] fncus2 = GetComparisonUnitList(fncal2, settings);

                    var insertedCorrSequ = new List<CorrelatedSequence>
                    {
                        new CorrelatedSequence
                        {
                            ComparisonUnitArray1 = null,
                            ComparisonUnitArray2 = fncus2,
                            CorrelationStatus = CorrelationStatus.Inserted
                        }
                    };

                    if (False)
                    {
                        var sb = new StringBuilder();
                        foreach (CorrelatedSequence item in insertedCorrSequ)
                            sb.Append(item).Append(Environment.NewLine);
                        string sbs = sb.ToString();
                        TestUtil.NotePad(sbs);
                    }

                    MarkRowsAsDeletedOrInserted(settings, insertedCorrSequ);

                    List<ComparisonUnitAtom> fnListOfComparisonUnitAtoms =
                        FlattenToComparisonUnitAtomList(insertedCorrSequ, settings);

                    AssembleAncestorUnidsInOrderToRebuildXmlTreeProperly(fnListOfComparisonUnitAtoms);

                    object newFootnoteEndnoteChildren = ProduceNewWmlMarkupFromCorrelatedSequence(partToUseAfter,
                        fnListOfComparisonUnitAtoms, settings);
                    var tempElement = new XElement(W.body, newFootnoteEndnoteChildren);
                    bool hasFootnoteReference = tempElement.Descendants(W.r).Any(r =>
                    {
                        var b = false;
                        if ((string) r.Elements(W.rPr).Elements(W.rStyle).Attributes(W.val).FirstOrDefault() ==
                            "FootnoteReference")
                            b = true;
                        if (r.Descendants(W.footnoteRef).Any())
                            b = true;
                        return b;
                    });
                    if (!hasFootnoteReference)
                    {
                        XElement firstPara = tempElement.Descendants(W.p).FirstOrDefault();
                        if (firstPara != null)
                        {
                            XElement firstRun = firstPara.Descendants(W.r).FirstOrDefault();
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

                    var newTempElement = (XElement) WordprocessingMLUtil.WmlOrderElementsPerStandard(tempElement);
                    XElement newContentElement = newTempElement
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
                            .FirstOrDefault(fnn => (string) fnn.Attribute(W.id) == afterId);
                        partToUseAfter = footnotesPartAfter;
                        partToUseAfterXDoc = footnotesPartAfterXDoc;
                    }
                    else
                    {
                        footnoteEndnoteBefore = endnotesPartBeforeXDoc
                            .Root
                            .Elements()
                            .FirstOrDefault(fnn => (string) fnn.Attribute(W.id) == afterId);
                        partToUseBefore = endnotesPartBefore;
                        partToUseBeforeXDoc = endnotesPartBeforeXDoc;
                    }

                    AddSha1HashToBlockLevelContent(partToUseBefore, footnoteEndnoteBefore, settings);

                    ComparisonUnitAtom[] fncal2 = CreateComparisonUnitAtomList(partToUseBefore, footnoteEndnoteBefore, settings);
                    ComparisonUnit[] fncus2 = GetComparisonUnitList(fncal2, settings);

                    var deletedCorrSequ = new List<CorrelatedSequence>
                    {
                        new CorrelatedSequence
                        {
                            ComparisonUnitArray1 = fncus2,
                            ComparisonUnitArray2 = null,
                            CorrelationStatus = CorrelationStatus.Deleted
                        }
                    };

                    if (False)
                    {
                        var sb = new StringBuilder();
                        foreach (CorrelatedSequence item in deletedCorrSequ)
                            sb.Append(item).Append(Environment.NewLine);
                        string sbs = sb.ToString();
                        TestUtil.NotePad(sbs);
                    }

                    MarkRowsAsDeletedOrInserted(settings, deletedCorrSequ);

                    List<ComparisonUnitAtom> fnListOfComparisonUnitAtoms =
                        FlattenToComparisonUnitAtomList(deletedCorrSequ, settings);

                    if (fnListOfComparisonUnitAtoms.Any())
                    {
                        AssembleAncestorUnidsInOrderToRebuildXmlTreeProperly(fnListOfComparisonUnitAtoms);

                        object newFootnoteEndnoteChildren = ProduceNewWmlMarkupFromCorrelatedSequence(partToUseBefore,
                            fnListOfComparisonUnitAtoms, settings);
                        var tempElement = new XElement(W.body, newFootnoteEndnoteChildren);
                        bool hasFootnoteReference = tempElement.Descendants(W.r).Any(r =>
                        {
                            var b = false;
                            if ((string) r.Elements(W.rPr).Elements(W.rStyle).Attributes(W.val).FirstOrDefault() ==
                                "FootnoteReference")
                                b = true;
                            if (r.Descendants(W.footnoteRef).Any())
                                b = true;
                            return b;
                        });
                        if (!hasFootnoteReference)
                        {
                            XElement firstPara = tempElement.Descendants(W.p).FirstOrDefault();
                            if (firstPara != null)
                            {
                                XElement firstRun = firstPara.Descendants(W.r).FirstOrDefault();
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

                        var newTempElement = (XElement) WordprocessingMLUtil.WmlOrderElementsPerStandard(tempElement);
                        XElement newContentElement = newTempElement.Descendants()
                            .FirstOrDefault(d => d.Name == W.footnote || d.Name == W.endnote);
                        if (newContentElement == null)
                            throw new OpenXmlPowerToolsException("Internal error");

                        footnoteEndnoteBefore.ReplaceNodes(newContentElement.Nodes());
                    }
                }
                else
                {
                    throw new OpenXmlPowerToolsException("Internal error");
                }
            }
        }

        private static void RectifyFootnoteEndnoteIds(
            MainDocumentPart mainDocumentPartBefore,
            MainDocumentPart mainDocumentPartAfter,
            MainDocumentPart mainDocumentPartWithRevisions,
            XDocument mainDocumentXDoc,
            WmlComparerSettings settings)
        {
            FootnotesPart footnotesPartBefore = mainDocumentPartBefore.FootnotesPart;
            EndnotesPart endnotesPartBefore = mainDocumentPartBefore.EndnotesPart;
            FootnotesPart footnotesPartAfter = mainDocumentPartAfter.FootnotesPart;
            EndnotesPart endnotesPartAfter = mainDocumentPartAfter.EndnotesPart;
            FootnotesPart footnotesPartWithRevisions = mainDocumentPartWithRevisions.FootnotesPart;
            EndnotesPart endnotesPartWithRevisions = mainDocumentPartWithRevisions.EndnotesPart;

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
                    .Where(e => (string) e.Attribute(W.id) != "-1" && (string) e.Attribute(W.id) != "0")
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
                    .Where(e => (string) e.Attribute(W.id) != "-1" && (string) e.Attribute(W.id) != "0")
                    .Remove();
            }

            var footnotesRefs = mainDocumentXDoc
                .Descendants(W.footnoteReference)
                .Select((fn, idx) =>
                {
                    return new
                    {
                        FootNote = fn,
                        Idx = idx
                    };
                });

            foreach (var fn in footnotesRefs)
            {
                var oldId = (string) fn.FootNote.Attribute(W.id);
                string newId = (fn.Idx + 1).ToString();
                fn.FootNote.Attribute(W.id).Value = newId;
                XElement footnote = footnotesPartAfterXDoc
                    .Root
                    .Elements()
                    .FirstOrDefault(e => (string) e.Attribute(W.id) == oldId);
                if (footnote == null)
                {
                    footnote = footnotesPartBeforeXDoc
                        .Root
                        .Elements()
                        .FirstOrDefault(e => (string) e.Attribute(W.id) == oldId);
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
                        Idx = idx
                    };
                });

            foreach (var fn in endnotesRefs)
            {
                var oldId = (string) fn.Endnote.Attribute(W.id);
                string newId = (fn.Idx + 1).ToString();
                fn.Endnote.Attribute(W.id).Value = newId;
                XElement endnote = endnotesPartAfterXDoc
                    .Root
                    .Elements()
                    .FirstOrDefault(e => (string) e.Attribute(W.id) == oldId);
                if (endnote == null)
                {
                    endnote = endnotesPartBeforeXDoc
                        .Root
                        .Elements()
                        .FirstOrDefault(e => (string) e.Attribute(W.id) == oldId);
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
                var newXDocRoot =
                    (XElement) WordprocessingMLUtil.WmlOrderElementsPerStandard(footnotesPartWithRevisionsXDoc.Root);
                footnotesPartWithRevisionsXDoc.Root.ReplaceWith(newXDocRoot);
                IgnorePt14Namespace(footnotesPartWithRevisionsXDoc.Root);
                footnotesPartWithRevisions.PutXDocument();
            }

            if (endnotesPartWithRevisionsXDoc != null)
            {
                MarkContentAsDeletedOrInserted(endnotesPartWithRevisionsXDoc, settings);
                CoalesceAdjacentRunsWithIdenticalFormatting(endnotesPartWithRevisionsXDoc);
                var newXDocRoot = (XElement) WordprocessingMLUtil.WmlOrderElementsPerStandard(endnotesPartWithRevisionsXDoc.Root);
                endnotesPartWithRevisionsXDoc.Root.ReplaceWith(newXDocRoot);
                IgnorePt14Namespace(endnotesPartWithRevisionsXDoc.Root);
                endnotesPartWithRevisions.PutXDocument();
            }
        }

        private static void ConjoinDeletedInsertedParagraphMarks(MainDocumentPart mainDocumentPart, XDocument newXDoc)
        {
            ConjoinMultipleParagraphMarks(newXDoc);
            if (mainDocumentPart.FootnotesPart != null)
            {
                XDocument fnXDoc = mainDocumentPart.FootnotesPart.GetXDocument();
                ConjoinMultipleParagraphMarks(fnXDoc);
                mainDocumentPart.FootnotesPart.PutXDocument();
            }

            if (mainDocumentPart.EndnotesPart != null)
            {
                XDocument fnXDoc = mainDocumentPart.EndnotesPart.GetXDocument();
                ConjoinMultipleParagraphMarks(fnXDoc);
                mainDocumentPart.EndnotesPart.PutXDocument();
            }
        }

        // it is possible, per the algorithm, for the algorithm to find that the paragraph mark for a single paragraph has been
        // inserted and deleted.  If the algorithm sets them to equal, then sometimes it will equate paragraph marks that should
        // not be equated.

        private static void ConjoinMultipleParagraphMarks(XDocument xDoc)
        {
            object newRoot = ConjoinTransform(xDoc.Root);
            xDoc.Root?.ReplaceWith(newRoot);
        }

        private static object ConjoinTransform(XNode node)
        {
            if (node is XElement element)
            {
                if (element.Name == W.p && element.Elements(W.pPr).Count() >= 2)
                {
                    var pPr = new XElement(element.Elements(W.pPr).First());
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
                    element.Nodes().Select(ConjoinTransform));
            }

            return node;
        }

        private static void FixUpRevisionIds(WordprocessingDocument wDocWithRevisions, XDocument newXDoc)
        {
            IEnumerable<XElement> footnoteRevisions = Enumerable.Empty<XElement>();
            if (wDocWithRevisions.MainDocumentPart.FootnotesPart != null)
            {
                XDocument fnxd = wDocWithRevisions.MainDocumentPart.FootnotesPart.GetXDocument();
                footnoteRevisions = fnxd
                    .Descendants()
                    .Where(d => d.Name == W.ins || d.Name == W.del);
            }

            IEnumerable<XElement> endnoteRevisions = Enumerable.Empty<XElement>();
            if (wDocWithRevisions.MainDocumentPart.EndnotesPart != null)
            {
                XDocument fnxd = wDocWithRevisions.MainDocumentPart.EndnotesPart.GetXDocument();
                endnoteRevisions = fnxd
                    .Descendants()
                    .Where(d => d.Name == W.ins || d.Name == W.del);
            }

            IEnumerable<XElement> mainRevisions = newXDoc
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
                        Idx = i + 1
                    };
                });
            foreach (var item in allRevisions)
                item.Rev.Attribute(W.id).Value = item.Idx.ToString();
            if (wDocWithRevisions.MainDocumentPart.FootnotesPart != null)
                wDocWithRevisions.MainDocumentPart.FootnotesPart.PutXDocument();
            if (wDocWithRevisions.MainDocumentPart.EndnotesPart != null)
                wDocWithRevisions.MainDocumentPart.EndnotesPart.PutXDocument();
        }

        private static void MoveLastSectPrToChildOfBody(XDocument newXDoc)
        {
            XElement lastParaWithSectPr = newXDoc
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

                private static void FixUpFootnotesEndnotesWithCustomMarkers(WordprocessingDocument wDocWithRevisions)
        {
#if FALSE

// this needs to change
      <w:del w:author = "Open-Xml-PowerTools"
             w:id = "7"
             w:date = "2017-06-07T12:23:22.8601285-07:00">
        <w:r>
          <w:rPr pt14:Unid = "ec75a71361c84562a757eee8b28fc229">
            <w:rFonts w:cs = "Times New Roman Bold"
                      pt14:Unid = "16bb355df5964ba09854f9152c97242b" />
            <w:b w:val = "0"
                 pt14:Unid = "9abcec54ad414791a5627cbb198e8aa9" />
            <w:bCs pt14:Unid = "71ecd2eba85e4bfaa92b3d618e2f8829" />
            <w:position w:val = "6"
                        pt14:Unid = "61793f6a5f494700b7f2a3a753ce9055" />
            <w:sz w:val = "16"
                  pt14:Unid = "60b3cd020c214d0ea07e5a68ae0e4efe" />
            <w:szCs w:val = "16"
                    pt14:Unid = "9ae61a724de44a75868180aac44ea380" />
          </w:rPr>
          <w:footnoteReference w:customMarkFollows = "1"
                               w:id = "1"
                               pt14:Status = "Deleted" />
        </w:r>
      </w:del>
      <w:del w:author = "Open-Xml-PowerTools"
             w:id = "8"
             w:date = "2017-06-07T12:23:22.8601285-07:00">
        <w:r>
          <w:rPr pt14:Unid = "445caef74a624e588e7adaa6d7775639">
            <w:rFonts w:cs = "Times New Roman Bold"
                      pt14:Unid = "5920885f8ec44c53bcaece2de7eafda2" />
            <w:b w:val = "0"
                 pt14:Unid = "023a29e2e6d44c3b8c5df47317ace4c6" />
            <w:bCs pt14:Unid = "e96e37daf9174b268ef4731df831df7d" />
            <w:position w:val = "6"
                        pt14:Unid = "be3f8ff7ed0745ae9340bb2706b28b1f" />
            <w:sz w:val = "16"
                  pt14:Unid = "6fbbde024e7c46b9b72435ae50065459" />
            <w:szCs w:val = "16"
                    pt14:Unid = "cc82e7bd75f441f2b609eae0672fb285" />
          </w:rPr>
          <w:delText>1</w:delText>
        </w:r>
      </w:del>

      // to this
      <w:del w:author = "Open-Xml-PowerTools"
             w:id = "7"
             w:date = "2017-06-07T12:23:22.8601285-07:00">
        <w:r>
          <w:rPr pt14:Unid = "ec75a71361c84562a757eee8b28fc229">
            <w:rFonts w:cs = "Times New Roman Bold"
                      pt14:Unid = "16bb355df5964ba09854f9152c97242b" />
            <w:b w:val = "0"
                 pt14:Unid = "9abcec54ad414791a5627cbb198e8aa9" />
            <w:bCs pt14:Unid = "71ecd2eba85e4bfaa92b3d618e2f8829" />
            <w:position w:val = "6"
                        pt14:Unid = "61793f6a5f494700b7f2a3a753ce9055" />
            <w:sz w:val = "16"
                  pt14:Unid = "60b3cd020c214d0ea07e5a68ae0e4efe" />
            <w:szCs w:val = "16"
                    pt14:Unid = "9ae61a724de44a75868180aac44ea380" />
          </w:rPr>
          <w:footnoteReference w:customMarkFollows = "1"
                               w:id = "1"
                               pt14:Status = "Deleted" />
          <w:delText>1</w:delText>
        </w:r>
      </w:del>
#endif

            // this is pretty random - a bug in Word prevents display of a document if the delText element does not immediately follow the footnoteReference element, in the same run.
            XDocument mainXDoc = wDocWithRevisions.MainDocumentPart.GetXDocument();
            var newRoot = (XElement) FootnoteEndnoteReferenceCleanupTransform(mainXDoc.Root);
            mainXDoc.Root?.ReplaceWith(newRoot);
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
                    bool hasFootnoteEndnoteReferencesThatNeedCleanedUp = element
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
                        IEnumerable<XElement> footnoteEndnoteReferencesToAdjust = clone
                            .Descendants()
                            .Where(d => d.Name == W.footnoteReference || d.Name == W.endnoteReference)
                            .Where(d => d.Attribute(W.customMarkFollows) != null);
                        foreach (XElement fnenr in footnoteEndnoteReferencesToAdjust)
                        {
                            XElement par = fnenr.Parent;
                            XElement gp = fnenr.Parent.Parent;
                            if (par.Name == W.r &&
                                gp.Name == W.del)
                            {
                                if (par.Element(W.delText) != null)
                                    continue;

                                XElement afterGp = gp.ElementsAfterSelf().FirstOrDefault();
                                if (afterGp == null)
                                    continue;

                                IEnumerable<XElement> afterGpDelText = afterGp.Elements(W.r).Elements(W.delText);
                                if (afterGpDelText.Any())
                                {
                                    par.Add(afterGpDelText); // this will clone and add to run that contains the reference
                                    afterGpDelText.Remove(); // this leaves an empty run, does not matter.
                                }
                            }

                            if (par.Name == W.r &&
                                gp.Name == W.ins)
                            {
                                if (par.Element(W.t) != null)
                                    continue;

                                XElement afterGp = gp.ElementsAfterSelf().FirstOrDefault();
                                if (afterGp == null)
                                    continue;

                                IEnumerable<XElement> afterGpText = afterGp.Elements(W.r).Elements(W.t);
                                if (afterGpText.Any())
                                {
                                    par.Add(afterGpText); // this will clone and add to run that contains the reference
                                    afterGpText.Remove(); // this leaves an empty run, does not matter.
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

        private static void FixUpRevMarkIds(WordprocessingDocument wDoc)
        {
            IEnumerable<XElement> revMarksToChange = wDoc
                .ContentParts()
                .Select(cp => cp.GetXDocument())
                .Select(xd => xd.Descendants().Where(d => d.Name == W.ins || d.Name == W.del))
                .SelectMany(m => m);
            var nextId = 0;
            foreach (XElement item in revMarksToChange)
            {
                XAttribute idAtt = item.Attribute(W.id);
                if (idAtt != null)
                    idAtt.Value = nextId++.ToString();
            }

            foreach (OpenXmlPart cp in wDoc.ContentParts())
                cp.PutXDocument();
        }

        private static void FixUpDocPrIds(WordprocessingDocument wDoc)
        {
            XName elementToFind = WP.docPr;
            IEnumerable<XElement> docPrToChange = wDoc
                .ContentParts()
                .Select(cp => cp.GetXDocument())
                .Select(xd => xd.Descendants().Where(d => d.Name == elementToFind))
                .SelectMany(m => m);
            var nextId = 1;
            foreach (XElement item in docPrToChange)
            {
                XAttribute idAtt = item.Attribute("id");
                if (idAtt != null)
                    idAtt.Value = nextId++.ToString();
            }

            foreach (OpenXmlPart cp in wDoc.ContentParts())
                cp.PutXDocument();
        }

        private static void FixUpShapeIds(WordprocessingDocument wDoc)
        {
            XName elementToFind = VML.shape;
            IEnumerable<XElement> shapeIdsToChange = wDoc
                .ContentParts()
                .Select(cp => cp.GetXDocument())
                .Select(xd => xd.Descendants().Where(d => d.Name == elementToFind))
                .SelectMany(m => m);
            var nextId = 1;
            foreach (XElement item in shapeIdsToChange)
            {
                int thisId = nextId++;

                XAttribute idAtt = item.Attribute("id");
                if (idAtt != null)
                    idAtt.Value = thisId.ToString();

                XElement oleObject = item.Parent.Element(O.OLEObject);
                if (oleObject != null)
                {
                    XAttribute shapeIdAtt = oleObject.Attribute("ShapeID");
                    if (shapeIdAtt != null)
                        shapeIdAtt.Value = thisId.ToString();
                }
            }

            foreach (OpenXmlPart cp in wDoc.ContentParts())
                cp.PutXDocument();
        }

        private static void FixUpShapeTypeIds(WordprocessingDocument wDoc)
        {
            XName elementToFind = VML.shapetype;
            IEnumerable<XElement> shapeTypeIdsToChange = wDoc
                .ContentParts()
                .Select(cp => cp.GetXDocument())
                .Select(xd => xd.Descendants().Where(d => d.Name == elementToFind))
                .SelectMany(m => m);
            var nextId = 1;
            foreach (XElement item in shapeTypeIdsToChange)
            {
                int thisId = nextId++;

                XAttribute idAtt = item.Attribute("id");
                if (idAtt != null)
                    idAtt.Value = thisId.ToString();

                XElement shape = item.Parent.Element(VML.shape);
                if (shape != null)
                {
                    XAttribute typeAtt = shape.Attribute("type");
                    if (typeAtt != null)
                        typeAtt.Value = thisId.ToString();
                }
            }

            foreach (OpenXmlPart cp in wDoc.ContentParts())
                cp.PutXDocument();
        }

        private static void AddFootnotesEndnotesStyles(WordprocessingDocument wDocWithRevisions)
        {
            XDocument mainXDoc = wDocWithRevisions.MainDocumentPart.GetXDocument();
            bool hasFootnotes = mainXDoc.Descendants(W.footnoteReference).Any();
            bool hasEndnotes = mainXDoc.Descendants(W.endnoteReference).Any();
            StyleDefinitionsPart styleDefinitionsPart = wDocWithRevisions.MainDocumentPart.StyleDefinitionsPart;
            XDocument sXDoc = styleDefinitionsPart.GetXDocument();
            if (hasFootnotes)
            {
                XElement footnoteTextStyle = sXDoc
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(s => (string) s.Attribute(W.styleId) == "FootnoteText");
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
                    XElement ftsElement = XElement.Parse(footnoteTextStyleMarkup);
                    sXDoc.Root.Add(ftsElement);
                }

                XElement footnoteTextCharStyle = sXDoc
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(s => (string) s.Attribute(W.styleId) == "FootnoteTextChar");
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
                    XElement fntcsElement = XElement.Parse(footnoteTextCharStyleMarkup);
                    sXDoc.Root.Add(fntcsElement);
                }

                XElement footnoteReferenceStyle = sXDoc
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(s => (string) s.Attribute(W.styleId) == "FootnoteReference");
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
                    XElement fnrsElement = XElement.Parse(footnoteReferenceStyleMarkup);
                    sXDoc.Root.Add(fnrsElement);
                }
            }

            if (hasEndnotes)
            {
                XElement endnoteTextStyle = sXDoc
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(s => (string) s.Attribute(W.styleId) == "EndnoteText");
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
                    XElement etsElement = XElement.Parse(endnoteTextStyleMarkup);
                    sXDoc.Root.Add(etsElement);
                }

                XElement endnoteTextCharStyle = sXDoc
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(s => (string) s.Attribute(W.styleId) == "EndnoteTextChar");
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
                    XElement entcsElement = XElement.Parse(endnoteTextCharStyleMarkup);
                    sXDoc.Root.Add(entcsElement);
                }

                XElement endnoteReferenceStyle = sXDoc
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(s => (string) s.Attribute(W.styleId) == "EndnoteReference");
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
                    XElement enrsElement = XElement.Parse(endnoteReferenceStyleMarkup);
                    sXDoc.Root.Add(enrsElement);
                }
            }

            if (hasFootnotes || hasEndnotes)
            {
                styleDefinitionsPart.PutXDocument();
            }
        }

        private static void CopyMissingStylesFromOneDocToAnother(WordprocessingDocument wDocFrom, WordprocessingDocument wDocTo)
        {
            XDocument revisionsStylesXDoc = wDocTo.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
            XDocument afterStylesXDoc = wDocFrom.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
            foreach (XElement style in afterStylesXDoc.Root.Elements(W.style))
            {
                var type = (string) style.Attribute(W.type);
                var styleId = (string) style.Attribute(W.styleId);
                XElement styleInRevDoc = revisionsStylesXDoc
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(st => (string) st.Attribute(W.type) == type &&
                                          (string) st.Attribute(W.styleId) == styleId);
                if (styleInRevDoc != null)
                    continue;

                var cloned = new XElement(style);
                if (cloned.Attribute(W._default) != null)
                    cloned.Attribute(W._default).Remove();
                revisionsStylesXDoc.Root.Add(cloned);
            }

            wDocTo.MainDocumentPart.StyleDefinitionsPart.PutXDocument();
        }

        private static void DeleteFootnotePropertiesInSettings(WordprocessingDocument wDocWithRevisions)
        {
            DocumentSettingsPart settingsPart = wDocWithRevisions.MainDocumentPart.DocumentSettingsPart;
            if (settingsPart != null)
            {
                XDocument sxDoc = settingsPart.GetXDocument();
                sxDoc.Root?.Elements().Where(e => e.Name == W.footnotePr || e.Name == W.endnotePr).Remove();
                settingsPart.PutXDocument();
            }
        }

        private static object CloneForStructureHash(XNode node)
        {
            if (node is XElement element)
            {
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Elements().Select(CloneForStructureHash));
            }

            return null;
        }

        private static List<CorrelatedSequence> FindCommonAtBeginningAndEnd(
            CorrelatedSequence unknown,
            WmlComparerSettings settings)
        {
            int lengthToCompare = Math.Min(unknown.ComparisonUnitArray1.Length, unknown.ComparisonUnitArray2.Length);

            int countCommonAtBeginning = unknown
                .ComparisonUnitArray1
                .Take(lengthToCompare)
                .Zip(unknown.ComparisonUnitArray2,
                    (pu1, pu2) => new
                    {
                        Pu1 = pu1,
                        Pu2 = pu2
                    })
                .TakeWhile(pair => pair.Pu1.SHA1Hash == pair.Pu2.SHA1Hash)
                .Count();

            if (countCommonAtBeginning != 0 && countCommonAtBeginning / (double) lengthToCompare < settings.DetailThreshold)
                countCommonAtBeginning = 0;

            if (countCommonAtBeginning != 0)
            {
                var newSequence = new List<CorrelatedSequence>();

                var csEqual = new CorrelatedSequence
                {
                    CorrelationStatus = CorrelationStatus.Equal,
                    ComparisonUnitArray1 = unknown
                        .ComparisonUnitArray1
                        .Take(countCommonAtBeginning)
                        .ToArray(),
                    ComparisonUnitArray2 = unknown
                        .ComparisonUnitArray2
                        .Take(countCommonAtBeginning)
                        .ToArray()
                };
                newSequence.Add(csEqual);

                int remainingLeft = unknown.ComparisonUnitArray1.Length - countCommonAtBeginning;
                int remainingRight = unknown.ComparisonUnitArray2.Length - countCommonAtBeginning;

                if (remainingLeft != 0 && remainingRight == 0)
                {
                    var csDeleted = new CorrelatedSequence
                    {
                        CorrelationStatus = CorrelationStatus.Deleted,
                        ComparisonUnitArray1 = unknown.ComparisonUnitArray1.Skip(countCommonAtBeginning).ToArray(),
                        ComparisonUnitArray2 = null
                    };
                    newSequence.Add(csDeleted);
                }
                else if (remainingLeft == 0 && remainingRight != 0)
                {
                    var csInserted = new CorrelatedSequence
                    {
                        CorrelationStatus = CorrelationStatus.Inserted,
                        ComparisonUnitArray1 = null,
                        ComparisonUnitArray2 = unknown.ComparisonUnitArray2.Skip(countCommonAtBeginning).ToArray()
                    };
                    newSequence.Add(csInserted);
                }
                else if (remainingLeft != 0 && remainingRight != 0)
                {
                    if (unknown.ComparisonUnitArray1[0] is ComparisonUnitWord first1 &&
                        unknown.ComparisonUnitArray2[0] is ComparisonUnitWord first2)
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

                        ComparisonUnit[] remainingInLeft = unknown
                            .ComparisonUnitArray1
                            .Skip(countCommonAtBeginning)
                            .ToArray();

                        ComparisonUnit[] remainingInRight = unknown
                            .ComparisonUnitArray2
                            .Skip(countCommonAtBeginning)
                            .ToArray();

                        ComparisonUnitAtom lastContentAtomLeft = unknown.ComparisonUnitArray1[countCommonAtBeginning - 1]
                            .DescendantContentAtoms()
                            .FirstOrDefault();

                        ComparisonUnitAtom lastContentAtomRight = unknown.ComparisonUnitArray2[countCommonAtBeginning - 1]
                            .DescendantContentAtoms()
                            .FirstOrDefault();

                        if (lastContentAtomLeft?.ContentElement.Name != W.pPr && lastContentAtomRight?.ContentElement.Name != W.pPr)
                        {
                            List<ComparisonUnit[]> split1 = SplitAtParagraphMark(remainingInLeft);
                            List<ComparisonUnit[]> split2 = SplitAtParagraphMark(remainingInRight);
                            if (split1.Count() == 1 && split2.Count() == 1)
                            {
                                var csUnknown2 = new CorrelatedSequence
                                {
                                    CorrelationStatus = CorrelationStatus.Unknown,
                                    ComparisonUnitArray1 = split1.First(),
                                    ComparisonUnitArray2 = split2.First()
                                };
                                newSequence.Add(csUnknown2);
                                return newSequence;
                            }

                            if (split1.Count == 2 && split2.Count == 2)
                            {
                                var csUnknown2 = new CorrelatedSequence
                                {
                                    CorrelationStatus = CorrelationStatus.Unknown,
                                    ComparisonUnitArray1 = split1.First(),
                                    ComparisonUnitArray2 = split2.First()
                                };
                                newSequence.Add(csUnknown2);

                                var csUnknown3 = new CorrelatedSequence
                                {
                                    CorrelationStatus = CorrelationStatus.Unknown,
                                    ComparisonUnitArray1 = split1.Skip(1).First(),
                                    ComparisonUnitArray2 = split2.Skip(1).First()
                                };
                                newSequence.Add(csUnknown3);

                                return newSequence;
                            }
                        }
                    }

                    var csUnknown = new CorrelatedSequence
                    {
                        CorrelationStatus = CorrelationStatus.Unknown,
                        ComparisonUnitArray1 = unknown.ComparisonUnitArray1.Skip(countCommonAtBeginning).ToArray(),
                        ComparisonUnitArray2 = unknown.ComparisonUnitArray2.Skip(countCommonAtBeginning).ToArray()
                    };
                    newSequence.Add(csUnknown);
                }
                else if (remainingLeft == 0 && remainingRight == 0)
                {
                    // nothing to do
                }

                return newSequence;
            }

            // if we get to here, then countCommonAtBeginning == 0

            int countCommonAtEnd = unknown
                .ComparisonUnitArray1
                .Reverse()
                .Take(lengthToCompare)
                .Zip(unknown
                        .ComparisonUnitArray2
                        .Reverse()
                        .Take(lengthToCompare),
                    (pu1, pu2) => new
                    {
                        Pu1 = pu1,
                        Pu2 = pu2
                    })
                .TakeWhile(pair => pair.Pu1.SHA1Hash == pair.Pu2.SHA1Hash)
                .Count();

            // never start a common section with a paragraph mark.  However, it is OK to set two paragraph marks as equal.
            while (true)
            {
                if (countCommonAtEnd <= 1)
                    break;

                ComparisonUnit firstCommon = unknown
                    .ComparisonUnitArray1
                    .Reverse()
                    .Take(countCommonAtEnd)
                    .LastOrDefault();

                if (!(firstCommon is ComparisonUnitWord firstCommonWord))
                    break;

                // if the word contains more than one atom, then not a paragraph mark
                if (firstCommonWord.Contents.Count() != 1)
                    break;

                if (!(firstCommonWord.Contents.First() is ComparisonUnitAtom firstCommonAtom))
                    break;

                if (firstCommonAtom.ContentElement.Name != W.pPr)
                    break;

                countCommonAtEnd--;
            }

            var isOnlyParagraphMark = false;
            if (countCommonAtEnd == 1)
            {
                ComparisonUnit firstCommon = unknown
                    .ComparisonUnitArray1
                    .Reverse()
                    .Take(countCommonAtEnd)
                    .LastOrDefault();

                if (firstCommon is ComparisonUnitWord firstCommonWord)
                {
                    // if the word contains more than one atom, then not a paragraph mark
                    if (firstCommonWord.Contents.Count == 1)
                    {
                        if (firstCommonWord.Contents.First() is ComparisonUnitAtom firstCommonAtom)
                        {
                            if (firstCommonAtom.ContentElement.Name == W.pPr)
                                isOnlyParagraphMark = true;
                        }
                    }
                }
            }

            if (countCommonAtEnd == 2)
            {
                ComparisonUnit firstCommon = unknown
                    .ComparisonUnitArray1
                    .Reverse()
                    .Take(countCommonAtEnd)
                    .LastOrDefault();

                ComparisonUnit secondCommon = unknown
                    .ComparisonUnitArray1
                    .Reverse()
                    .Take(countCommonAtEnd)
                    .FirstOrDefault();

                if (firstCommon is ComparisonUnitWord firstCommonWord && secondCommon is ComparisonUnitWord secondCommonWord)
                {
                    // if the word contains more than one atom, then not a paragraph mark
                    if (firstCommonWord.Contents.Count == 1 && secondCommonWord.Contents.Count == 1)
                    {
                        if (firstCommonWord.Contents.First() is ComparisonUnitAtom firstCommonAtom &&
                            secondCommonWord.Contents.First() is ComparisonUnitAtom secondCommonAtom)
                        {
                            if (secondCommonAtom.ContentElement.Name == W.pPr)
                                isOnlyParagraphMark = true;
                        }
                    }
                }
            }

            if (!isOnlyParagraphMark && countCommonAtEnd != 0 &&
                countCommonAtEnd / (double) lengthToCompare < settings.DetailThreshold)
            {
                countCommonAtEnd = 0;
            }

            // If the following test is not there, the test below sets the end paragraph mark of the entire document equal to the end paragraph
            // mark of the first paragraph in the other document, causing lines to be out of order.
            // [InlineData("WC010-Para-Before-Table-Unmodified.docx", "WC010-Para-Before-Table-Mod.docx", 3)]
            if (isOnlyParagraphMark)
            {
                countCommonAtEnd = 0;
            }

            if (countCommonAtEnd == 0)
            {
                return null;
            }

            // if countCommonAtEnd != 0, and if it contains a paragraph mark, then if there are comparison units in the same paragraph before the common at end (in either version)
            // then we want to put all of those comparison units into a single unknown, where they must be resolved against each other.  We don't want those comparison units to go into the middle unknown comparison unit.

            if (countCommonAtEnd != 0)
            {
                var remainingInLeftParagraph = 0;
                var remainingInRightParagraph = 0;

                List<ComparisonUnit> commonEndSeq = unknown
                    .ComparisonUnitArray1
                    .Reverse()
                    .Take(countCommonAtEnd)
                    .Reverse()
                    .ToList();

                ComparisonUnit firstOfCommonEndSeq = commonEndSeq.First();
                if (firstOfCommonEndSeq is ComparisonUnitWord)
                {
                    // are there any paragraph marks in the common seq at end?
                    //if (commonEndSeq.Any(cu => cu.Contents.OfType<ComparisonUnitAtom>().First().ContentElement.Name == W.pPr))
                    if (commonEndSeq.Any(cu =>
                    {
                        ComparisonUnitAtom firstComparisonUnitAtom = cu.Contents.OfType<ComparisonUnitAtom>().FirstOrDefault();
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

                                ComparisonUnitAtom firstComparisonUnitAtom =
                                    cu.Contents.OfType<ComparisonUnitAtom>().FirstOrDefault();
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

                                ComparisonUnitAtom firstComparisonUnitAtom =
                                    cu.Contents.OfType<ComparisonUnitAtom>().FirstOrDefault();
                                if (firstComparisonUnitAtom == null)
                                    return true;

                                return firstComparisonUnitAtom.ContentElement.Name != W.pPr;
                            })
                            .Count();
                    }
                }

                var newSequence = new List<CorrelatedSequence>();

                int beforeCommonParagraphLeft = unknown.ComparisonUnitArray1.Length - remainingInLeftParagraph - countCommonAtEnd;
                int beforeCommonParagraphRight =
                    unknown.ComparisonUnitArray2.Length - remainingInRightParagraph - countCommonAtEnd;

                if (beforeCommonParagraphLeft != 0 && beforeCommonParagraphRight == 0)
                {
                    var csDeleted = new CorrelatedSequence();
                    csDeleted.CorrelationStatus = CorrelationStatus.Deleted;
                    csDeleted.ComparisonUnitArray1 = unknown.ComparisonUnitArray1.Take(beforeCommonParagraphLeft).ToArray();
                    csDeleted.ComparisonUnitArray2 = null;
                    newSequence.Add(csDeleted);
                }
                else if (beforeCommonParagraphLeft == 0 && beforeCommonParagraphRight != 0)
                {
                    var csInserted = new CorrelatedSequence();
                    csInserted.CorrelationStatus = CorrelationStatus.Inserted;
                    csInserted.ComparisonUnitArray1 = null;
                    csInserted.ComparisonUnitArray2 = unknown.ComparisonUnitArray2.Take(beforeCommonParagraphRight).ToArray();
                    newSequence.Add(csInserted);
                }
                else if (beforeCommonParagraphLeft != 0 && beforeCommonParagraphRight != 0)
                {
                    var csUnknown = new CorrelatedSequence();
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
                    var csDeleted = new CorrelatedSequence();
                    csDeleted.CorrelationStatus = CorrelationStatus.Deleted;
                    csDeleted.ComparisonUnitArray1 = unknown.ComparisonUnitArray1.Skip(beforeCommonParagraphLeft)
                        .Take(remainingInLeftParagraph).ToArray();
                    csDeleted.ComparisonUnitArray2 = null;
                    newSequence.Add(csDeleted);
                }
                else if (remainingInLeftParagraph == 0 && remainingInRightParagraph != 0)
                {
                    var csInserted = new CorrelatedSequence();
                    csInserted.CorrelationStatus = CorrelationStatus.Inserted;
                    csInserted.ComparisonUnitArray1 = null;
                    csInserted.ComparisonUnitArray2 = unknown.ComparisonUnitArray2.Skip(beforeCommonParagraphRight)
                        .Take(remainingInRightParagraph).ToArray();
                    newSequence.Add(csInserted);
                }
                else if (remainingInLeftParagraph != 0 && remainingInRightParagraph != 0)
                {
                    var csUnknown = new CorrelatedSequence();
                    csUnknown.CorrelationStatus = CorrelationStatus.Unknown;
                    csUnknown.ComparisonUnitArray1 = unknown.ComparisonUnitArray1.Skip(beforeCommonParagraphLeft)
                        .Take(remainingInLeftParagraph).ToArray();
                    csUnknown.ComparisonUnitArray2 = unknown.ComparisonUnitArray2.Skip(beforeCommonParagraphRight)
                        .Take(remainingInRightParagraph).ToArray();
                    newSequence.Add(csUnknown);
                }
                else if (remainingInLeftParagraph == 0 && remainingInRightParagraph == 0)
                {
                    // nothing to do
                }

                var csEqual = new CorrelatedSequence();
                csEqual.CorrelationStatus = CorrelationStatus.Equal;
                csEqual.ComparisonUnitArray1 = unknown.ComparisonUnitArray1
                    .Skip(unknown.ComparisonUnitArray1.Length - countCommonAtEnd).ToArray();
                csEqual.ComparisonUnitArray2 = unknown.ComparisonUnitArray2
                    .Skip(unknown.ComparisonUnitArray2.Length - countCommonAtEnd).ToArray();
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
                ComparisonUnitAtom atom = cua[i].DescendantContentAtoms().FirstOrDefault();
                if (atom != null && atom.ContentElement.Name == W.pPr)
                    break;
            }

            if (i == cua.Length)
            {
                return new List<ComparisonUnit[]>
                {
                    cua
                };
            }

            return new List<ComparisonUnit[]>
            {
                cua.Take(i).ToArray(),
                cua.Skip(i).ToArray()
            };
        }

        private static object CoalesceRecurse(
            OpenXmlPart part,
            IEnumerable<ComparisonUnitAtom> list,
            int level,
            WmlComparerSettings settings)
        {
            IEnumerable<IGrouping<string, ComparisonUnitAtom>> grouped = list.GroupBy(ca =>
                {
                    if (level >= ca.AncestorElements.Length)
                        return "";

                    return ca.AncestorUnids[level];
                })
                .Where(g => g.Key != "");

            // if there are no deeper children, then we're done.
            if (!grouped.Any())
                return null;

            if (False)
            {
                var sb = new StringBuilder();
                foreach (IGrouping<string, ComparisonUnitAtom> group in grouped)
                {
                    sb.AppendFormat("Group Key: {0}", @group.Key);
                    sb.Append(Environment.NewLine);
                    foreach (ComparisonUnitAtom groupChildItem in @group)
                    {
                        sb.Append("  ");
                        sb.Append(groupChildItem.ToString(0));
                        sb.Append(Environment.NewLine);
                    }

                    sb.Append(Environment.NewLine);
                }

                string sbs = sb.ToString();
                TestUtil.NotePad(sbs);
            }

            List<object> elementList = grouped
                .Select(g =>
                {
                    XElement ancestorBeingConstructed =
                        g.First().AncestorElements[level]; // these will all be the same, by definition

                    // need to group by corr stat
                    List<IGrouping<string, ComparisonUnitAtom>> groupedChildren = g
                        .GroupAdjacent(gc =>
                        {
                            var key = "";
                            if (level < gc.AncestorElements.Length - 1)
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
                        List<object> newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                string[] spl = gc.Key.Split('|');
                                if (spl[0] == "")
                                {
                                    return (object) gc.Select(gcc =>
                                    {
                                        var dup = new XElement(gcc.ContentElement);
                                        if (spl[1] == "Deleted")
                                            dup.Add(new XAttribute(PtOpenXml.Status, "Deleted"));
                                        else if (spl[1] == "Inserted")
                                            dup.Add(new XAttribute(PtOpenXml.Status, "Inserted"));
                                        return dup;
                                    });
                                }

                                return CoalesceRecurse(part, gc, level + 1, settings);
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
                        List<object> newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                string[] spl = gc.Key.Split('|');
                                if (spl[0] == "")
                                {
                                    return (object) gc.Select(gcc =>
                                    {
                                        var dup = new XElement(gcc.ContentElement);
                                        if (spl[1] == "Deleted")
                                            dup.Add(new XAttribute(PtOpenXml.Status, "Deleted"));
                                        else if (spl[1] == "Inserted")
                                            dup.Add(new XAttribute(PtOpenXml.Status, "Inserted"));
                                        return dup;
                                    });
                                }

                                return CoalesceRecurse(part, gc, level + 1, settings);
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
                        List<object> newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                string textOfTextElement = gc.Select(gce => gce.ContentElement.Value).StringConcatenate();
                                bool del = gc.First().CorrelationStatus == CorrelationStatus.Deleted;
                                bool ins = gc.First().CorrelationStatus == CorrelationStatus.Inserted;
                                if (del)
                                    return (object) new XElement(W.delText,
                                        new XAttribute(PtOpenXml.Status, "Deleted"),
                                        GetXmlSpaceAttribute(textOfTextElement),
                                        textOfTextElement);
                                if (ins)
                                    return (object) new XElement(W.t,
                                        new XAttribute(PtOpenXml.Status, "Inserted"),
                                        GetXmlSpaceAttribute(textOfTextElement),
                                        textOfTextElement);

                                return (object) new XElement(W.t,
                                    GetXmlSpaceAttribute(textOfTextElement),
                                    textOfTextElement);
                            })
                            .ToList();
                        return newChildElements;
                    }

                    if (ancestorBeingConstructed.Name == W.drawing)
                    {
                        List<object> newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                bool del = gc.First().CorrelationStatus == CorrelationStatus.Deleted;
                                if (del)
                                {
                                    return (object) gc.Select(gcc =>
                                    {
                                        var newDrawing = new XElement(gcc.ContentElement);
                                        newDrawing.Add(new XAttribute(PtOpenXml.Status, "Deleted"));

                                        OpenXmlPart openXmlPartOfDeletedContent = gc.First().Part;
                                        OpenXmlPart openXmlPartInNewDocument = part;
                                        return gc.Select(gce =>
                                        {
                                            Package packageOfDeletedContent = openXmlPartOfDeletedContent.OpenXmlPackage.Package;
                                            Package packageOfNewContent = openXmlPartInNewDocument.OpenXmlPackage.Package;
                                            PackagePart partInDeletedDocument = packageOfDeletedContent.GetPart(part.Uri);
                                            PackagePart partInNewDocument = packageOfNewContent.GetPart(part.Uri);

                                            return MoveRelatedPartsToDestination(
                                                partInDeletedDocument,
                                                partInNewDocument,
                                                newDrawing);
                                        });
                                    });
                                }

                                bool ins = gc.First().CorrelationStatus == CorrelationStatus.Inserted;
                                if (ins)
                                {
                                    return gc.Select(gcc =>
                                    {
                                        var newDrawing = new XElement(gcc.ContentElement);
                                        newDrawing.Add(new XAttribute(PtOpenXml.Status, "Inserted"));

                                        OpenXmlPart openXmlPartOfInsertedContent = gc.First().Part;
                                        OpenXmlPart openXmlPartInNewDocument = part;
                                        return gc.Select(gce =>
                                        {
                                            Package packageOfSourceContent = openXmlPartOfInsertedContent.OpenXmlPackage.Package;
                                            Package packageOfNewContent = openXmlPartInNewDocument.OpenXmlPackage.Package;
                                            PackagePart partInDeletedDocument = packageOfSourceContent.GetPart(part.Uri);
                                            PackagePart partInNewDocument = packageOfNewContent.GetPart(part.Uri);

                                            return MoveRelatedPartsToDestination(
                                                partInDeletedDocument,
                                                partInNewDocument,
                                                newDrawing);
                                        });
                                    });
                                }

                                return gc.Select(gcc => gcc.ContentElement);
                            })
                            .ToList();

                        return newChildElements;
                    }

                    if (ancestorBeingConstructed.Name == M.oMath || ancestorBeingConstructed.Name == M.oMathPara)
                    {
                        List<IEnumerable<XElement>> newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                bool del = gc.First().CorrelationStatus == CorrelationStatus.Deleted;
                                if (del)
                                {
                                    return gc.Select(gcc =>
                                        new XElement(W.del,
                                            new XAttribute(W.author, settings.AuthorForRevisions),
                                            new XAttribute(W.id, _maxId++),
                                            new XAttribute(W.date, settings.DateTimeForRevisions),
                                            gcc.ContentElement));
                                }

                                bool ins = gc.First().CorrelationStatus == CorrelationStatus.Inserted;
                                if (ins)
                                {
                                    return gc.Select(gcc =>
                                        new XElement(W.ins,
                                            new XAttribute(W.author, settings.AuthorForRevisions),
                                            new XAttribute(W.id, _maxId++),
                                            new XAttribute(W.date, settings.DateTimeForRevisions),
                                            gcc.ContentElement));
                                }

                                return gc.Select(gcc => gcc.ContentElement);
                            })
                            .ToList();
                        return newChildElements;
                    }

                    if (AllowableRunChildren.Contains(ancestorBeingConstructed.Name))
                    {
                        List<IEnumerable<XElement>> newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                bool del = gc.First().CorrelationStatus == CorrelationStatus.Deleted;
                                bool ins = gc.First().CorrelationStatus == CorrelationStatus.Inserted;
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

                                if (ins)
                                {
                                    return gc.Select(gcc =>
                                    {
                                        var dup = new XElement(ancestorBeingConstructed.Name,
                                            ancestorBeingConstructed.Attributes().Where(a => a.Name.Namespace != PtOpenXml.pt),
                                            new XAttribute(PtOpenXml.Status, "Inserted"));
                                        return dup;
                                    });
                                }

                                return gc.Select(gcc => gcc.ContentElement);
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
                        return ReconstructElement(part, g, ancestorBeingConstructed, VML.shapetype, VML.shape, O.OLEObject, level,
                            settings);
                    if (ancestorBeingConstructed.Name == W.ruby)
                        return ReconstructElement(part, g, ancestorBeingConstructed, W.rubyPr, null, null, level, settings);

                    return (object) ReconstructElement(part, g, ancestorBeingConstructed, null, null, null, level, settings);
                })
                .ToList();
            return elementList;
        }

        private static XElement ReconstructElement(
            OpenXmlPart part,
            IGrouping<string, ComparisonUnitAtom> g,
            XElement ancestorBeingConstructed,
            XName props1XName,
            XName props2XName,
            XName props3XName,
            int level,
            WmlComparerSettings settings)
        {
            object newChildElements = CoalesceRecurse(part, g, level + 1, settings);

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

        private static void SetAfterUnids(CorrelatedSequence unknown)
        {
            if (unknown.ComparisonUnitArray1.Length == 1 && unknown.ComparisonUnitArray2.Length == 1)
            {
                if (unknown.ComparisonUnitArray1[0] is ComparisonUnitGroup cua1 &&
                    unknown.ComparisonUnitArray2[0] is ComparisonUnitGroup cua2 &&
                    cua1.ComparisonUnitGroupType == cua2.ComparisonUnitGroupType)
                {
                    ComparisonUnitGroupType groupType = cua1.ComparisonUnitGroupType;
                    IEnumerable<ComparisonUnitAtom> da1 = cua1.DescendantContentAtoms();
                    IEnumerable<ComparisonUnitAtom> da2 = cua2.DescendantContentAtoms();
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
                    foreach (XElement ae in da1.First().AncestorElements)
                    {
                        if (ae.Name != takeThruName)
                        {
                            relevantAncestors.Add(ae);
                            continue;
                        }

                        relevantAncestors.Add(ae);
                        break;
                    }

                    string[] unidList = relevantAncestors
                        .Select(a =>
                        {
                            var unid = (string) a.Attribute(PtOpenXml.Unid);
                            if (unid == null)
                                throw new OpenXmlPowerToolsException("Internal error");

                            return unid;
                        })
                        .ToArray();

                    foreach (ComparisonUnitAtom da in da2)
                    {
                        IEnumerable<XElement> ancestorsToSet = da.AncestorElements.Take(unidList.Length);
                        var zipped = ancestorsToSet.Zip(unidList, (a, u) =>
                            new
                            {
                                Ancestor = a,
                                Unid = u
                            });

                        foreach (var z in zipped)
                        {
                            XAttribute unid = z.Ancestor.Attribute(PtOpenXml.Unid);

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
            int maxd = Math.Min(unknown.ComparisonUnitArray1.Length, unknown.ComparisonUnitArray2.Length);
            if (maxd < 3)
                return null;

            if (unknown.ComparisonUnitArray1.FirstOrDefault() is ComparisonUnitGroup firstInCu1 &&
                unknown.ComparisonUnitArray2.FirstOrDefault() is ComparisonUnitGroup firstInCu2)
            {
                if ((firstInCu1.ComparisonUnitGroupType == ComparisonUnitGroupType.Paragraph ||
                     firstInCu1.ComparisonUnitGroupType == ComparisonUnitGroupType.Table ||
                     firstInCu1.ComparisonUnitGroupType == ComparisonUnitGroupType.Row) &&
                    (firstInCu2.ComparisonUnitGroupType == ComparisonUnitGroupType.Paragraph ||
                     firstInCu2.ComparisonUnitGroupType == ComparisonUnitGroupType.Table ||
                     firstInCu2.ComparisonUnitGroupType == ComparisonUnitGroupType.Row))
                {
                    ComparisonUnitGroupType groupType = firstInCu1.ComparisonUnitGroupType;

                    // Next want to do the lcs algorithm on this.
                    // potentially, we will find all paragraphs are correlated, but they may not be for two reasons-
                    // - if there were changes that were not tracked
                    // - if the anomalies in the change tracking cause there to be a mismatch in the number of paragraphs
                    // therefore we are going to do the whole LCS algorithm thing and at the end of the process, we set
                    // up the correlated sequence list where correlated paragraphs are together in their own unknown
                    // correlated sequence.

                    ComparisonUnit[] cul1 = unknown.ComparisonUnitArray1;
                    ComparisonUnit[] cul2 = unknown.ComparisonUnitArray2;
                    var currentLongestCommonSequenceLength = 0;
                    var currentLongestCommonSequenceAtomCount = 0;
                    int currentI1 = -1;
                    int currentI2 = -1;
                    for (var i1 = 0; i1 < cul1.Length; i1++)
                    {
                        for (var i2 = 0; i2 < cul2.Length; i2++)
                        {
                            var thisSequenceLength = 0;
                            var thisSequenceAtomCount = 0;
                            int thisI1 = i1;
                            int thisI2 = i2;
                            while (true)
                            {
                                bool match = cul1[thisI1] is ComparisonUnitGroup group1 &&
                                             cul2[thisI2] is ComparisonUnitGroup group2 &&
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

                    // here we want to have some sort of threshold, and if the currentLongestCommonSequenceLength is not
                    // longer than the threshold, then don't do anything
                    var doCorrelation = false;
                    if (currentLongestCommonSequenceLength == 1)
                    {
                        int numberOfAtoms1 = unknown.ComparisonUnitArray1[currentI1].DescendantContentAtoms().Count();
                        int numberOfAtoms2 = unknown.ComparisonUnitArray2[currentI2].DescendantContentAtoms().Count();
                        if (numberOfAtoms1 > 16 && numberOfAtoms2 > 16)
                        {
                            doCorrelation = true;
                        }
                    }
                    else if (currentLongestCommonSequenceLength > 1 && currentLongestCommonSequenceLength <= 3)
                    {
                        int numberOfAtoms1 = unknown
                            .ComparisonUnitArray1
                            .Skip(currentI1)
                            .Take(currentLongestCommonSequenceLength)
                            .Select(z => z.DescendantContentAtoms().Count())
                            .Sum();

                        int numberOfAtoms2 = unknown
                            .ComparisonUnitArray2
                            .Skip(currentI2)
                            .Take(currentLongestCommonSequenceLength)
                            .Select(z => z.DescendantContentAtoms().Count())
                            .Sum();

                        if (numberOfAtoms1 > 32 && numberOfAtoms2 > 32)
                        {
                            doCorrelation = true;
                        }
                    }
                    else if (currentLongestCommonSequenceLength > 3)
                    {
                        doCorrelation = true;
                    }

                    if (doCorrelation)
                    {
                        var newListOfCorrelatedSequence = new List<CorrelatedSequence>();

                        if (currentI1 > 0 && currentI2 == 0)
                        {
                            var deletedCorrelatedSequence = new CorrelatedSequence
                            {
                                CorrelationStatus = CorrelationStatus.Deleted,
                                ComparisonUnitArray1 = cul1.Take(currentI1).ToArray(),
                                ComparisonUnitArray2 = null
                            };
                            newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                        }
                        else if (currentI1 == 0 && currentI2 > 0)
                        {
                            var insertedCorrelatedSequence = new CorrelatedSequence
                            {
                                CorrelationStatus = CorrelationStatus.Inserted,
                                ComparisonUnitArray1 = null,
                                ComparisonUnitArray2 = cul2.Take(currentI2).ToArray()
                            };
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                        }
                        else if (currentI1 > 0 && currentI2 > 0)
                        {
                            var unknownCorrelatedSequence = new CorrelatedSequence
                            {
                                CorrelationStatus = CorrelationStatus.Unknown,
                                ComparisonUnitArray1 = cul1.Take(currentI1).ToArray(),
                                ComparisonUnitArray2 = cul2.Take(currentI2).ToArray()
                            };
                            newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
                        }
                        else if (currentI1 == 0 && currentI2 == 0)
                        {
                            // nothing to do
                        }

                        for (var i = 0; i < currentLongestCommonSequenceLength; i++)
                        {
                            var unknownCorrelatedSequence = new CorrelatedSequence
                            {
                                CorrelationStatus = CorrelationStatus.Unknown,
                                ComparisonUnitArray1 = cul1
                                    .Skip(currentI1)
                                    .Skip(i)
                                    .Take(1)
                                    .ToArray(),
                                ComparisonUnitArray2 = cul2
                                    .Skip(currentI2)
                                    .Skip(i)
                                    .Take(1)
                                    .ToArray()
                            };
                            newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
                        }

                        int endI1 = currentI1 + currentLongestCommonSequenceLength;
                        int endI2 = currentI2 + currentLongestCommonSequenceLength;

                        if (endI1 < cul1.Length && endI2 == cul2.Length)
                        {
                            var deletedCorrelatedSequence = new CorrelatedSequence
                            {
                                CorrelationStatus = CorrelationStatus.Deleted,
                                ComparisonUnitArray1 = cul1.Skip(endI1).ToArray(),
                                ComparisonUnitArray2 = null
                            };
                            newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                        }
                        else if (endI1 == cul1.Length && endI2 < cul2.Length)
                        {
                            var insertedCorrelatedSequence = new CorrelatedSequence
                            {
                                CorrelationStatus = CorrelationStatus.Inserted,
                                ComparisonUnitArray1 = null,
                                ComparisonUnitArray2 = cul2.Skip(endI2).ToArray()
                            };
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                        }
                        else if (endI1 < cul1.Length && endI2 < cul2.Length)
                        {
                            var unknownCorrelatedSequence = new CorrelatedSequence
                            {
                                CorrelationStatus = CorrelationStatus.Unknown,
                                ComparisonUnitArray1 = cul1.Skip(endI1).ToArray(),
                                ComparisonUnitArray2 = cul2.Skip(endI2).ToArray()
                            };
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
    }
}
