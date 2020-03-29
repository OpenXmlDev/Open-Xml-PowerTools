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
        // the following gets a flattened list of ComparisonUnitAtoms, with status indicated in each ComparisonUnitAtom: Deleted, Inserted, or Equal

        // for any deleted or inserted rows, we go into the w:trPr properties, and add the appropriate w:ins or w:del element, and therefore
        // when generating the document, the appropriate row will be marked as deleted or inserted.

        public static List<WmlComparerRevision> GetRevisions(WmlDocument source, WmlComparerSettings settings)
        {
            using (var ms = new MemoryStream())
            {
                ms.Write(source.DocumentByteArray, 0, source.DocumentByteArray.Length);
                using (var wDoc = WordprocessingDocument.Open(ms, true))
                {
                    TestForInvalidContent(wDoc);
                    RemoveExistingPowerToolsMarkup(wDoc);

                    var contentParent = wDoc.MainDocumentPart.GetXDocument().Root?.Element(W.body);
                    var atomList =
                        CreateComparisonUnitAtomList(wDoc.MainDocumentPart, contentParent, settings).ToArray();

                    if (False)
                    {
                        var sb = new StringBuilder();
                        foreach (var item in atomList)
                        {
                            sb.Append(item + Environment.NewLine);
                        }

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
                                    new XAttribute(XNamespace.Xmlns + "w",
                                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
                                    a.RevTrackElement.Attributes().Where(a2 => a2.Name != W.id && a2.Name != PtOpenXml.Unid));
                                key += rt.ToString(SaveOptions.DisableFormatting);
                            }

                            return key;
                        })
                        .ToList();

                    var revisions = grouped
                        .Where(k => k.Key != "Equal")
                        .ToList();

                    if (False)
                    {
                        var sb = new StringBuilder();
                        foreach (var item in revisions)
                        {
                            sb.Append(item.Key + Environment.NewLine);
                        }

                        var sbs = sb.ToString();
                        TestUtil.NotePad(sbs);
                    }

                    var mainDocPartRevisionList = revisions
                        .Select(rg =>
                        {
                            var rev = new WmlComparerRevision();
                            if (rg.Key.StartsWith("Inserted"))
                            {
                                rev.RevisionType = WmlComparerRevisionType.Inserted;
                            }
                            else if (rg.Key.StartsWith("Deleted"))
                            {
                                rev.RevisionType = WmlComparerRevisionType.Deleted;
                            }

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
                                    .Select(rgc => rgc.ContentElement.Name == W.pPr ? NewLine : rgc.ContentElement.Value)
                                    .StringConcatenate();
                            }

                            return rev;
                        })
                        .ToList();

                    var footnotesRevisionList =
                        GetFootnoteEndnoteRevisionList(wDoc.MainDocumentPart.FootnotesPart, W.footnote, settings);
                    var endnotesRevisionList =
                        GetFootnoteEndnoteRevisionList(wDoc.MainDocumentPart.EndnotesPart, W.endnote, settings);

                    var finalRevisionList = mainDocPartRevisionList
                        .Concat(footnotesRevisionList)
                        .Concat(endnotesRevisionList)
                        .ToList();

                    return finalRevisionList;
                }
            }
        }

        private static IEnumerable<WmlComparerRevision> GetFootnoteEndnoteRevisionList(
            OpenXmlPart footnotesEndnotesPart,
            XName footnoteEndnoteElementName,
            WmlComparerSettings settings)
        {
            if (footnotesEndnotesPart == null)
            {
                return Enumerable.Empty<WmlComparerRevision>();
            }

            var xDoc = footnotesEndnotesPart.GetXDocument();
            var footnotesEndnotes =
                xDoc.Root?.Elements(footnoteEndnoteElementName) ?? throw new OpenXmlPowerToolsException("Invalid document.");

            var revisionsForPart = new List<WmlComparerRevision>();
            foreach (var fn in footnotesEndnotes)
            {
                var atomList = CreateComparisonUnitAtomList(footnotesEndnotesPart, fn, settings).ToArray();

                if (False)
                {
                    var sb = new StringBuilder();
                    foreach (var item in atomList)
                    {
                        sb.Append(item + Environment.NewLine);
                    }

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
                                new XAttribute(XNamespace.Xmlns + "w",
                                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
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
                        {
                            rev.RevisionType = WmlComparerRevisionType.Inserted;
                        }
                        else if (rg.Key.StartsWith("Deleted"))
                        {
                            rev.RevisionType = WmlComparerRevisionType.Deleted;
                        }

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
                                .Select(rgc => rgc.ContentElement.Name == W.pPr ? NewLine : rgc.ContentElement.Value)
                                .StringConcatenate();
                        }

                        return rev;
                    });

                revisionsForPart.AddRange(thisNoteRevisionList);
            }

            return revisionsForPart;
        }
    }
}