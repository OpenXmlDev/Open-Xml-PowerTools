// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static partial class WmlComparer
    {
        private static WmlDocument PreProcessMarkup(WmlDocument source, int startingIdForFootnotesEndnotes)
        {
            // open and close to get rid of MC content
            using (var ms = new MemoryStream())
            {
                ms.Write(source.DocumentByteArray, 0, source.DocumentByteArray.Length);
                var os = new OpenSettings
                {
                    MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(
                        MarkupCompatibilityProcessMode.ProcessAllParts,
                        FileFormatVersions.Office2007)
                };

                using (var wDoc = WordprocessingDocument.Open(ms, true, os))
                {
                    var unused = wDoc.MainDocumentPart.RootElement;
                    if (wDoc.MainDocumentPart.FootnotesPart != null)
                    {
                        // contrary to what you might think, looking at the API, it is necessary to access the root element of each part to cause
                        // the SDK to process MC markup.
                        var unused1 = wDoc.MainDocumentPart.FootnotesPart.RootElement;
                    }

                    if (wDoc.MainDocumentPart.EndnotesPart != null)
                    {
                        var unused1 = wDoc.MainDocumentPart.EndnotesPart.RootElement;
                    }
                }

                source = new WmlDocument(source.FileName, ms.ToArray());
            }

            // open and close to get rid of MC content
            using (var ms = new MemoryStream())
            {
                ms.Write(source.DocumentByteArray, 0, source.DocumentByteArray.Length);
                var os = new OpenSettings
                {
                    MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(
                        MarkupCompatibilityProcessMode.ProcessAllParts,
                        FileFormatVersions.Office2007)
                };

                using (var wDoc = WordprocessingDocument.Open(ms, true, os))
                {
                    TestForInvalidContent(wDoc);
                    RemoveExistingPowerToolsMarkup(wDoc);

                    // Removing content controls, field codes, and bookmarks is a no-no for many use cases.
                    // We need content controls, e.g., on the title page. Field codes are required for
                    // automatic cross-references, which require bookmarks.
                    // TODO: Revisit
                    var msSettings = new SimplifyMarkupSettings
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
                        RemoveHyperlinks = true
                    };
                    MarkupSimplifier.SimplifyMarkup(wDoc, msSettings);
                    ChangeFootnoteEndnoteReferencesToUniqueRange(wDoc, startingIdForFootnotesEndnotes);
                    AddUnidsToMarkupInContentParts(wDoc);
                    AddFootnotesEndnotesParts(wDoc);
                    FillInEmptyFootnotesEndnotes(wDoc);
                }

                return new WmlDocument(source.FileName, ms.ToArray());
            }
        }

        private static void TestForInvalidContent(WordprocessingDocument wDoc)
        {
            foreach (var part in wDoc.ContentParts())
            {
                var xDoc = part.GetXDocument();
                if (xDoc.Descendants(W.altChunk).Any())
                {
                    throw new OpenXmlPowerToolsException("Unsupported document, contains w:altChunk");
                }

                if (xDoc.Descendants(W.subDoc).Any())
                {
                    throw new OpenXmlPowerToolsException("Unsupported document, contains w:subDoc");
                }

                if (xDoc.Descendants(W.contentPart).Any())
                {
                    throw new OpenXmlPowerToolsException("Unsupported document, contains w:contentPart");
                }
            }
        }

        private static void RemoveExistingPowerToolsMarkup(WordprocessingDocument wDoc)
        {
            wDoc.MainDocumentPart
                .GetXDocument()
                .Root?
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
                    .Root?
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
                    .Root?
                    .Descendants()
                    .Attributes()
                    .Where(a => a.Name.Namespace == PtOpenXml.pt)
                    .Where(a => a.Name != PtOpenXml.Unid)
                    .Remove();

                enPart.PutXDocument();
            }
        }

        private static void ChangeFootnoteEndnoteReferencesToUniqueRange(
            WordprocessingDocument wDoc,
            int startingIdForFootnotesEndnotes)
        {
            var mainDocPart = wDoc.MainDocumentPart;
            var footnotesPart = wDoc.MainDocumentPart.FootnotesPart;
            var endnotesPart = wDoc.MainDocumentPart.EndnotesPart;

            var document =
                mainDocPart.GetXDocument().Root ?? throw new OpenXmlPowerToolsException("Invalid document.");

            var footnotes = footnotesPart?.GetXDocument().Root;
            var endnotes = endnotesPart?.GetXDocument().Root;

            var references = document
                .Descendants()
                .Where(d => d.Name == W.footnoteReference || d.Name == W.endnoteReference);

            foreach (var r in references)
            {
                var oldId = (string)r.Attribute(W.id);
                var newId = startingIdForFootnotesEndnotes.ToString();
                startingIdForFootnotesEndnotes++;
                r.SetAttributeValue(W.id, newId);
                if (r.Name == W.footnoteReference)
                {
                    var fn = footnotes?
                        .Elements()
                        .FirstOrDefault(e => (string)e.Attribute(W.id) == oldId);

                    if (fn == null)
                    {
                        throw new OpenXmlPowerToolsException("Invalid document");
                    }

                    fn.SetAttributeValue(W.id, newId);
                }
                else
                {
                    var en = endnotes?
                        .Elements()
                        .FirstOrDefault(e => (string)e.Attribute(W.id) == oldId);

                    if (en == null)
                    {
                        throw new OpenXmlPowerToolsException("Invalid document");
                    }

                    en.SetAttributeValue(W.id, newId);
                }
            }

            mainDocPart.PutXDocument();
            footnotesPart?.PutXDocument();
            endnotesPart?.PutXDocument();
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

        private static void AssignUnidToAllElements(XElement contentParent)
        {
            var content = contentParent.Descendants();
            foreach (var d in content)
            {
                if (d.Attribute(PtOpenXml.Unid) == null)
                {
                    var unid = Guid.NewGuid().ToString().Replace("-", "");
                    var newAtt = new XAttribute(PtOpenXml.Unid, unid);
                    d.Add(newAtt);
                }
            }
        }

        [SuppressMessage("ReSharper", "CoVariantArrayConversion")]
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

        private static void FillInEmptyFootnotesEndnotes(WordprocessingDocument wDoc)
        {
            var emptyFootnote = XElement.Parse(
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

            var emptyEndnote = XElement.Parse(
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
                var fnRoot = footnotePart.GetXDocument().Root ?? throw new ArgumentException();
                foreach (var fn in fnRoot.Elements(W.footnote))
                {
                    if (!fn.HasElements)
                    {
                        fn.Add(emptyFootnote);
                    }
                }

                footnotePart.PutXDocument();
            }

            var endnotePart = wDoc.MainDocumentPart.EndnotesPart;
            if (endnotePart != null)
            {
                var fnRoot = endnotePart.GetXDocument().Root ?? throw new ArgumentException();
                foreach (var fn in fnRoot.Elements(W.endnote))
                {
                    if (!fn.HasElements)
                    {
                        fn.Add(emptyEndnote);
                    }
                }

                endnotePart.PutXDocument();
            }
        }
    }
}