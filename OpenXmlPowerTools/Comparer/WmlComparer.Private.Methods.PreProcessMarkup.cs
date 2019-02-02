// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

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

                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true, os))
                {
                    OpenXmlPartRootElement unused = wDoc.MainDocumentPart.RootElement;
                    if (wDoc.MainDocumentPart.FootnotesPart != null)
                    {
                        // contrary to what you might think, looking at the API, it is necessary to access the root element of each part to cause
                        // the SDK to process MC markup.
                        OpenXmlPartRootElement unused1 = wDoc.MainDocumentPart.FootnotesPart.RootElement;
                    }

                    if (wDoc.MainDocumentPart.EndnotesPart != null)
                    {
                        OpenXmlPartRootElement unused1 = wDoc.MainDocumentPart.EndnotesPart.RootElement;
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

                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true, os))
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
            foreach (OpenXmlPart part in wDoc.ContentParts())
            {
                XDocument xDoc = part.GetXDocument();
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
                .Root?
                .Descendants()
                .Attributes()
                .Where(a => a.Name.Namespace == PtOpenXml.pt)
                .Where(a => a.Name != PtOpenXml.Unid)
                .Remove();

            wDoc.MainDocumentPart.PutXDocument();

            FootnotesPart fnPart = wDoc.MainDocumentPart.FootnotesPart;
            if (fnPart != null)
            {
                XDocument fnXDoc = fnPart.GetXDocument();
                fnXDoc
                    .Root?
                    .Descendants()
                    .Attributes()
                    .Where(a => a.Name.Namespace == PtOpenXml.pt)
                    .Where(a => a.Name != PtOpenXml.Unid)
                    .Remove();

                fnPart.PutXDocument();
            }

            EndnotesPart enPart = wDoc.MainDocumentPart.EndnotesPart;
            if (enPart != null)
            {
                XDocument enXDoc = enPart.GetXDocument();
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
            MainDocumentPart mainDocPart = wDoc.MainDocumentPart;
            FootnotesPart footnotesPart = wDoc.MainDocumentPart.FootnotesPart;
            EndnotesPart endnotesPart = wDoc.MainDocumentPart.EndnotesPart;

            XElement document =
                mainDocPart.GetXDocument().Root ?? throw new OpenXmlPowerToolsException("Invalid document.");

            XElement footnotes = footnotesPart?.GetXDocument().Root;
            XElement endnotes = endnotesPart?.GetXDocument().Root;

            IEnumerable<XElement> references = document
                .Descendants()
                .Where(d => d.Name == W.footnoteReference || d.Name == W.endnoteReference);

            foreach (XElement r in references)
            {
                var oldId = (string) r.Attribute(W.id);
                string newId = startingIdForFootnotesEndnotes.ToString();
                startingIdForFootnotesEndnotes++;
                r.SetAttributeValue(W.id, newId);
                if (r.Name == W.footnoteReference)
                {
                    XElement fn = footnotes?
                        .Elements()
                        .FirstOrDefault(e => (string) e.Attribute(W.id) == oldId);

                    if (fn == null)
                    {
                        throw new OpenXmlPowerToolsException("Invalid document");
                    }

                    fn.SetAttributeValue(W.id, newId);
                }
                else
                {
                    XElement en = endnotes?
                        .Elements()
                        .FirstOrDefault(e => (string) e.Attribute(W.id) == oldId);

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
            XDocument mdp = wDoc.MainDocumentPart.GetXDocument();
            AssignUnidToAllElements(mdp.Root);
            IgnorePt14Namespace(mdp.Root);
            wDoc.MainDocumentPart.PutXDocument();

            if (wDoc.MainDocumentPart.FootnotesPart != null)
            {
                XDocument p = wDoc.MainDocumentPart.FootnotesPart.GetXDocument();
                AssignUnidToAllElements(p.Root);
                IgnorePt14Namespace(p.Root);
                wDoc.MainDocumentPart.FootnotesPart.PutXDocument();
            }

            if (wDoc.MainDocumentPart.EndnotesPart != null)
            {
                XDocument p = wDoc.MainDocumentPart.EndnotesPart.GetXDocument();
                AssignUnidToAllElements(p.Root);
                IgnorePt14Namespace(p.Root);
                wDoc.MainDocumentPart.EndnotesPart.PutXDocument();
            }
        }

        private static void AssignUnidToAllElements(XElement contentParent)
        {
            IEnumerable<XElement> content = contentParent.Descendants();
            foreach (XElement d in content)
            {
                if (d.Attribute(PtOpenXml.Unid) == null)
                {
                    string unid = Guid.NewGuid().ToString().Replace("-", "");
                    var newAtt = new XAttribute(PtOpenXml.Unid, unid);
                    d.Add(newAtt);
                }
            }
        }

        [SuppressMessage("ReSharper", "CoVariantArrayConversion")]
        private static void AddFootnotesEndnotesParts(WordprocessingDocument wDoc)
        {
            MainDocumentPart mdp = wDoc.MainDocumentPart;
            if (mdp.FootnotesPart == null)
            {
                mdp.AddNewPart<FootnotesPart>();
                XDocument newFootnotes = wDoc.MainDocumentPart.FootnotesPart.GetXDocument();
                newFootnotes.Declaration.Standalone = "yes";
                newFootnotes.Declaration.Encoding = "UTF-8";
                newFootnotes.Add(new XElement(W.footnotes, NamespaceAttributes));
                mdp.FootnotesPart.PutXDocument();
            }

            if (mdp.EndnotesPart == null)
            {
                mdp.AddNewPart<EndnotesPart>();
                XDocument newEndnotes = wDoc.MainDocumentPart.EndnotesPart.GetXDocument();
                newEndnotes.Declaration.Standalone = "yes";
                newEndnotes.Declaration.Encoding = "UTF-8";
                newEndnotes.Add(new XElement(W.endnotes, NamespaceAttributes));
                mdp.EndnotesPart.PutXDocument();
            }
        }

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

            FootnotesPart footnotePart = wDoc.MainDocumentPart.FootnotesPart;
            if (footnotePart != null)
            {
                XElement fnRoot = footnotePart.GetXDocument().Root ?? throw new ArgumentException();
                foreach (XElement fn in fnRoot.Elements(W.footnote))
                {
                    if (!fn.HasElements)
                        fn.Add(emptyFootnote);
                }

                footnotePart.PutXDocument();
            }

            EndnotesPart endnotePart = wDoc.MainDocumentPart.EndnotesPart;
            if (endnotePart != null)
            {
                XElement fnRoot = endnotePart.GetXDocument().Root ?? throw new ArgumentException();
                foreach (XElement fn in fnRoot.Elements(W.endnote))
                {
                    if (!fn.HasElements)
                        fn.Add(emptyEndnote);
                }

                endnotePart.PutXDocument();
            }
        }
    }
}
