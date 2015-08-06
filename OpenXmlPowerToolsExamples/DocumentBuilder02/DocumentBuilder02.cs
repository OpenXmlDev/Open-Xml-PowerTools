#define WhitePaper
#define DeleteNotes
#define Shred
#define Reassemble

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

// Insert content into a document
// Delete content from a document
// Shred a document
// Assemble it again, insert TOC

class DocumentBuilderExample02
{
#if Shred
    private class DocumentInfo
    {
        public int DocumentNumber;
        public int Start;
        public int Count;
    }
#endif

    static void Main(string[] args)
    {
        // Insert an abstract and author biography into a white paper.
        List<Source> sources = null;

#if WhitePaper
        sources = new List<Source>()
        {
            new Source(new WmlDocument("../../WhitePaper.docx"), 0, 1, true),
            new Source(new WmlDocument("../../Abstract.docx"), false),
            new Source(new WmlDocument("../../AuthorBiography.docx"), false),
            new Source(new WmlDocument("../../WhitePaper.docx"), 1, false),
        };
        DocumentBuilder.BuildDocument(sources, "AssembledPaper.docx");
#endif

#if DeleteNotes
        // Delete all paragraphs with a specific style.
        using (WordprocessingDocument doc =
            WordprocessingDocument.Open("../../Notes.docx", false))
        {
            sources = doc
                .MainDocumentPart
                .GetXDocument()
                .Root
                .Element(W.body)
                .Elements()
                .Select((p, i) => new
                {
                    Paragraph = p,
                    Index = i,
                })
                .GroupAdjacent(pi => (string)pi.Paragraph
                    .Elements(W.pPr)
                    .Elements(W.pStyle)
                    .Attributes(W.val)
                    .FirstOrDefault() != "Note")
                .Where(g => g.Key == true)
                .Select(g => new Source(
                    new WmlDocument("../../Notes.docx"), g.First().Index,
                        g.Last().Index - g.First().Index + 1, true))
                .ToList();
        }
        DocumentBuilder.BuildDocument(sources, "NewNotes.docx");
#endif

#if Shred
        // Shred a document into multiple parts for each section
        List<DocumentInfo> documentList;
        using (WordprocessingDocument doc =
            WordprocessingDocument.Open("../../Spec.docx", false))
        {
            var sectionCounts = doc
                .MainDocumentPart
                .GetXDocument()
                .Root
                .Element(W.body)
                .Elements()
                .Rollup(0, (pi, last) => (string)pi
                    .Elements(W.pPr)
                    .Elements(W.pStyle)
                    .Attributes(W.val)
                    .FirstOrDefault() == "Heading1" ? last + 1 : last);
            var beforeZipped = doc
                .MainDocumentPart
                .GetXDocument()
                .Root
                .Element(W.body)
                .Elements()
                .Select((p, i) => new
                {
                    Paragraph = p,
                    Index = i,
                });
            var zipped = PtExtensions.PtZip(beforeZipped, sectionCounts, (pi, sc) => new
                {
                    Paragraph = pi.Paragraph,
                    Index = pi.Index,
                    SectionIndex = sc,
                });
            documentList = zipped
                .GroupAdjacent(p => p.SectionIndex)
                .Select(g => new DocumentInfo
                {
                    DocumentNumber = g.Key,
                    Start = g.First().Index,
                    Count = g.Last().Index - g.First().Index + 1,
                })
                .ToList();
        }
        foreach (var doc in documentList)
        {
            string fileName = String.Format("Section{0:000}.docx", doc.DocumentNumber);
            List<Source> documentSource = new List<Source> {
                new Source(new WmlDocument("../../Spec.docx"), doc.Start, doc.Count, true)
            };
            DocumentBuilder.BuildDocument(documentSource, fileName);
        }
#endif

#if Reassemble
        // Re-assemble the parts into a single document.
        sources = new DirectoryInfo(".")
            .GetFiles("Section*.docx")
            .Select(d => new Source(new WmlDocument(d.FullName), true))
            .ToList();
        DocumentBuilder.BuildDocument(sources, "ReassembledSpec.docx");
        using (WordprocessingDocument doc =
            WordprocessingDocument.Open("ReassembledSpec.docx", true))
        {
            ReferenceAdder.AddToc(doc, "/w:document/w:body/w:p[1]",
                @"TOC \o '1-3' \h \z \u", null, null);
        }
#endif
    }
}
