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
    private class DocumentInfo
    {
        public int DocumentNumber;
        public int Start;
        public int Count;
    }

    static void Main(string[] args)
    {
        var n = DateTime.Now;
        var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
        tempDi.Create();

        // Insert an abstract and author biography into a white paper.
        List<Source> sources = null;

        sources = new List<Source>()
        {
            new Source(new WmlDocument("../../WhitePaper.docx"), 0, 1, true),
            new Source(new WmlDocument("../../Abstract.docx"), false),
            new Source(new WmlDocument("../../AuthorBiography.docx"), false),
            new Source(new WmlDocument("../../WhitePaper.docx"), 1, false),
        };
        DocumentBuilder.BuildDocument(sources, Path.Combine(tempDi.FullName, "AssembledPaper.docx"));

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
        DocumentBuilder.BuildDocument(sources, Path.Combine(tempDi.FullName, "NewNotes.docx"));

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
            DocumentBuilder.BuildDocument(documentSource, Path.Combine(tempDi.FullName, fileName));
        }

        // Re-assemble the parts into a single document.
        sources = tempDi
            .GetFiles("Section*.docx")
            .Select(d => new Source(new WmlDocument(d.FullName), true))
            .ToList();
        DocumentBuilder.BuildDocument(sources, Path.Combine(tempDi.FullName, "ReassembledSpec.docx"));
    }
}
