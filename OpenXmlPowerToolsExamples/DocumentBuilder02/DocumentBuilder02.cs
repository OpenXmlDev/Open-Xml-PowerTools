// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

/// <summary>
/// Demonstrates muliple operations:
/// (1) insert content into a document;
/// (2) delete content from a document;
/// (3) shred a document; and
/// (4) assemble it again, insert TOC.
/// </summary>
internal class DocumentBuilderExample02
{
    private static void Main()
    {
        DateTime n = DateTime.Now;
        var tempDi = new DirectoryInfo(
            $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");

        tempDi.Create();

        // Insert an abstract and author biography into a white paper.
        var sources = new List<Source>
        {
            new(new WmlDocument("../../../WhitePaper.docx"), 0, 1, true),
            new(new WmlDocument("../../../Abstract.docx"), false),
            new(new WmlDocument("../../../AuthorBiography.docx"), false),
            new(new WmlDocument("../../../WhitePaper.docx"), 1, false)
        };

        DocumentBuilder.BuildDocument(sources, Path.Combine(tempDi.FullName, "AssembledPaper.docx"));

        // Delete all paragraphs with a specific style.
        using (WordprocessingDocument doc = WordprocessingDocument.Open("../../../Notes.docx", false))
        {
            sources = doc
                .MainDocumentPart!
                .GetXElement()!
                .Elements(W.body)
                .Elements()
                .Select((p, i) => new
                {
                    Paragraph = p,
                    Index = i
                })
                .GroupAdjacent(pi => (string)pi.Paragraph
                    .Elements(W.pPr)
                    .Elements(W.pStyle)
                    .Attributes(W.val)
                    .FirstOrDefault() != "Note")
                .Where(g => g.Key)
                .Select(g => new Source(
                    new WmlDocument("../../../Notes.docx"), g.First().Index,
                    g.Last().Index - g.First().Index + 1, true))
                .ToList();
        }

        DocumentBuilder.BuildDocument(sources, Path.Combine(tempDi.FullName, "NewNotes.docx"));

        // Shred a document into multiple parts for each section
        List<DocumentInfo> documentList;
        using (WordprocessingDocument doc = WordprocessingDocument.Open("../../../Spec.docx", false))
        {
            IEnumerable<int> sectionCounts = doc
                .MainDocumentPart
                .GetXElement()
                .Elements(W.body)
                .Elements()
                .Rollup(0, (pi, last) => (string)pi
                    .Elements(W.pPr)
                    .Elements(W.pStyle)
                    .Attributes(W.val)
                    .FirstOrDefault() == "Heading1"
                    ? last + 1
                    : last);

            var beforeZipped = doc
                .MainDocumentPart
                .GetXElement()
                .Elements(W.body)
                .Elements()
                .Select((p, i) => new
                {
                    Paragraph = p,
                    Index = i
                });

            var zipped = beforeZipped.PtZip(sectionCounts, (pi, sc) => new
            {
                pi.Paragraph,
                pi.Index,
                SectionIndex = sc
            });

            documentList = zipped
                .GroupAdjacent(p => p.SectionIndex)
                .Select(g => new DocumentInfo
                {
                    DocumentNumber = g.Key,
                    Start = g.First().Index,
                    Count = g.Last().Index - g.First().Index + 1
                })
                .ToList();
        }

        foreach (DocumentInfo doc in documentList)
        {
            var fileName = $"Section{doc.DocumentNumber:000}.docx";
            var documentSource = new List<Source>
            {
                new(new WmlDocument("../../../Spec.docx"), doc.Start, doc.Count, true)
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

    private class DocumentInfo
    {
        public int DocumentNumber;
        public int Start;
        public int Count;
    }
}
