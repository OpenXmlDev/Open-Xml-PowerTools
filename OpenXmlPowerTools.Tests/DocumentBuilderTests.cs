/***************************************************************************

Copyright (c) Microsoft Corporation 2012-2015.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

***************************************************************************/

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenXmlPowerTools;
using Xunit;

namespace OxPt
{
    public class DbTests
    {
        [Fact]
        public void DB001_DocumentBuilderKeepSections()
        {
            string name = "DB001-Sections.docx";
            FileInfo sourceDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            List<Source> sources = null;
            sources = new List<Source>()
            {
                new Source(new WmlDocument(sourceDocx.FullName), true),
            };
            var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-processed-by-DocumentBuilder.docx")));
            DocumentBuilder.BuildDocument(sources, processedDestDocx.FullName);
        }

        [Fact]
        public void DB002_DocumentBuilderKeepSectionsDiscardHeaders()
        {
            FileInfo source1Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB002-Sections-With-Headers.docx"));
            FileInfo source2Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB002-Landscape-Section.docx"));

            List<Source> sources = null;
            sources = new List<Source>()
            {
                new Source(new WmlDocument(source1Docx.FullName)) { KeepSections = true },
                new Source(new WmlDocument(source2Docx.FullName)) { KeepSections = true, DiscardHeadersAndFootersInKeptSections = true },
                new Source(new WmlDocument(source1Docx.FullName)) { KeepSections = true },
            };
            var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB002-Keep-Sections-Discard-Headers-And-Footers.docx"));
            DocumentBuilder.BuildDocument(sources, processedDestDocx.FullName);
        }

        [Fact]
        public void DB003_DocumentBuilderOnlyDefaultHeader()
        {
            FileInfo source1Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB003-Only-Default-Header.docx"));
            FileInfo source2Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB002-Landscape-Section.docx"));

            List<Source> sources = null;
            sources = new List<Source>()
            {
                new Source(new WmlDocument(source1Docx.FullName)) { KeepSections = true },
                new Source(new WmlDocument(source2Docx.FullName)) { KeepSections = true, DiscardHeadersAndFootersInKeptSections = true },
                new Source(new WmlDocument(source1Docx.FullName)) { KeepSections = true },
            };
            var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB003-Only-Default-Header.docx"));
            DocumentBuilder.BuildDocument(sources, processedDestDocx.FullName);
        }

        [Fact]
        public void DB004_DocumentBuilderNoHeaders()
        {
            FileInfo source1Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB004-No-Headers.docx"));
            FileInfo source2Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB002-Landscape-Section.docx"));

            List<Source> sources = null;
            sources = new List<Source>()
            {
                new Source(new WmlDocument(source1Docx.FullName)) { KeepSections = true },
                new Source(new WmlDocument(source2Docx.FullName)) { KeepSections = true, DiscardHeadersAndFootersInKeptSections = true },
                new Source(new WmlDocument(source1Docx.FullName)) { KeepSections = true },
            };
            var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB003-Only-Default-Header.docx"));
            DocumentBuilder.BuildDocument(sources, processedDestDocx.FullName);
        }

        [Fact]
        public void DB005_HeadersWithRefsToImages()
        {
            FileInfo source1Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB005-Headers-With-Images.docx"));
            FileInfo source2Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB002-Landscape-Section.docx"));

            List<Source> sources = null;
            sources = new List<Source>()
            {
                new Source(new WmlDocument(source1Docx.FullName)) { KeepSections = true },
                new Source(new WmlDocument(source2Docx.FullName)) { KeepSections = true, DiscardHeadersAndFootersInKeptSections = true },
                new Source(new WmlDocument(source1Docx.FullName)) { KeepSections = true },
            };
            var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB005.docx"));
            DocumentBuilder.BuildDocument(sources, processedDestDocx.FullName);
        }

        [Fact]
        public void DB006_Example_DocumentBuilder01()
        {
            FileInfo source1 = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB006-Source1.docx"));
            FileInfo source2 = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB006-Source2.docx"));
            FileInfo source3 = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB006-Source3.docx"));
            List<Source> sources = null;

            // Create new document from 10 paragraphs starting at paragraph 5 of Source1.docx
            sources = new List<Source>()
            {
                new Source(new WmlDocument(source1.FullName), 5, 10, true),
            };
            var out1 = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB006-Out1.docx"));
            DocumentBuilder.BuildDocument(sources, out1.FullName);
            Validate(out1);

            // Create new document from paragraph 1, and paragraphs 5 through end of Source3.docx.
            // This effectively 'deletes' paragraphs 2-4
            sources = new List<Source>()
            {
                new Source(new WmlDocument(source3.FullName), 0, 1, false),
                new Source(new WmlDocument(source3.FullName), 4, false),
            };
            var out2 = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB006-Out2.docx"));
            DocumentBuilder.BuildDocument(sources, out2.FullName);
            Validate(out2);

            // Create a new document that consists of the entirety of Source1.docx and Source2.docx.  Use
            // the section information (headings and footers) from source1.
            sources = new List<Source>()
            {
                new Source(new WmlDocument(source1.FullName), true),
                new Source(new WmlDocument(source2.FullName), false),
            };
            var out3 = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB006-Out3.docx"));
            DocumentBuilder.BuildDocument(sources, out3.FullName);
            Validate(out3);

            // Create a new document that consists of the entirety of Source1.docx and Source2.docx.  Use
            // the section information (headings and footers) from source2.
            sources = new List<Source>()
            {
                new Source(new WmlDocument(source1.FullName), false),
                new Source(new WmlDocument(source2.FullName), true),
            };
            var out4 = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB006-Out4.docx"));
            DocumentBuilder.BuildDocument(sources, out4.FullName);
            Validate(out4);

            // Create a new document that consists of the first 5 paragraphs of Source1.docx and the first
            // five paragraphs of Source2.docx.  This example returns a new WmlDocument, when you then can
            // serialize to a SharePoint document library, or use in some other interesting scenario.
            sources = new List<Source>()
            {
                new Source(new WmlDocument(source1.FullName), 0, 5, false),
                new Source(new WmlDocument(source2.FullName), 0, 5, true),
            };
            WmlDocument wmlOut5 = DocumentBuilder.BuildDocument(sources);
            var out5 = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB006-Out5.docx"));
            
            wmlOut5.SaveAs(out5.FullName);  // save it to the file system, but we could just as easily done something
                                            // else with it.
            Validate(out5);
        }

        [Fact]
        public void DB007_Example_DocumentBuilder02_WhitePaper()
        {
            FileInfo spec = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB007-Spec.docx"));
            FileInfo whitePaper = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB007-WhitePaper.docx"));
            FileInfo paperAbstract = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB007-Abstract.docx"));
            FileInfo authorBio = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB007-AuthorBiography.docx"));

            List<Source> sources = null;
            sources = new List<Source>()
            {
                new Source(new WmlDocument(whitePaper.FullName), 0, 1, true),
                new Source(new WmlDocument(paperAbstract.FullName), false),
                new Source(new WmlDocument(authorBio.FullName), false),
                new Source(new WmlDocument(whitePaper.FullName), 1, false),
            };
            var assembledPaper = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB007-AssembledPaper.docx"));
            DocumentBuilder.BuildDocument(sources, assembledPaper.FullName);
            Validate(assembledPaper);
        }

        [Fact]
        public void DB008_DeleteParasWithGivenStyle()
        {
            FileInfo notes = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB007-Notes.docx"));

            List<Source> sources = null;
            // Delete all paragraphs with a specific style.
            using (WordprocessingDocument doc = WordprocessingDocument.Open(notes.FullName, false))
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
                        new WmlDocument(notes.FullName), g.First().Index,
                            g.Last().Index - g.First().Index + 1, true))
                    .ToList();
            }
            var newNotes = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB008-NewNotes.docx"));
            DocumentBuilder.BuildDocument(sources, newNotes.FullName);
            Validate(newNotes);
        }

        private class DocumentInfo
        {
            public int DocumentNumber;
            public int Start;
            public int Count;
        }

        [Fact]
        public void DB009_ShredDocument()
        {
            FileInfo spec = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB007-Spec.docx"));
            // Shred a document into multiple parts for each section
            List<DocumentInfo> documentList;
            using (WordprocessingDocument doc = WordprocessingDocument.Open(spec.FullName, false))
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
                string fileName = String.Format("DB009-Section{0:000}.docx", doc.DocumentNumber);
                var fiSection = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, fileName));
                List<Source> documentSource = new List<Source> {
                    new Source(new WmlDocument(spec.FullName), doc.Start, doc.Count, true)
                };
                DocumentBuilder.BuildDocument(documentSource, fiSection.FullName);
                Validate(fiSection);
            }

            // Re-assemble the parts into a single document.
            List<Source> sources = TestUtil.TempDir
                .GetFiles("DB009-Section*.docx")
                .Select(d => new Source(new WmlDocument(d.FullName), true))
                .ToList();
            var fiReassembled = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB009-Reassembled.docx"));

            DocumentBuilder.BuildDocument(sources, fiReassembled.FullName);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(fiReassembled.FullName, true))
            {
                ReferenceAdder.AddToc(doc, "/w:document/w:body/w:p[1]",
                    @"TOC \o '1-3' \h \z \u", null, null);
            }
            Validate(fiReassembled);
        }


        [Fact]
        public void DB010_InsertUsingInsertId()
        {
            FileInfo front = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB010-FrontMatter.docx"));
            FileInfo insert01 = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB010-Insert-01.docx"));
            FileInfo insert02 = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB010-Insert-02.docx"));
            FileInfo template = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB010-Template.docx"));

            WmlDocument doc1 = new WmlDocument(template.FullName);
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(doc1.DocumentByteArray, 0, doc1.DocumentByteArray.Length);
                using (WordprocessingDocument doc = WordprocessingDocument.Open(mem, true))
                {
                    XDocument xDoc = doc.MainDocumentPart.GetXDocument();
                    XElement frontMatterPara = xDoc.Root.Descendants(W.txbxContent).Elements(W.p).FirstOrDefault();
                    frontMatterPara.ReplaceWith(
                        new XElement(PtOpenXml.Insert,
                            new XAttribute("Id", "Front")));
                    XElement tbl = xDoc.Root.Element(W.body).Elements(W.tbl).FirstOrDefault();
                    XElement firstCell = tbl.Descendants(W.tr).First().Descendants(W.p).First();
                    firstCell.ReplaceWith(
                        new XElement(PtOpenXml.Insert,
                            new XAttribute("Id", "Liz")));
                    XElement secondCell = tbl.Descendants(W.tr).Skip(1).First().Descendants(W.p).First();
                    secondCell.ReplaceWith(
                        new XElement(PtOpenXml.Insert,
                            new XAttribute("Id", "Eric")));
                    doc.MainDocumentPart.PutXDocument();
                }
                doc1.DocumentByteArray = mem.ToArray();
            }

            List<Source> sources = new List<Source>()
            {
                new Source(doc1, true),
                new Source(new WmlDocument(insert01.FullName), "Liz"),
                new Source(new WmlDocument(insert02.FullName), "Eric"),
                new Source(new WmlDocument(front.FullName), "Front"),
            };
            var out1 = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB010-Inserted.docx"));
            DocumentBuilder.BuildDocument(sources, out1.FullName);
            Validate(out1);
        }

        private void Validate(FileInfo fi)
        {
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(fi.FullName, true))
            {
                OpenXmlValidator v = new OpenXmlValidator();
                var errors = v.Validate(wDoc).Where(ve => !s_ExpectedErrors.Contains(ve.Description));

                //if (errors.Count() != 0)
                //{
                //    StringBuilder sb = new StringBuilder();
                //    foreach (var item in errors)
                //    {
                //        sb.Append(item.Description).Append(Environment.NewLine);
                //    }
                //    var s = sb.ToString();
                //    Console.WriteLine(s);
                //}

                Assert.Equal(0, errors.Count());
            }
        }

        private static List<string> s_ExpectedErrors = new List<string>()
        {
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowFirstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowLastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowFirstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowLastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddVBand' attribute is not declared.",
            "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:updateFields'.",
        };


        
    }
}
