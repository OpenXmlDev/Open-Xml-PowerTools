using Codeuctivity.OpenXmlPowerTools;
using Codeuctivity.OpenXmlPowerTools.DocumentBuilder;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace Codeuctivity.Tests
{
    public class DbTests
    {
        [Fact]
        public void DB001_DocumentBuilderKeepSections()
        {
            var name = "DB001-Sections.docx";
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
            var sources = new List<Source>()
            {
                new Source(new WmlDocument(sourceDocx.FullName), true),
            };
            var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-processed-by-DocumentBuilder.docx")));
            DocumentBuilder.BuildDocument(sources, processedDestDocx.FullName);
        }

        [Fact]
        public void DB002_DocumentBuilderKeepSectionsDiscardHeaders()
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1Docx = new FileInfo(Path.Combine(sourceDir.FullName, "DB002-Sections-With-Headers.docx"));
            var source2Docx = new FileInfo(Path.Combine(sourceDir.FullName, "DB002-Landscape-Section.docx"));
            var sources = new List<Source>()
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
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1Docx = new FileInfo(Path.Combine(sourceDir.FullName, "DB003-Only-Default-Header.docx"));
            var source2Docx = new FileInfo(Path.Combine(sourceDir.FullName, "DB002-Landscape-Section.docx"));
            var sources = new List<Source>()
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
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1Docx = new FileInfo(Path.Combine(sourceDir.FullName, "DB004-No-Headers.docx"));
            var source2Docx = new FileInfo(Path.Combine(sourceDir.FullName, "DB002-Landscape-Section.docx"));
            var sources = new List<Source>()
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
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1Docx = new FileInfo(Path.Combine(sourceDir.FullName, "DB005-Headers-With-Images.docx"));
            var source2Docx = new FileInfo(Path.Combine(sourceDir.FullName, "DB002-Landscape-Section.docx"));
            var sources = new List<Source>()
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
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1 = new FileInfo(Path.Combine(sourceDir.FullName, "DB006-Source1.docx"));
            var source2 = new FileInfo(Path.Combine(sourceDir.FullName, "DB006-Source2.docx"));
            var source3 = new FileInfo(Path.Combine(sourceDir.FullName, "DB006-Source3.docx"));

            // Create new document from 10 paragraphs starting at paragraph 5 of Source1.docx
            var sources = new List<Source>()
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
            var wmlOut5 = DocumentBuilder.BuildDocument(sources);
            var out5 = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB006-Out5.docx"));

            wmlOut5.SaveAs(out5.FullName);  // save it to the file system, but we could just as easily done something
                                            // else with it.
            Validate(out5);
        }

        [Fact]
        public void DB007_Example_DocumentBuilder02_WhitePaper()
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var whitePaper = new FileInfo(Path.Combine(sourceDir.FullName, "DB007-WhitePaper.docx"));
            var paperAbstract = new FileInfo(Path.Combine(sourceDir.FullName, "DB007-Abstract.docx"));
            var authorBio = new FileInfo(Path.Combine(sourceDir.FullName, "DB007-AuthorBiography.docx"));
            var sources = new List<Source>()
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
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var notes = new FileInfo(Path.Combine(sourceDir.FullName, "DB007-Notes.docx"));

            List<Source> sources = null;
            // Delete all paragraphs with a specific style.
            using (var doc = WordprocessingDocument.Open(notes.FullName, false))
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

        [Theory]
        [InlineData("DB009-00010", "DB/HeadersFooters/Src/Content-Controls.docx", "DB/HeadersFooters/Dest/Fax.docx", "Templafy")]
        [InlineData("DB009-00020", "DB/HeadersFooters/Src/Letterhead.docx", "DB/HeadersFooters/Dest/Fax.docx", "Templafy")]
        [InlineData("DB009-00030", "DB/HeadersFooters/Src/Letterhead-with-Watermark.docx", "DB/HeadersFooters/Dest/Fax.docx", "Templafy")]
        [InlineData("DB009-00040", "DB/HeadersFooters/Src/Logo.docx", "DB/HeadersFooters/Dest/Fax.docx", "Templafy")]
        [InlineData("DB009-00050", "DB/HeadersFooters/Src/Watermark-1.docx", "DB/HeadersFooters/Dest/Fax.docx", "Templafy")]
        [InlineData("DB009-00060", "DB/HeadersFooters/Src/Watermark-2.docx", "DB/HeadersFooters/Dest/Fax.docx", "Templafy")]
        [InlineData("DB009-00070", "DB/HeadersFooters/Src/Disclaimer.docx", "DB/HeadersFooters/Dest/Fax.docx", "Templafy")]
        [InlineData("DB009-00080", "DB/HeadersFooters/Src/Footer.docx", "DB/HeadersFooters/Dest/Fax.docx", "Templafy")]
        [InlineData("DB009-00110", "DB/HeadersFooters/Src/Content-Controls.docx", "DB/HeadersFooters/Dest/Letter.docx", "Templafy")]
        [InlineData("DB009-00120", "DB/HeadersFooters/Src/Letterhead.docx", "DB/HeadersFooters/Dest/Letter.docx", "Templafy")]
        [InlineData("DB009-00130", "DB/HeadersFooters/Src/Letterhead-with-Watermark.docx", "DB/HeadersFooters/Dest/Letter.docx", "Templafy")]
        [InlineData("DB009-00140", "DB/HeadersFooters/Src/Logo.docx", "DB/HeadersFooters/Dest/Letter.docx", "Templafy")]
        [InlineData("DB009-00150", "DB/HeadersFooters/Src/Watermark-1.docx", "DB/HeadersFooters/Dest/Letter.docx", "Templafy")]
        [InlineData("DB009-00160", "DB/HeadersFooters/Src/Watermark-2.docx", "DB/HeadersFooters/Dest/Letter.docx", "Templafy")]
        [InlineData("DB009-00170", "DB/HeadersFooters/Src/Disclaimer.docx", "DB/HeadersFooters/Dest/Letter.docx", "Templafy")]
        [InlineData("DB009-00180", "DB/HeadersFooters/Src/Footer.docx", "DB/HeadersFooters/Dest/Letter.docx", "Templafy")]
        public void DB009_ImportIntoHeadersFooters(string testId, string src, string dest, string insertId)
        {
            // Load the source document
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceDocxFi = new FileInfo(Path.Combine(sourceDir.FullName, src));
            var wmlSourceDocument = new WmlDocument(sourceDocxFi.FullName);

            // Load the dest document
            var destDocxFi = new FileInfo(Path.Combine(sourceDir.FullName, dest));
            var wmlDestDocument = new WmlDocument(destDocxFi.FullName);

            // Create the dir for the test
            var rootTempDir = TestUtil.TempDir;
            var thisTestTempDir = new DirectoryInfo(Path.Combine(rootTempDir.FullName, testId));
            if (thisTestTempDir.Exists)
            {
                Assert.True(false, "Duplicate test id: " + testId);
            }
            else
            {
                thisTestTempDir.Create();
            }

            var tempDirFullName = thisTestTempDir.FullName;

            // Copy src DOCX to temp directory, for ease of review

            while (true)
            {
                try
                {
                    var sourceDocxCopiedToDestFileName = new FileInfo(Path.Combine(tempDirFullName, sourceDocxFi.Name));
                    if (!sourceDocxCopiedToDestFileName.Exists)
                    {
                        wmlSourceDocument.SaveAs(sourceDocxCopiedToDestFileName.FullName);
                    }
                    break;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                }
            }

            // Copy dest DOCX to temp directory, for ease of review

            while (true)
            {
                try
                {
                    var destDocxCopiedToDestFileName = new FileInfo(Path.Combine(tempDirFullName, destDocxFi.Name));
                    if (!destDocxCopiedToDestFileName.Exists)
                    {
                        wmlDestDocument.SaveAs(destDocxCopiedToDestFileName.FullName);
                    }
                    break;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                }
            }

            var sources = new List<Source>()
            {
                new Source(wmlDestDocument),
                new Source(wmlSourceDocument, insertId),
            };

            var outFi = new FileInfo(Path.Combine(tempDirFullName, "Output.docx"));
            DocumentBuilder.BuildDocument(sources, outFi.FullName);
            Validate(outFi);
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
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var spec = new FileInfo(Path.Combine(sourceDir.FullName, "DB007-Spec.docx"));
            // Shred a document into multiple parts for each section
            List<DocumentInfo> documentList;
            using (var doc = WordprocessingDocument.Open(spec.FullName, false))
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
                var zipped = beforeZipped.PtZip(sectionCounts, (pi, sc) => new
                {
                    pi.Paragraph,
                    pi.Index,
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
                var fileName = string.Format("DB009-Section{0:000}.docx", doc.DocumentNumber);
                var fiSection = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, fileName));
                var documentSource = new List<Source> {
                    new Source(new WmlDocument(spec.FullName), doc.Start, doc.Count, true)
                };
                DocumentBuilder.BuildDocument(documentSource, fiSection.FullName);
                Validate(fiSection);
            }

            // Re-assemble the parts into a single document.
            var sources = TestUtil.TempDir
                .GetFiles("DB009-Section*.docx")
                .Select(d => new Source(new WmlDocument(d.FullName), true))
                .ToList();
            var fiReassembled = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB009-Reassembled.docx"));

            DocumentBuilder.BuildDocument(sources, fiReassembled.FullName);
            using (var doc = WordprocessingDocument.Open(fiReassembled.FullName, true))
            {
                ReferenceAdder.AddToc(doc, "/w:document/w:body/w:p[1]",
                    @"TOC \o '1-3' \h \z \u", null, null);
            }
            Validate(fiReassembled);
        }

        [Fact]
        public void DB010_InsertUsingInsertId()
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var front = new FileInfo(Path.Combine(sourceDir.FullName, "DB010-FrontMatter.docx"));
            var insert01 = new FileInfo(Path.Combine(sourceDir.FullName, "DB010-Insert-01.docx"));
            var insert02 = new FileInfo(Path.Combine(sourceDir.FullName, "DB010-Insert-02.docx"));
            var template = new FileInfo(Path.Combine(sourceDir.FullName, "DB010-Template.docx"));

            var doc1 = new WmlDocument(template.FullName);
            using (var mem = new MemoryStream())
            {
                mem.Write(doc1.DocumentByteArray, 0, doc1.DocumentByteArray.Length);
                using (var doc = WordprocessingDocument.Open(mem, true))
                {
                    var xDoc = doc.MainDocumentPart.GetXDocument();
                    var frontMatterPara = xDoc.Root.Descendants(W.txbxContent).Elements(W.p).FirstOrDefault();
                    frontMatterPara.ReplaceWith(
                        new XElement(PtOpenXml.Insert,
                            new XAttribute("Id", "Front")));
                    var tbl = xDoc.Root.Element(W.body).Elements(W.tbl).FirstOrDefault();
                    var firstCell = tbl.Descendants(W.tr).First().Descendants(W.p).First();
                    firstCell.ReplaceWith(
                        new XElement(PtOpenXml.Insert,
                            new XAttribute("Id", "Liz")));
                    var secondCell = tbl.Descendants(W.tr).Skip(1).First().Descendants(W.p).First();
                    secondCell.ReplaceWith(
                        new XElement(PtOpenXml.Insert,
                            new XAttribute("Id", "Eric")));
                    doc.MainDocumentPart.PutXDocument();
                }
                doc1.DocumentByteArray = mem.ToArray();
            }

            var sources = new List<Source>()
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

        [Fact]
        public void DB011_BodyAndHeaderWithShapes()
        {
            // Both of theses documents have a shape with a DocProperties ID of 1.
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1 = new FileInfo(Path.Combine(sourceDir.FullName, "DB011-Header-With-Shape.docx"));
            var source2 = new FileInfo(Path.Combine(sourceDir.FullName, "DB011-Body-With-Shape.docx"));
            var sources = new List<Source>()
            {
                new Source(new WmlDocument(source1.FullName)),
                new Source(new WmlDocument(source2.FullName)),
            };
            var processedDestDocx =
                new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB011-Body-And-Header-With-Shapes.docx"));
            DocumentBuilder.BuildDocument(sources, processedDestDocx.FullName);
            Validate(processedDestDocx);

            ValidateUniqueDocPrIds(processedDestDocx);
        }

        [Fact]
        public void DB012_NumberingsWithSameAbstractNumbering()
        {
            // This document has three numbering definitions that use the same abstract numbering definition.
            var name = "DB012-Lists-With-Different-Numberings.docx";
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
            var sources = new List<Source>()
            {
                new Source(new WmlDocument(sourceDocx.FullName)),
            };
            var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName,
                sourceDocx.Name.Replace(".docx", "-processed-by-DocumentBuilder.docx")));
            DocumentBuilder.BuildDocument(sources, processedDestDocx.FullName);

            using var wDoc = WordprocessingDocument.Open(processedDestDocx.FullName, false);
            var numberingRoot = wDoc.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root;
            Assert.Equal(3, numberingRoot.Elements(W.num).Count());
        }

        [Fact]
        public void DB013a_LocalizedStyleIds_Heading()
        {
            // Each of these documents have changed the font color of the Heading 1 style, one to red, the other to green.
            // One of the documents were created with English as the Word display language, the other with Danish as the language.
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1 =
                new FileInfo(Path.Combine(sourceDir.FullName, "DB013a-Red-Heading1-English.docx"));
            var source2 = new FileInfo(Path.Combine(sourceDir.FullName,
                "DB013a-Green-Heading1-Danish.docx"));
            List<Source> sources = null;

            sources = new List<Source>()
            {
                new Source(new WmlDocument(source1.FullName)),
                new Source(new WmlDocument(source2.FullName)),
            };
            var processedDestDocx =
                new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB013a-Colored-Heading1.docx"));
            DocumentBuilder.BuildDocument(sources, processedDestDocx.FullName);

            using var wDoc = WordprocessingDocument.Open(processedDestDocx.FullName, false);
            var styles = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument().Root.Elements(W.style).ToArray();
            Assert.Equal(1, styles.Count(s => s.Element(W.name).Attribute(W.val).Value == "heading 1"));

            var styleIds = new HashSet<string>(styles.Select(s => s.Attribute(W.styleId).Value));
            var paragraphStylesIds = new HashSet<string>(wDoc.MainDocumentPart.GetXDocument()
                .Descendants(W.pStyle)
                .Select(p => p.Attribute(W.val).Value));
            Assert.Subset(styleIds, paragraphStylesIds);
        }

        [Fact]
        public void DB013b_LocalizedStyleIds_List()
        {
            // Each of these documents have changed the font color of the List Paragraph style, one to orange, the other to blue.
            // One of the documents were created with English as the Word display language, the other with Danish as the language.
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1 =
                new FileInfo(Path.Combine(sourceDir.FullName, "DB013b-Orange-List-Danish.docx"));
            var source2 = new FileInfo(Path.Combine(sourceDir.FullName,
                "DB013b-Blue-List-English.docx"));
            List<Source> sources = null;

            sources = new List<Source>()
            {
                new Source(new WmlDocument(source1.FullName)),
                new Source(new WmlDocument(source2.FullName)),
            };
            var processedDestDocx =
                new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB013b-Colored-List.docx"));
            DocumentBuilder.BuildDocument(sources, processedDestDocx.FullName);

            using var wDoc = WordprocessingDocument.Open(processedDestDocx.FullName, false);
            var styles = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument().Root.Elements(W.style).ToArray();
            Assert.Equal(1, styles.Count(s => s.Element(W.name).Attribute(W.val).Value == "List Paragraph"));

            var styleIds = new HashSet<string>(styles.Select(s => s.Attribute(W.styleId).Value));
            var paragraphStylesIds = new HashSet<string>(wDoc.MainDocumentPart.GetXDocument()
                .Descendants(W.pStyle)
                .Select(p => p.Attribute(W.val).Value));
            Assert.Subset(styleIds, paragraphStylesIds);
        }

        [Fact]
        public void DB014_KeepWebExtensions()
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source = new FileInfo(Path.Combine(sourceDir.FullName, "DB014-WebExtensions.docx"));
            var sources = new List<Source>()
            {
                new Source(new WmlDocument(source.FullName)),
            };
            var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB014-WebExtensions.docx"));
            DocumentBuilder.BuildDocument(sources, processedDestDocx.FullName);
            Validate(processedDestDocx);

            using var wDoc = WordprocessingDocument.Open(processedDestDocx.FullName, false);
            Assert.NotNull(wDoc.WebExTaskpanesPart);
            Assert.Equal(2, wDoc.WebExTaskpanesPart.Taskpanes.ChildElements.Count);
            Assert.Equal(2, wDoc.WebExTaskpanesPart.WebExtensionParts.Count());
        }

        [Fact]
        public void DB015_LatentStyles()
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source = new FileInfo(Path.Combine(sourceDir.FullName, "DB015-LatentStyles.docx"));
            var sources = new List<Source>()
            {
                new Source(new WmlDocument(source.FullName)),
            };
            var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB015-LatentStyles.docx"));
            DocumentBuilder.BuildDocument(sources, processedDestDocx.FullName);
            Validate(processedDestDocx);

            //using (WordprocessingDocument wDoc = WordprocessingDocument.Open(processedDestDocx.FullName, false))
            //{
            //    Assert.NotNull(wDoc.WebExTaskpanesPart);
            //    Assert.Equal(2, wDoc.WebExTaskpanesPart.Taskpanes.ChildElements.Count);
            //    Assert.Equal(2, wDoc.WebExTaskpanesPart.WebExtensionParts.Count());
            //}
        }

        [Fact]
        public void DB0016_DocDefaultStyles()
        {
            var name = "DB0016-DocDefaultStyles.docx";
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
            var sources = new List<Source>()
            {
                new Source(new WmlDocument(sourceDocx.FullName), true),
            };
            var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName,
                sourceDocx.Name.Replace(".docx", "-processed-by-DocumentBuilder.docx")));
            DocumentBuilder.BuildDocument(sources, processedDestDocx.FullName);

            using var wDoc = WordprocessingDocument.Open(processedDestDocx.FullName, true);
            var styles = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument().Root.Elements(W.docDefaults).ToArray();
            Assert.Single(styles);
        }

        [Fact]
        public void DB017_ApplyHeaderAndFooterToAllDocs()
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1 = new FileInfo(Path.Combine(sourceDir.FullName, "DB017-ApplyHeaderAndFooterToAllDocs-Portrait-TwoColumns.docx"));
            var source2 = new FileInfo(Path.Combine(sourceDir.FullName, "DB017-ApplyHeaderAndFooterToAllDocs-Landscape-SingleColumn.docx"));

            var sources = new List<Source>()
            {
                new Source(new WmlDocument(source1.FullName)){KeepSections = true},
                new Source(new WmlDocument(source2.FullName)){KeepSections = true, DiscardHeadersAndFootersInKeptSections = true},
            };

            var processedDestDocx = new FileInfo(Path.Combine(Path.Combine(TestUtil.TempDir.FullName), "DB017-ApplyHeaderAndFooterToAllDocs.docx"));
            DocumentBuilder.BuildDocument(sources, processedDestDocx.FullName);
        }

        [Theory]
        [InlineData("DB100-00010", "DB/GlossaryDocuments/CellLevelContentControl-built.docx", "DB/GlossaryDocuments/BaseDocument.docx,0,4", "DB/GlossaryDocuments/CellLevelContentControl.docx", "DB/GlossaryDocuments/BaseDocument.docx,4", null, null, null)]
        [InlineData("DB100-00020", "DB/GlossaryDocuments/InlineContentControl-built.docx", "DB/GlossaryDocuments/BaseDocument.docx,0,4", "DB/GlossaryDocuments/InlineContentControl.docx", "DB/GlossaryDocuments/BaseDocument.docx,4", null, null, null)]
        [InlineData("DB100-00030", "DB/GlossaryDocuments/MultilineWithBulletPoints-built.docx", "DB/GlossaryDocuments/BaseDocument.docx,0,4", "DB/GlossaryDocuments/MultilineWithBulletPoints.docx", "DB/GlossaryDocuments/BaseDocument.docx,4", null, null, null)]
        [InlineData("DB100-00040", "DB/GlossaryDocuments/NestedContentControl-built.docx", "DB/GlossaryDocuments/BaseDocument.docx,0,4", "DB/GlossaryDocuments/NestedContentControl.docx", "DB/GlossaryDocuments/BaseDocument.docx,4", null, null, null)]
        [InlineData("DB100-00050", "DB/GlossaryDocuments/RowLevelContentControl-built.docx", "DB/GlossaryDocuments/BaseDocument.docx,0,4", "DB/GlossaryDocuments/RowLevelContentControl.docx", "DB/GlossaryDocuments/BaseDocument.docx,4", null, null, null)]
        [InlineData("DB100-00060", "DB/GlossaryDocuments/ContentControlDanishProofingLanguage-built.docx", "DB/GlossaryDocuments/BaseDocument.docx,0,4", "DB/GlossaryDocuments/ContentControlDanishProofingLanguage.docx", "DB/GlossaryDocuments/BaseDocument.docx,4", null, null, null)]
        [InlineData("DB100-00070", "DB/GlossaryDocuments/ContentControlEnglishProofingLanguage-built.docx", "DB/GlossaryDocuments/BaseDocument.docx,0,4", "DB/GlossaryDocuments/ContentControlEnglishProofingLanguage.docx", "DB/GlossaryDocuments/BaseDocument.docx,4", null, null, null)]
        [InlineData("DB100-00080", "DB/GlossaryDocuments/ContentControlMixedProofingLanguage-built.docx", "DB/GlossaryDocuments/BaseDocument.docx,0,4", "DB/GlossaryDocuments/ContentControlMixedProofingLanguage.docx", "DB/GlossaryDocuments/BaseDocument.docx,4", null, null, null)]
        [InlineData("DB100-00090", "DB/GlossaryDocuments/ContentControlWithContent-built.docx", "DB/GlossaryDocuments/BaseDocument.docx,0,4", "DB/GlossaryDocuments/ContentControlWithContent.docx", "DB/GlossaryDocuments/BaseDocument.docx,4", null, null, null)]
        [InlineData("DB100-00100", "DB/GlossaryDocuments/FooterContent-built.docx", "DB/GlossaryDocuments/BaseDocument.docx,0,4", "DB/GlossaryDocuments/FooterContent.docx", "DB/GlossaryDocuments/BaseDocument.docx,4", null, null, null)]
        [InlineData("DB100-00110", "DB/GlossaryDocuments/HeaderContent-built.docx", "DB/GlossaryDocuments/BaseDocument.docx,0,4", "DB/GlossaryDocuments/HeaderContent.docx", "DB/GlossaryDocuments/BaseDocument.docx,4", null, null, null)]
        [InlineData("DB100-00200", null, "DB/GlossaryDocuments/BaseDocument.docx", "DB/GlossaryDocuments/CellLevelContentControl.docx", "DB/GlossaryDocuments/NestedContentControl.docx", null, null, null)]
        public void WithGlossaryDocuments(string testId, string baseline, string src1, string src2, string src3, string src4, string src5, string src6)
        {
            var rawSources = new string[] { src1, src2, src3, src4, src5, src6, };
            var sourcesStr = rawSources.Where(s => s != null).ToArray();

            // Load the source documents
            var sources = sourcesStr.Select(s =>
            {
                var spl = s.Split(',');
                if (spl.Length == 1)
                {
                    var sourceDir = new DirectoryInfo("../../../../TestFiles/");
                    var sourceFi = new FileInfo(Path.Combine(sourceDir.FullName, s));
                    var wmlSource = new WmlDocument(sourceFi.FullName);
                    return new Source(wmlSource);
                }
                else if (spl.Length == 2)
                {
                    var start = int.Parse(spl[1]);
                    var sourceDir = new DirectoryInfo("../../../../TestFiles/");
                    var sourceFi = new FileInfo(Path.Combine(sourceDir.FullName, spl[0]));
                    return new Source(sourceFi.FullName, start, true);
                }
                else
                {
                    var start = int.Parse(spl[1]);
                    var count = int.Parse(spl[2]);
                    var sourceDir = new DirectoryInfo("../../../../TestFiles/");
                    var sourceFi = new FileInfo(Path.Combine(sourceDir.FullName, spl[0]));
                    return new Source(sourceFi.FullName, start, count, true);
                }
            })
                .ToList();

            // Create the dir for the test
            var rootTempDir = TestUtil.TempDir;
            var thisTestTempDir = new DirectoryInfo(Path.Combine(rootTempDir.FullName, testId));
            if (thisTestTempDir.Exists)
            {
                Assert.True(false, "Duplicate test id: " + testId);
            }
            else
            {
                thisTestTempDir.Create();
            }

            var tempDirFullName = thisTestTempDir.FullName;

            // Copy sources to temp directory, for ease of review

            foreach (var item in sources)
            {
                var fi = new FileInfo(item.WmlDocument.FileName);
                var sourceCopiedToDestFi = new FileInfo(Path.Combine(tempDirFullName, fi.Name));
                if (!sourceCopiedToDestFi.Exists)
                {
                    File.Copy(item.WmlDocument.FileName, sourceCopiedToDestFi.FullName);
                }
            }

            if (baseline != null)
            {
                var sourceDir = new DirectoryInfo("../../../../TestFiles/");
                var baselineFi = new FileInfo(Path.Combine(sourceDir.FullName, baseline));
                var baselineCopiedToDestFileName = new FileInfo(Path.Combine(tempDirFullName, baselineFi.Name));
                File.Copy(baselineFi.FullName, baselineCopiedToDestFileName.FullName);
            }

            // Use DocumentBuilder to build the destination document

            var outFi = new FileInfo(Path.Combine(tempDirFullName, "Output.docx"));
            var settings = new DocumentBuilderSettings();
            DocumentBuilder.BuildDocument(sources, outFi.FullName, settings);
            Validate(outFi);
        }

        private void Validate(FileInfo fi)
        {
            using var wDoc = WordprocessingDocument.Open(fi.FullName, true);
            var v = new OpenXmlValidator();
            var errors = v.Validate(wDoc).Where(ve =>
            {
                var found = s_ExpectedErrors.Any(xe => ve.Description.Contains(xe));
                return !found;
            });

            if (errors.Any())
            {
                var sb = new StringBuilder();
                foreach (var item in errors)
                {
                    sb.Append(item.Description).Append(Environment.NewLine);
                }
                var s = sb.ToString();
                Assert.True(false, s);
            }
        }

        private static readonly List<string> s_ExpectedErrors = new List<string>()
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
            "The attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:name' has invalid value 'useWord2013TrackBottomHyphenation'. The Enumeration constraint failed.",
            "The 'http://schemas.microsoft.com/office/word/2012/wordml:restartNumberingAfterBreak' attribute is not declared.",
            "Attribute 'id' should have unique value. Its current value '",
            "The 'urn:schemas-microsoft-com:mac:vml:blur' attribute is not declared.",
            "Attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:id' should have unique value. Its current value '",
            "The element has unexpected child element 'http://schemas.microsoft.com/office/word/2012/wordml:",
            "The element has invalid child element 'http://schemas.microsoft.com/office/word/2012/wordml:",
            "The 'urn:schemas-microsoft-com:mac:vml:complextextbox' attribute is not declared.",
            "http://schemas.microsoft.com/office/word/2010/wordml:",
            "http://schemas.microsoft.com/office/word/2008/9/12/wordml:",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:allStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:customStyles' attribute is not declared.",
        };

        private void ValidateUniqueDocPrIds(FileInfo fi)
        {
            using var doc = WordprocessingDocument.Open(fi.FullName, false);
            var docPrIds = new HashSet<string>();
            foreach (var item in doc.MainDocumentPart.GetXDocument().Descendants(WP.docPr))
            {
                Assert.True(docPrIds.Add(item.Attribute(NoNamespace.id).Value));
            }

            foreach (var header in doc.MainDocumentPart.HeaderParts)
            {
                foreach (var item in header.GetXDocument().Descendants(WP.docPr))
                {
                    Assert.True(docPrIds.Add(item.Attribute(NoNamespace.id).Value));
                }
            }

            foreach (var footer in doc.MainDocumentPart.FooterParts)
            {
                foreach (var item in footer.GetXDocument().Descendants(WP.docPr))
                {
                    Assert.True(docPrIds.Add(item.Attribute(NoNamespace.id).Value));
                }
            }

            if (doc.MainDocumentPart.FootnotesPart != null)
            {
                foreach (var item in doc.MainDocumentPart.FootnotesPart.GetXDocument().Descendants(WP.docPr))
                {
                    Assert.True(docPrIds.Add(item.Attribute(NoNamespace.id).Value));
                }
            }

            if (doc.MainDocumentPart.EndnotesPart != null)
            {
                foreach (var item in doc.MainDocumentPart.EndnotesPart.GetXDocument().Descendants(WP.docPr))
                {
                    Assert.True(docPrIds.Add(item.Attribute(NoNamespace.id).Value));
                }
            }
        }
    }
}