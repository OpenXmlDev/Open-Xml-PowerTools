// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

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
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using Xunit;

#if !ELIDE_XUNIT_TESTS

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
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Load the source document
            FileInfo sourceDocxFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, src));
            WmlDocument wmlSourceDocument = new WmlDocument(sourceDocxFi.FullName);

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Load the dest document
            FileInfo destDocxFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, dest));
            WmlDocument wmlDestDocument = new WmlDocument(destDocxFi.FullName);

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create the dir for the test
            var rootTempDir = TestUtil.TempDir;
            var thisTestTempDir = new DirectoryInfo(Path.Combine(rootTempDir.FullName, testId));
            if (thisTestTempDir.Exists)
                Assert.True(false, "Duplicate test id: " + testId);
            else
                thisTestTempDir.Create();
            var tempDirFullName = thisTestTempDir.FullName;

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Copy src DOCX to temp directory, for ease of review

            while (true)
            {
                try
                {
                    ////////// CODE TO REPEAT UNTIL SUCCESS //////////
                    var sourceDocxCopiedToDestFileName = new FileInfo(Path.Combine(tempDirFullName, sourceDocxFi.Name));
                    if (!sourceDocxCopiedToDestFileName.Exists)
                        wmlSourceDocument.SaveAs(sourceDocxCopiedToDestFileName.FullName);
                    //////////////////////////////////////////////////
                    break;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                }
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Copy dest DOCX to temp directory, for ease of review

            while (true)
            {
                try
                {
                    ////////// CODE TO REPEAT UNTIL SUCCESS //////////
                    var destDocxCopiedToDestFileName = new FileInfo(Path.Combine(tempDirFullName, destDocxFi.Name));
                    if (!destDocxCopiedToDestFileName.Exists)
                        wmlDestDocument.SaveAs(destDocxCopiedToDestFileName.FullName);
                    //////////////////////////////////////////////////
                    break;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                }
            }

            List<Source> sources = new List<Source>()
            {
                new Source(wmlDestDocument),
                new Source(wmlSourceDocument, insertId),
            };

            var outFi = new FileInfo(Path.Combine(tempDirFullName, "Output.docx"));
            DocumentBuilder.BuildDocument(sources, outFi.FullName);
            Validate(outFi);
        }

#if false
        [Theory]
        [InlineData("DB999-00010", "DBTEMP/03DE57384B87AA6C2A3BDE87DDDD7F880DC55E.docx", true)]
        [InlineData("DB999-00020", "DBTEMP/0D3DEB27ED036116466BED616B2056CDD2783A.docx", false)]
        [InlineData("DB999-00030", "DBTEMP/421628B3F4B03B123CA8EDDA5009E449F5F47D.docx", false)]
        [InlineData("DB999-00040", "DBTEMP/58D4E8661C7F44FE33392B89B0A3CB0AF1684F.docx", false)]
        [InlineData("DB999-00050", "DBTEMP/67EBCA627D6D584CAB3EB1DF2E4C3982023DEE.docx", true)]
        [InlineData("DB999-00060", "DBTEMP/A529643E2FC3E2C682FA86DEE0A1B3064DCEE0.docx", false)]
        [InlineData("DB999-00070", "DBTEMP/E794032F0422B440D3C564F0E09E395519127D.docx", false)]
        [InlineData("DB999-00080", "DBTEMP/1FF1ADF30B24978E9449754459C743D3BC67ED.docx", false)]
        [InlineData("DB999-00090", "DBTEMP/5E685927DA2FECB88DE9CAF0BECEC88BC118A7.docx", false)]
        [InlineData("DB999-00100", "DBTEMP/6427BCF5C18B55D627B95F3E14924050628C5B.docx", false)]
        [InlineData("DB999-00110", "DBTEMP/91691E0D3AB89E9927A2BAC5D385BB6277648F.docx", false)]
        [InlineData("DB999-00120", "DBTEMP/9533BC5710190EA01DA86D29CD06880395C4AF.docx", false)]
        [InlineData("DB999-00130", "DBTEMP/E9CD8C556AA52CA7D31DADB51A201EEF580AA8.docx", false)]
        [InlineData("DB999-00140", "DBTEMP/21D3CE149C30B791F9A8BE092828E1469A9047.docx", false)]
        [InlineData("DB999-00150", "DBTEMP/AC0CB8CE43A7ECAE995BB542D4FB1060FB835B.docx", false)]
        [InlineData("DB999-00160", "DBTEMP/C61F69B52EC8B0E2C784C932B26F3C613AE671.docx", false)]
        [InlineData("DB999-00170", "DBTEMP/1DF04A9130B3EF858ACA6837A706A429904973.dotm", false)]
        [InlineData("DB999-00180", "DBTEMP/6E9F26B708DE6076B2C731B97AAA5288D839AB.docm", false)]
        [InlineData("DB999-00190", "DBTEMP/A6649726EA0BD7545932DDD51403D83E4D5917.docx", false)]
        [InlineData("DB999-00200", "DBTEMP/C8AE8AD0A73F24B7CFCFD11918B337CF2B90C9.docx", false)]
        [InlineData("DB999-00210", "DBTEMP/BC46A7FBB212EFD10878A39D91AE3ECAADDAB0.docx", false)]
        [InlineData("DB999-00220", "DBTEMP/B6F0E938B508676B322C47F3E0E29C8D786DB2.docm", false)]
        [InlineData("DB999-00230", "DBTEMP/D4D8694A51DECA243AF748B3232BE565EEE19D.docx", false)]
        [InlineData("DB999-00240", "DBTEMP/F20B3CE72BF635462E22BA3CA81CA9D57F6FEB.docx", false)]
        [InlineData("DB999-00250", "DBTEMP/74ED106FF88C1B195D97C466E00BECCB636A03.docx", false)]
        [InlineData("DB999-00260", "DBTEMP/4421A4B7B6ECC2813070309AA2D86C4BCA4AEF.docx", false)]
        [InlineData("DB999-00270", "DBTEMP/BC7D91B993807518F3D430B7C6592AFD6BD91C.docx", false)]
        [InlineData("DB999-00280", "DBTEMP/3006E76FE65E8A25A91ED204EEBEE6D6D62A44.docx", false)]
        [InlineData("DB999-00290", "DBTEMP/6254B74778BFFCD1799F4F2B3B01C2025AABB2.docx", false)]
        [InlineData("DB999-00300", "DBTEMP/5AD0A0BD99676B268D8E7C1F69238FB9B6149E.docx", false)]
        [InlineData("DB999-00310", "DBTEMP/2D58495ECCF30ED9507B707C689CA9C9D4B049.docx", false)]

        public void DB999_DocumentBuilder(string testId, string src, bool shouldThrow)
        {
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Load the source document
            FileInfo sourceDocxFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, src));
            WmlDocument wmlSourceDocument = new WmlDocument(sourceDocxFi.FullName);

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create the dir for the test
            var rootTempDir = TestUtil.TempDir;
            var thisTestTempDir = new DirectoryInfo(Path.Combine(rootTempDir.FullName, testId));
            if (thisTestTempDir.Exists)
                Assert.True(false, "Duplicate test id: " + testId);
            else
                thisTestTempDir.Create();
            var tempDirFullName = thisTestTempDir.FullName;

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Copy src DOCX to temp directory, for ease of review

            while (true)
            {
                try
                {
                    ////////// CODE TO REPEAT UNTIL SUCCESS //////////
                    var sourceDocxCopiedToDestFileName = new FileInfo(Path.Combine(tempDirFullName, sourceDocxFi.Name));
                    if (!sourceDocxCopiedToDestFileName.Exists)
                        wmlSourceDocument.SaveAs(sourceDocxCopiedToDestFileName.FullName);
                    //////////////////////////////////////////////////
                    break;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                }
            }

            List<string> expectedErrors;
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(wmlSourceDocument.DocumentByteArray, 0, wmlSourceDocument.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, false))
                {
                    OpenXmlValidator validator = new OpenXmlValidator();
                    expectedErrors = validator.Validate(wDoc)
                        .Select(e => e.Description)
                        .Distinct()
                        .ToList();
                }
            }
            foreach (var item in s_ExpectedErrors)
                expectedErrors.Add(item);

            List<Source> sources = new List<Source>()
            {
                new Source(wmlSourceDocument, true),
            };

            var outFi = new FileInfo(Path.Combine(tempDirFullName, "Output.docx"));

            if (shouldThrow)
            {
                Assert.Throws<DocumentBuilderException>(() => DocumentBuilder.BuildDocument(sources, outFi.FullName));
            }
            else
            {
                var outWml = DocumentBuilder.BuildDocument(sources);
                outWml.SaveAs(outFi.FullName);

                using (MemoryStream ms = new MemoryStream())
                {
                    ms.Write(outWml.DocumentByteArray, 0, outWml.DocumentByteArray.Length);
                    using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, false))
                    {
                        OpenXmlValidator validator = new OpenXmlValidator();
                        var errors = validator.Validate(wDoc).Where(e =>
                        {
                            var str = e.Description;
                            foreach (var ee in expectedErrors)
                            {
                                if (str.Contains(ee))
                                    return false;
                            }
                            return true;
                        });
                        if (errors.Count() != 0)
                        {
                            var message = errors.Select(e => e.Description + Environment.NewLine).StringConcatenate();
                            Assert.True(false, message);
                        }
                    }
                }

            }
        }
#endif

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

        [Fact]
        public void DB011_BodyAndHeaderWithShapes()
        {
            // Both of theses documents have a shape with a DocProperties ID of 1.
            FileInfo source1 = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB011-Header-With-Shape.docx"));
            FileInfo source2 = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB011-Body-With-Shape.docx"));
            List<Source> sources = null;

            sources = new List<Source>()
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
            string name = "DB012-Lists-With-Different-Numberings.docx";
            FileInfo sourceDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            List<Source> sources = null;
            sources = new List<Source>()
            {
                new Source(new WmlDocument(sourceDocx.FullName)),
            };
            var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName,
                sourceDocx.Name.Replace(".docx", "-processed-by-DocumentBuilder.docx")));
            DocumentBuilder.BuildDocument(sources, processedDestDocx.FullName);

            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(processedDestDocx.FullName, false))
            {
                var numberingRoot = wDoc.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root;
                Assert.Equal(3, numberingRoot.Elements(W.num).Count());
            }
        }

        [Fact]
        public void DB013a_LocalizedStyleIds_Heading()
        {
            // Each of these documents have changed the font color of the Heading 1 style, one to red, the other to green.
            // One of the documents were created with English as the Word display language, the other with Danish as the language.
            FileInfo source1 =
                new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB013a-Red-Heading1-English.docx"));
            FileInfo source2 = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName,
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

            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(processedDestDocx.FullName, false))
            {
                var styles = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument().Root.Elements(W.style).ToArray();
                Assert.Equal(1, styles.Count(s => s.Element(W.name).Attribute(W.val).Value == "heading 1"));

                var styleIds = new HashSet<string>(styles.Select(s => s.Attribute(W.styleId).Value));
                var paragraphStylesIds = new HashSet<string>(wDoc.MainDocumentPart.GetXDocument()
                    .Descendants(W.pStyle)
                    .Select(p => p.Attribute(W.val).Value));
                Assert.Subset(styleIds, paragraphStylesIds);
            }
        }

        [Fact]
        public void DB013b_LocalizedStyleIds_List()
        {
            // Each of these documents have changed the font color of the List Paragraph style, one to orange, the other to blue.
            // One of the documents were created with English as the Word display language, the other with Danish as the language.
            FileInfo source1 =
                new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB013b-Orange-List-Danish.docx"));
            FileInfo source2 = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName,
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

            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(processedDestDocx.FullName, false))
            {
                var styles = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument().Root.Elements(W.style).ToArray();
                Assert.Equal(1, styles.Count(s => s.Element(W.name).Attribute(W.val).Value == "List Paragraph"));

                var styleIds = new HashSet<string>(styles.Select(s => s.Attribute(W.styleId).Value));
                var paragraphStylesIds = new HashSet<string>(wDoc.MainDocumentPart.GetXDocument()
                    .Descendants(W.pStyle)
                    .Select(p => p.Attribute(W.val).Value));
                Assert.Subset(styleIds, paragraphStylesIds);
            }
        }

        [Fact]
        public void DB014_KeepWebExtensions()
        {
            FileInfo source = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "DB014-WebExtensions.docx"));
            List<Source> sources = null;

            sources = new List<Source>()
            {
                new Source(new WmlDocument(source.FullName)),
            };
            var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB014-WebExtensions.docx"));
            DocumentBuilder.BuildDocument(sources, processedDestDocx.FullName);
            Validate(processedDestDocx);

            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(processedDestDocx.FullName, false))
            {
                Assert.NotNull(wDoc.WebExTaskpanesPart);
                Assert.Equal(2, wDoc.WebExTaskpanesPart.Taskpanes.ChildElements.Count);
                Assert.Equal(2, wDoc.WebExTaskpanesPart.WebExtensionParts.Count());
            }
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

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Load the source documents
            List<Source> sources = sourcesStr.Select(s =>
            {
                var spl = s.Split(',');
                if (spl.Length == 1)
                {
                    var sourceFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, s));
                    var wmlSource = new WmlDocument(sourceFi.FullName);
                    return new Source(wmlSource);
                }
                else if (spl.Length == 2)
                {
                    var start = int.Parse(spl[1]);
                    var sourceFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, spl[0]));
                    return new Source(sourceFi.FullName, start, true);
                }
                else
                {
                    var start = int.Parse(spl[1]);
                    var count = int.Parse(spl[2]);
                    var sourceFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, spl[0]));
                    return new Source(sourceFi.FullName, start, count, true);
                }
            })
                .ToList();

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create the dir for the test
            var rootTempDir = TestUtil.TempDir;
            var thisTestTempDir = new DirectoryInfo(Path.Combine(rootTempDir.FullName, testId));
            if (thisTestTempDir.Exists)
                Assert.True(false, "Duplicate test id: " + testId);
            else
                thisTestTempDir.Create();
            var tempDirFullName = thisTestTempDir.FullName;

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Copy sources to temp directory, for ease of review

            foreach (var item in sources)
            {
                var fi = new FileInfo(item.WmlDocument.FileName);
                var sourceCopiedToDestFi = new FileInfo(Path.Combine(tempDirFullName, fi.Name));
                if (!sourceCopiedToDestFi.Exists)
                    File.Copy(item.WmlDocument.FileName, sourceCopiedToDestFi.FullName);
            }

            if (baseline != null)
            {
                var baselineFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, baseline));
                var baselineCopiedToDestFileName = new FileInfo(Path.Combine(tempDirFullName, baselineFi.Name));
                File.Copy(baselineFi.FullName, baselineCopiedToDestFileName.FullName);
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Use DocumentBuilder to build the destination document

            var outFi = new FileInfo(Path.Combine(tempDirFullName, "Output.docx"));
            DocumentBuilderSettings settings = new DocumentBuilderSettings();
            DocumentBuilder.BuildDocument(sources, outFi.FullName, settings);
            Validate(outFi);
        }

        private void Validate(FileInfo fi)
        {
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(fi.FullName, true))
            {
                OpenXmlValidator v = new OpenXmlValidator();
                var errors = v.Validate(wDoc).Where(ve =>
                {
                    var found = s_ExpectedErrors.Any(xe => ve.Description.Contains(xe));
                    return !found;
                });

                if (errors.Count() != 0)
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (var item in errors)
                    {
                        sb.Append(item.Description).Append(Environment.NewLine);
                    }
                    var s = sb.ToString();
                    Assert.True(false, s);
                }
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
            using (WordprocessingDocument doc = WordprocessingDocument.Open(fi.FullName, false))
            {
                var docPrIds = new HashSet<string>();
                foreach (var item in doc.MainDocumentPart.GetXDocument().Descendants(WP.docPr))
                    Assert.True(docPrIds.Add(item.Attribute(NoNamespace.id).Value));
                foreach (var header in doc.MainDocumentPart.HeaderParts)
                foreach (var item in header.GetXDocument().Descendants(WP.docPr))
                    Assert.True(docPrIds.Add(item.Attribute(NoNamespace.id).Value));
                foreach (var footer in doc.MainDocumentPart.FooterParts)
                foreach (var item in footer.GetXDocument().Descendants(WP.docPr))
                    Assert.True(docPrIds.Add(item.Attribute(NoNamespace.id).Value));
                if (doc.MainDocumentPart.FootnotesPart != null)
                    foreach (var item in doc.MainDocumentPart.FootnotesPart.GetXDocument().Descendants(WP.docPr))
                        Assert.True(docPrIds.Add(item.Attribute(NoNamespace.id).Value));
                if (doc.MainDocumentPart.EndnotesPart != null)
                    foreach (var item in doc.MainDocumentPart.EndnotesPart.GetXDocument().Descendants(WP.docPr))
                        Assert.True(docPrIds.Add(item.Attribute(NoNamespace.id).Value));
            }
        }
    }
}
#endif
