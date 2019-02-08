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
using OpenXmlPowerTools;
using Xunit;
using System.Diagnostics;

/****************************************************************************************************************/
// Large tests have been commented out below.  If and when there is an effort to improve performance for WmlComparer,
// then uncomment.  Performance isn't bad, but certainly is possible to improve.
/****************************************************************************************************************/

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class WcTests
    {
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static bool s_OpenWord = false;
        public static bool m_OpenTempDirInExplorer = false;
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        [Theory]
        [InlineData("RC-0010", "RC/RC001-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC/RC001-After1.docx</DocName>
                <Color>LightYellow</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
              <RcInfo>
                <DocName>RC/RC001-After2.docx</DocName>
                <Color>LightPink</Color>
                <Revisor>From Fred</Revisor>
              </RcInfo>
            </Root>")]
        [InlineData("RC-0020", "RC/RC002-Image.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC/RC002-Image-After1.docx</DocName>
                <Color>LightBlue</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [InlineData("RC-0030", "RC/RC002-Image-After1.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC/RC002-Image.docx</DocName>
                <Color>LightBlue</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [InlineData("RC-0040", "WC/WC027-Twenty-Paras-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>WC/WC027-Twenty-Paras-After-1.docx</DocName>
                <Color>LightBlue</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [InlineData("RC-0050", "WC/WC027-Twenty-Paras-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>WC/WC027-Twenty-Paras-After-3.docx</DocName>
                <Color>LightBlue</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [InlineData("RC-0060", "RC/RC003-Multi-Paras.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC/RC003-Multi-Paras-After.docx</DocName>
                <Color>LightBlue</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [InlineData("RC-0070", "RC/RC004-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC/RC004-After1.docx</DocName>
                <Color>LightYellow</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
              <RcInfo>
                <DocName>RC/RC004-After2.docx</DocName>
                <Color>LightPink</Color>
                <Revisor>From Fred</Revisor>
              </RcInfo>
            </Root>")]
        [InlineData("RC-0080", "RC/RC005-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC/RC005-After1.docx</DocName>
                <Color>LightYellow</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [InlineData("RC-0090", "RC/RC006-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC/RC006-After1.docx</DocName>
                <Color>LightYellow</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [InlineData("RC-0100", "RC/RC007-Endnotes-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC/RC007-Endnotes-After.docx</DocName>
                <Color>LightYellow</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]

        public void WC001_Consolidate(string testId, string originalName, string revisedDocumentsXml)
        {
            FileInfo originalDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, originalName));

            var rootTempDir = TestUtil.TempDir;
            var thisTestTempDir = new DirectoryInfo(Path.Combine(rootTempDir.FullName, testId));
            if (thisTestTempDir.Exists)
                Assert.True(false, "Duplicate test id: " + testId);
            else
                thisTestTempDir.Create();

            var originalCopiedToDestDocx = new FileInfo(Path.Combine(thisTestTempDir.FullName, originalDocx.Name));
            if (!originalCopiedToDestDocx.Exists)
            {
                var wml1 = new WmlDocument(originalDocx.FullName);
                var wml2 = WordprocessingMLUtil.BreakLinkToTemplate(wml1);
                wml2.SaveAs(originalCopiedToDestDocx.FullName);
            }

            var revisedDocumentsXElement = XElement.Parse(revisedDocumentsXml);
            var revisedDocumentsArray = revisedDocumentsXElement
                .Elements()
                .Select(z =>
                {
                    FileInfo revisedDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, z.Element("DocName").Value));
                    var revisedCopiedToDestDocx = new FileInfo(Path.Combine(thisTestTempDir.FullName, revisedDocx.Name));
                    var wml1 = new WmlDocument(revisedDocx.FullName);
                    var wml2 = WordprocessingMLUtil.BreakLinkToTemplate(wml1);
                    wml2.SaveAs(revisedCopiedToDestDocx.FullName);
                    return new WmlRevisedDocumentInfo()
                    {
                        RevisedDocument = new WmlDocument(revisedCopiedToDestDocx.FullName),
                        Color = ColorParser.FromName(z.Element("Color")?.Value),
                        Revisor = z.Element("Revisor")?.Value,
                    };
                })
                .ToList();

            var consolidatedDocxName = originalCopiedToDestDocx.Name.Replace(".docx", "-Consolidated.docx");
            var consolidatedDocumentFi = new FileInfo(Path.Combine(thisTestTempDir.FullName, consolidatedDocxName));

            WmlDocument source1Wml = new WmlDocument(originalCopiedToDestDocx.FullName);
            WmlComparerSettings settings = new WmlComparerSettings();
            settings.DebugTempFileDi = thisTestTempDir;
            WmlDocument consolidatedWml = WmlComparer.Consolidate(
                source1Wml,
                revisedDocumentsArray,
                settings);
            var wml3 = WordprocessingMLUtil.BreakLinkToTemplate(consolidatedWml);
            wml3.SaveAs(consolidatedDocumentFi.FullName);

            var validationErrors = "";
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(consolidatedWml.DocumentByteArray, 0, consolidatedWml.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    OpenXmlValidator validator = new OpenXmlValidator();
                    var errors = validator.Validate(wDoc).Where(e => !ExpectedErrors.Contains(e.Description));
                    if (errors.Count() > 0)
                    {

                        var ind = "  ";
                        var sb = new StringBuilder();
                        foreach (var err in errors)
                        {
#if true
                            sb.Append("Error" + Environment.NewLine);
                            sb.Append(ind + "ErrorType: " + err.ErrorType.ToString() + Environment.NewLine);
                            sb.Append(ind + "Description: " + err.Description + Environment.NewLine);
                            sb.Append(ind + "Part: " + err.Part.Uri.ToString() + Environment.NewLine);
                            sb.Append(ind + "XPath: " + err.Path.XPath + Environment.NewLine);
#else
                        sb.Append("            \"" + err.Description + "\"," + Environment.NewLine);
#endif
                        }
                        validationErrors = sb.ToString();
                    }
                }
            }

            /************************************************************************************************************************/

            if (s_OpenWord)
            {
                FileInfo wordExe = new FileInfo(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");
                WordRunner.RunWord(wordExe, consolidatedDocumentFi);
                WordRunner.RunWord(wordExe, originalCopiedToDestDocx);

                var revisedList = revisedDocumentsXElement
                    .Elements()
                    .Select(z =>
                    {
                        FileInfo revisedDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, z.Element("DocName").Value));
                        var revisedCopiedToDestDocx = new FileInfo(Path.Combine(thisTestTempDir.FullName, revisedDocx.Name));
                        return revisedCopiedToDestDocx;
                    })
                    .ToList();
                foreach (var item in revisedList)
                    WordRunner.RunWord(wordExe, item);
            }

            /************************************************************************************************************************/

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Open Windows Explorer
            if (m_OpenTempDirInExplorer)
            {
                while (true)
                {
                    try
                    {
                        ////////// CODE TO REPEAT UNTIL SUCCESS //////////
                        var semaphorFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "z_ExplorerOpenedSemaphore.txt"));
                        if (!semaphorFi.Exists)
                        {
                            File.WriteAllText(semaphorFi.FullName, "");
                            TestUtil.Explorer(thisTestTempDir);
                        }
                        //////////////////////////////////////////////////
                        break;
                    }
                    catch (IOException)
                    {
                        System.Threading.Thread.Sleep(50);
                    }
                }
            }

            if (validationErrors != "")
                Assert.True(false, validationErrors);
        }

        [Theory]
        [InlineData("WCB-1000", "CA/CA001-Plain.docx", "CA/CA001-Plain-Mod.docx")]
        [InlineData("WCB-1010", "WC/WC001-Digits.docx", "WC/WC001-Digits-Mod.docx")]
        [InlineData("WCB-1020", "WC/WC001-Digits.docx", "WC/WC001-Digits-Deleted-Paragraph.docx")]
        [InlineData("WCB-1030", "WC/WC001-Digits-Deleted-Paragraph.docx", "WC/WC001-Digits.docx")]
        [InlineData("WCB-1040", "WC/WC002-Unmodified.docx", "WC/WC002-DiffInMiddle.docx")]
        [InlineData("WCB-1050", "WC/WC002-Unmodified.docx", "WC/WC002-DiffAtBeginning.docx")]
        [InlineData("WCB-1060", "WC/WC002-Unmodified.docx", "WC/WC002-DeleteAtBeginning.docx")]
        [InlineData("WCB-1070", "WC/WC002-Unmodified.docx", "WC/WC002-InsertAtBeginning.docx")]
        [InlineData("WCB-1080", "WC/WC002-Unmodified.docx", "WC/WC002-InsertAtEnd.docx")]
        [InlineData("WCB-1090", "WC/WC002-Unmodified.docx", "WC/WC002-DeleteAtEnd.docx")]
        [InlineData("WCB-1100", "WC/WC002-Unmodified.docx", "WC/WC002-DeleteInMiddle.docx")]
        [InlineData("WCB-1110", "WC/WC002-Unmodified.docx", "WC/WC002-InsertInMiddle.docx")]
        [InlineData("WCB-1120", "WC/WC002-DeleteInMiddle.docx", "WC/WC002-Unmodified.docx")]
        //[InlineData("WCB-1130", "WC/WC004-Large.docx", "WC/WC004-Large-Mod.docx")]
        [InlineData("WCB-1140", "WC/WC006-Table.docx", "WC/WC006-Table-Delete-Row.docx")]
        [InlineData("WCB-1150", "WC/WC006-Table-Delete-Row.docx", "WC/WC006-Table.docx")]
        [InlineData("WCB-1160", "WC/WC006-Table.docx", "WC/WC006-Table-Delete-Contests-of-Row.docx")]
        [InlineData("WCB-1170", "WC/WC007-Unmodified.docx", "WC/WC007-Longest-At-End.docx")]
        [InlineData("WCB-1180", "WC/WC007-Unmodified.docx", "WC/WC007-Deleted-at-Beginning-of-Para.docx")]
        [InlineData("WCB-1190", "WC/WC007-Unmodified.docx", "WC/WC007-Moved-into-Table.docx")]
        [InlineData("WCB-1200", "WC/WC009-Table-Unmodified.docx", "WC/WC009-Table-Cell-1-1-Mod.docx")]
        [InlineData("WCB-1210", "WC/WC010-Para-Before-Table-Unmodified.docx", "WC/WC010-Para-Before-Table-Mod.docx")]
        [InlineData("WCB-1220", "WC/WC011-Before.docx", "WC/WC011-After.docx")]
        [InlineData("WCB-1230", "WC/WC012-Math-Before.docx", "WC/WC012-Math-After.docx")]
        [InlineData("WCB-1240", "WC/WC013-Image-Before.docx", "WC/WC013-Image-After.docx")]
        [InlineData("WCB-1250", "WC/WC013-Image-Before.docx", "WC/WC013-Image-After2.docx")]
        [InlineData("WCB-1260", "WC/WC013-Image-Before2.docx", "WC/WC013-Image-After2.docx")]
        [InlineData("WCB-1270", "WC/WC014-SmartArt-Before.docx", "WC/WC014-SmartArt-After.docx")]
        [InlineData("WCB-1280", "WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-After.docx")]
        [InlineData("WCB-1290", "WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-Deleted-After.docx")]
        [InlineData("WCB-1300", "WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-Deleted-After2.docx")]
        [InlineData("WCB-1310", "WC/WC015-Three-Paragraphs.docx", "WC/WC015-Three-Paragraphs-After.docx")]
        [InlineData("WCB-1320", "WC/WC016-Para-Image-Para.docx", "WC/WC016-Para-Image-Para-w-Deleted-Image.docx")]
        [InlineData("WCB-1330", "WC/WC017-Image.docx", "WC/WC017-Image-After.docx")]
        [InlineData("WCB-1340", "WC/WC018-Field-Simple-Before.docx", "WC/WC018-Field-Simple-After-1.docx")]
        [InlineData("WCB-1350", "WC/WC018-Field-Simple-Before.docx", "WC/WC018-Field-Simple-After-2.docx")]
        [InlineData("WCB-1360", "WC/WC019-Hyperlink-Before.docx", "WC/WC019-Hyperlink-After-1.docx")]
        [InlineData("WCB-1370", "WC/WC019-Hyperlink-Before.docx", "WC/WC019-Hyperlink-After-2.docx")]
        [InlineData("WCB-1380", "WC/WC020-FootNote-Before.docx", "WC/WC020-FootNote-After-1.docx")]
        [InlineData("WCB-1390", "WC/WC020-FootNote-Before.docx", "WC/WC020-FootNote-After-2.docx")]
        [InlineData("WCB-1400", "WC/WC021-Math-Before-1.docx", "WC/WC021-Math-After-1.docx")]
        [InlineData("WCB-1410", "WC/WC021-Math-Before-2.docx", "WC/WC021-Math-After-2.docx")]
        [InlineData("WCB-1420", "WC/WC022-Image-Math-Para-Before.docx", "WC/WC022-Image-Math-Para-After.docx")]
        [InlineData("WCB-1430", "WC/WC023-Table-4-Row-Image-Before.docx", "WC/WC023-Table-4-Row-Image-After-Delete-1-Row.docx")]
        [InlineData("WCB-1440", "WC/WC024-Table-Before.docx", "WC/WC024-Table-After.docx")]
        [InlineData("WCB-1450", "WC/WC024-Table-Before.docx", "WC/WC024-Table-After2.docx")]
        [InlineData("WCB-1460", "WC/WC025-Simple-Table-Before.docx", "WC/WC025-Simple-Table-After.docx")]
        [InlineData("WCB-1470", "WC/WC026-Long-Table-Before.docx", "WC/WC026-Long-Table-After-1.docx")]
        [InlineData("WCB-1480", "WC/WC027-Twenty-Paras-Before.docx", "WC/WC027-Twenty-Paras-After-1.docx")]
        [InlineData("WCB-1490", "WC/WC027-Twenty-Paras-After-1.docx", "WC/WC027-Twenty-Paras-Before.docx")]
        [InlineData("WCB-1500", "WC/WC027-Twenty-Paras-Before.docx", "WC/WC027-Twenty-Paras-After-2.docx")]
        [InlineData("WCB-1510", "WC/WC030-Image-Math-Before.docx", "WC/WC030-Image-Math-After.docx")]
        [InlineData("WCB-1520", "WC/WC031-Two-Maths-Before.docx", "WC/WC031-Two-Maths-After.docx")]
        [InlineData("WCB-1530", "WC/WC032-Para-with-Para-Props.docx", "WC/WC032-Para-with-Para-Props-After.docx")]
        [InlineData("WCB-1540", "WC/WC033-Merged-Cells-Before.docx", "WC/WC033-Merged-Cells-After1.docx")]
        [InlineData("WCB-1550", "WC/WC033-Merged-Cells-Before.docx", "WC/WC033-Merged-Cells-After2.docx")]
        [InlineData("WCB-1560", "WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After1.docx")]
        [InlineData("WCB-1570", "WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After2.docx")]
        [InlineData("WCB-1580", "WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After3.docx")]
        [InlineData("WCB-1590", "WC/WC034-Footnotes-After3.docx", "WC/WC034-Footnotes-Before.docx")]
        [InlineData("WCB-1600", "WC/WC035-Footnote-Before.docx", "WC/WC035-Footnote-After.docx")]
        [InlineData("WCB-1610", "WC/WC035-Footnote-After.docx", "WC/WC035-Footnote-Before.docx")]
        [InlineData("WCB-1620", "WC/WC036-Footnote-With-Table-Before.docx", "WC/WC036-Footnote-With-Table-After.docx")]
        [InlineData("WCB-1630", "WC/WC036-Footnote-With-Table-After.docx", "WC/WC036-Footnote-With-Table-Before.docx")]
        [InlineData("WCB-1640", "WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After1.docx")]
        [InlineData("WCB-1650", "WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After2.docx")]
        [InlineData("WCB-1660", "WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After3.docx")]
        [InlineData("WCB-1670", "WC/WC034-Endnotes-After3.docx", "WC/WC034-Endnotes-Before.docx")]
        [InlineData("WCB-1680", "WC/WC035-Endnote-Before.docx", "WC/WC035-Endnote-After.docx")]
        [InlineData("WCB-1690", "WC/WC035-Endnote-After.docx", "WC/WC035-Endnote-Before.docx")]
        [InlineData("WCB-1700", "WC/WC036-Endnote-With-Table-Before.docx", "WC/WC036-Endnote-With-Table-After.docx")]
        [InlineData("WCB-1710", "WC/WC036-Endnote-With-Table-After.docx", "WC/WC036-Endnote-With-Table-Before.docx")]
        [InlineData("WCB-1720", "WC/WC038-Document-With-BR-Before.docx", "WC/WC038-Document-With-BR-After.docx")]
        [InlineData("WCB-1730", "RC/RC001-Before.docx", "RC/RC001-After1.docx")]
        [InlineData("WCB-1740", "RC/RC002-Image.docx", "RC/RC002-Image-After1.docx")]
        //[InlineData("WCB-1000", "", "")]
        //[InlineData("WCB-1000", "", "")]
        //[InlineData("WCB-1000", "", "")]
        //[InlineData("WCB-1000", "", "")]
        //[InlineData("WCB-1000", "", "")]
        //[InlineData("WCB-1000", "", "")]
        //[InlineData("WCB-1000", "", "")]


        public void WC002_Consolidate_Bulk_Test(string testId, string name1, string name2)
        {
            FileInfo source1Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
            FileInfo source2Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

            var rootTempDir = TestUtil.TempDir;
            var thisTestTempDir = new DirectoryInfo(Path.Combine(rootTempDir.FullName, testId));
            if (thisTestTempDir.Exists)
                Assert.True(false, "Duplicate test id: " + testId);
            else
                thisTestTempDir.Create();

            var source1CopiedToDestDocx = new FileInfo(Path.Combine(thisTestTempDir.FullName, source1Docx.Name));
            var source2CopiedToDestDocx = new FileInfo(Path.Combine(thisTestTempDir.FullName, source2Docx.Name));
            if (!source1CopiedToDestDocx.Exists)
            {
                var wml1 = new WmlDocument(source1Docx.FullName);
                var wml2 = WordprocessingMLUtil.BreakLinkToTemplate(wml1);
                wml2.SaveAs(source1CopiedToDestDocx.FullName);
            }
            if (!source2CopiedToDestDocx.Exists)
            {
                var wml1 = new WmlDocument(source2Docx.FullName);
                var wml2 = WordprocessingMLUtil.BreakLinkToTemplate(wml1);
                wml2.SaveAs(source2CopiedToDestDocx.FullName);
            }

            /************************************************************************************************************************/

            if (s_OpenWord)
            {
                FileInfo source1DocxForWord = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
                FileInfo source2DocxForWord = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

                var source1CopiedToDestDocxForWord = new FileInfo(Path.Combine(thisTestTempDir.FullName, source1Docx.Name.Replace(".docx", "-For-Word.docx")));
                var source2CopiedToDestDocxForWord = new FileInfo(Path.Combine(thisTestTempDir.FullName, source2Docx.Name.Replace(".docx", "-For-Word.docx")));
                if (!source1CopiedToDestDocxForWord.Exists)
                    File.Copy(source1Docx.FullName, source1CopiedToDestDocxForWord.FullName);
                if (!source2CopiedToDestDocxForWord.Exists)
                    File.Copy(source2Docx.FullName, source2CopiedToDestDocxForWord.FullName);

                FileInfo wordExe = new FileInfo(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");
                var path = new DirectoryInfo(@"C:\Users\Eric\Documents\WindowsPowerShellModules\Open-Xml-PowerTools\TestFiles");
                WordRunner.RunWord(wordExe, source2CopiedToDestDocxForWord);
                WordRunner.RunWord(wordExe, source1CopiedToDestDocxForWord);
            }

            /************************************************************************************************************************/

            var before = source1CopiedToDestDocx.Name.Replace(".docx", "");
            var after = source2CopiedToDestDocx.Name.Replace(".docx", "");
            var docxWithRevisionsFi = new FileInfo(Path.Combine(thisTestTempDir.FullName, before + "-COMPARE-" + after + ".docx"));
            var docxConsolidatedFi = new FileInfo(Path.Combine(thisTestTempDir.FullName, before + "-CONSOLIDATED-" + after + ".docx"));

            WmlDocument source1Wml = new WmlDocument(source1CopiedToDestDocx.FullName);
            WmlDocument source2Wml = new WmlDocument(source2CopiedToDestDocx.FullName);
            WmlComparerSettings settings = new WmlComparerSettings();
            WmlDocument comparedWml = WmlComparer.Compare(source1Wml, source2Wml, settings);
            WordprocessingMLUtil.BreakLinkToTemplate(comparedWml).SaveAs(docxWithRevisionsFi.FullName);

            List<WmlRevisedDocumentInfo> revisedDocInfo = new List<WmlRevisedDocumentInfo>()
            {
                new WmlRevisedDocumentInfo()
                {
                    RevisedDocument = source2Wml,
                    Color = Color.LightBlue,
                    Revisor = "Revised by Eric White",
                }
            };
            WmlDocument consolidatedWml = WmlComparer.Consolidate(
                source1Wml,
                revisedDocInfo,
                settings);
            WordprocessingMLUtil.BreakLinkToTemplate(consolidatedWml).SaveAs(docxConsolidatedFi.FullName);

            string validationErrors = "";
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(consolidatedWml.DocumentByteArray, 0, consolidatedWml.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    OpenXmlValidator validator = new OpenXmlValidator();
                    var errors = validator.Validate(wDoc).Where(e => !ExpectedErrors.Contains(e.Description));
                    if (errors.Count() > 0)
                    {

                        var ind = "  ";
                        var sb = new StringBuilder();
                        foreach (var err in errors)
                        {
#if true
                            sb.Append("Error" + Environment.NewLine);
                            sb.Append(ind + "ErrorType: " + err.ErrorType.ToString() + Environment.NewLine);
                            sb.Append(ind + "Description: " + err.Description + Environment.NewLine);
                            sb.Append(ind + "Part: " + err.Part.Uri.ToString() + Environment.NewLine);
                            sb.Append(ind + "XPath: " + err.Path.XPath + Environment.NewLine);
#else
                        sb.Append("            \"" + err.Description + "\"," + Environment.NewLine);
#endif
                        }
                        validationErrors = sb.ToString();
                    }
                }
            }

            /************************************************************************************************************************/

            if (s_OpenWord)
            {
                FileInfo wordExe = new FileInfo(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");
                WordRunner.RunWord(wordExe, docxConsolidatedFi);
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Open Windows Explorer
            if (m_OpenTempDirInExplorer)
            {
                while (true)
                {
                    try
                    {
                        ////////// CODE TO REPEAT UNTIL SUCCESS //////////
                        var semaphorFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "z_ExplorerOpenedSemaphore.txt"));
                        if (!semaphorFi.Exists)
                        {
                            File.WriteAllText(semaphorFi.FullName, "");
                            TestUtil.Explorer(thisTestTempDir);
                        }
                        //////////////////////////////////////////////////
                        break;
                    }
                    catch (IOException)
                    {
                        System.Threading.Thread.Sleep(50);
                    }
                }
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            if (validationErrors != "")
                Assert.True(false, validationErrors);
        }

        [Theory]
        [InlineData("WC-1000", "CA/CA001-Plain.docx", "CA/CA001-Plain-Mod.docx", 1)]
        [InlineData("WC-1010", "WC/WC001-Digits.docx", "WC/WC001-Digits-Mod.docx", 4)]
        [InlineData("WC-1020", "WC/WC001-Digits.docx", "WC/WC001-Digits-Deleted-Paragraph.docx", 1)]
        [InlineData("WC-1030", "WC/WC001-Digits-Deleted-Paragraph.docx", "WC/WC001-Digits.docx", 1)]
        [InlineData("WC-1040", "WC/WC002-Unmodified.docx", "WC/WC002-DiffInMiddle.docx", 2)]
        [InlineData("WC-1050", "WC/WC002-Unmodified.docx", "WC/WC002-DiffAtBeginning.docx", 2)]
        [InlineData("WC-1060", "WC/WC002-Unmodified.docx", "WC/WC002-DeleteAtBeginning.docx", 1)]
        [InlineData("WC-1070", "WC/WC002-Unmodified.docx", "WC/WC002-InsertAtBeginning.docx", 1)]
        [InlineData("WC-1080", "WC/WC002-Unmodified.docx", "WC/WC002-InsertAtEnd.docx", 1)]
        [InlineData("WC-1090", "WC/WC002-Unmodified.docx", "WC/WC002-DeleteAtEnd.docx", 1)]
        [InlineData("WC-1100", "WC/WC002-Unmodified.docx", "WC/WC002-DeleteInMiddle.docx", 1)]
        [InlineData("WC-1110", "WC/WC002-Unmodified.docx", "WC/WC002-InsertInMiddle.docx", 1)]
        [InlineData("WC-1120", "WC/WC002-DeleteInMiddle.docx", "WC/WC002-Unmodified.docx", 1)]
        //[InlineData("WC-1130", "WC/WC004-Large.docx", "WC/WC004-Large-Mod.docx", 2)]
        [InlineData("WC-1140", "WC/WC006-Table.docx", "WC/WC006-Table-Delete-Row.docx", 1)]
        [InlineData("WC-1150", "WC/WC006-Table-Delete-Row.docx", "WC/WC006-Table.docx", 1)]
        [InlineData("WC-1160", "WC/WC006-Table.docx", "WC/WC006-Table-Delete-Contests-of-Row.docx", 2)]
        [InlineData("WC-1170", "WC/WC007-Unmodified.docx", "WC/WC007-Longest-At-End.docx", 2)]
        [InlineData("WC-1180", "WC/WC007-Unmodified.docx", "WC/WC007-Deleted-at-Beginning-of-Para.docx", 1)]
        [InlineData("WC-1190", "WC/WC007-Unmodified.docx", "WC/WC007-Moved-into-Table.docx", 2)]
        [InlineData("WC-1200", "WC/WC009-Table-Unmodified.docx", "WC/WC009-Table-Cell-1-1-Mod.docx", 1)]
        [InlineData("WC-1210", "WC/WC010-Para-Before-Table-Unmodified.docx", "WC/WC010-Para-Before-Table-Mod.docx", 3)]
        [InlineData("WC-1220", "WC/WC011-Before.docx", "WC/WC011-After.docx", 2)]
        [InlineData("WC-1230", "WC/WC012-Math-Before.docx", "WC/WC012-Math-After.docx", 2)]
        [InlineData("WC-1240", "WC/WC013-Image-Before.docx", "WC/WC013-Image-After.docx", 2)]
        [InlineData("WC-1250", "WC/WC013-Image-Before.docx", "WC/WC013-Image-After2.docx", 2)]
        [InlineData("WC-1260", "WC/WC013-Image-Before2.docx", "WC/WC013-Image-After2.docx", 2)]
        [InlineData("WC-1270", "WC/WC014-SmartArt-Before.docx", "WC/WC014-SmartArt-After.docx", 2)]
        [InlineData("WC-1280", "WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-After.docx", 2)]
        [InlineData("WC-1310", "WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-Deleted-After.docx", 3)]
        [InlineData("WC-1320", "WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-Deleted-After2.docx", 1)]
        [InlineData("WC-1330", "WC/WC015-Three-Paragraphs.docx", "WC/WC015-Three-Paragraphs-After.docx", 3)]
        [InlineData("WC-1340", "WC/WC016-Para-Image-Para.docx", "WC/WC016-Para-Image-Para-w-Deleted-Image.docx", 1)]
        [InlineData("WC-1350", "WC/WC017-Image.docx", "WC/WC017-Image-After.docx", 3)]
        [InlineData("WC-1360", "WC/WC018-Field-Simple-Before.docx", "WC/WC018-Field-Simple-After-1.docx", 2)]
        [InlineData("WC-1370", "WC/WC018-Field-Simple-Before.docx", "WC/WC018-Field-Simple-After-2.docx", 3)]
        [InlineData("WC-1380", "WC/WC019-Hyperlink-Before.docx", "WC/WC019-Hyperlink-After-1.docx", 3)]
        [InlineData("WC-1390", "WC/WC019-Hyperlink-Before.docx", "WC/WC019-Hyperlink-After-2.docx", 5)]
        [InlineData("WC-1400", "WC/WC020-FootNote-Before.docx", "WC/WC020-FootNote-After-1.docx", 3)]
        [InlineData("WC-1410", "WC/WC020-FootNote-Before.docx", "WC/WC020-FootNote-After-2.docx", 5)]
        [InlineData("WC-1420", "WC/WC021-Math-Before-1.docx", "WC/WC021-Math-After-1.docx", 9)]
        [InlineData("WC-1430", "WC/WC021-Math-Before-2.docx", "WC/WC021-Math-After-2.docx", 6)]
        [InlineData("WC-1440", "WC/WC022-Image-Math-Para-Before.docx", "WC/WC022-Image-Math-Para-After.docx", 10)]
        [InlineData("WC-1450", "WC/WC023-Table-4-Row-Image-Before.docx", "WC/WC023-Table-4-Row-Image-After-Delete-1-Row.docx", 7)]
        [InlineData("WC-1460", "WC/WC024-Table-Before.docx", "WC/WC024-Table-After.docx", 1)]
        [InlineData("WC-1470", "WC/WC024-Table-Before.docx", "WC/WC024-Table-After2.docx", 7)]
        [InlineData("WC-1480", "WC/WC025-Simple-Table-Before.docx", "WC/WC025-Simple-Table-After.docx", 4)]
        [InlineData("WC-1500", "WC/WC026-Long-Table-Before.docx", "WC/WC026-Long-Table-After-1.docx", 2)]
        [InlineData("WC-1510", "WC/WC027-Twenty-Paras-Before.docx", "WC/WC027-Twenty-Paras-After-1.docx", 2)]
        [InlineData("WC-1520", "WC/WC027-Twenty-Paras-After-1.docx", "WC/WC027-Twenty-Paras-Before.docx", 2)]
        [InlineData("WC-1530", "WC/WC027-Twenty-Paras-Before.docx", "WC/WC027-Twenty-Paras-After-2.docx", 4)]
        [InlineData("WC-1540", "WC/WC030-Image-Math-Before.docx", "WC/WC030-Image-Math-After.docx", 2)]
        [InlineData("WC-1550", "WC/WC031-Two-Maths-Before.docx", "WC/WC031-Two-Maths-After.docx", 4)]
        [InlineData("WC-1560", "WC/WC032-Para-with-Para-Props.docx", "WC/WC032-Para-with-Para-Props-After.docx", 3)]
        [InlineData("WC-1570", "WC/WC033-Merged-Cells-Before.docx", "WC/WC033-Merged-Cells-After1.docx", 2)]
        [InlineData("WC-1580", "WC/WC033-Merged-Cells-Before.docx", "WC/WC033-Merged-Cells-After2.docx", 4)]
        [InlineData("WC-1600", "WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After1.docx", 1)]
        [InlineData("WC-1610", "WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After2.docx", 4)]
        [InlineData("WC-1620", "WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After3.docx", 3)]
        [InlineData("WC-1630", "WC/WC034-Footnotes-After3.docx", "WC/WC034-Footnotes-Before.docx", 3)]
        [InlineData("WC-1640", "WC/WC035-Footnote-Before.docx", "WC/WC035-Footnote-After.docx", 2)]
        [InlineData("WC-1650", "WC/WC035-Footnote-After.docx", "WC/WC035-Footnote-Before.docx", 2)]
        [InlineData("WC-1660", "WC/WC036-Footnote-With-Table-Before.docx", "WC/WC036-Footnote-With-Table-After.docx", 5)]
        [InlineData("WC-1670", "WC/WC036-Footnote-With-Table-After.docx", "WC/WC036-Footnote-With-Table-Before.docx", 5)]
        [InlineData("WC-1680", "WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After1.docx", 1)]
        [InlineData("WC-1700", "WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After2.docx", 4)]
        [InlineData("WC-1710", "WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After3.docx", 7)]
        [InlineData("WC-1720", "WC/WC034-Endnotes-After3.docx", "WC/WC034-Endnotes-Before.docx", 7)]
        [InlineData("WC-1730", "WC/WC035-Endnote-Before.docx", "WC/WC035-Endnote-After.docx", 2)]
        [InlineData("WC-1740", "WC/WC035-Endnote-After.docx", "WC/WC035-Endnote-Before.docx", 2)]
        [InlineData("WC-1750", "WC/WC036-Endnote-With-Table-Before.docx", "WC/WC036-Endnote-With-Table-After.docx", 6)]
        [InlineData("WC-1760", "WC/WC036-Endnote-With-Table-After.docx", "WC/WC036-Endnote-With-Table-Before.docx", 6)]
        [InlineData("WC-1770", "WC/WC037-Textbox-Before.docx", "WC/WC037-Textbox-After1.docx", 2)]
        [InlineData("WC-1780", "WC/WC038-Document-With-BR-Before.docx", "WC/WC038-Document-With-BR-After.docx", 2)]
        [InlineData("WC-1800", "RC/RC001-Before.docx", "RC/RC001-After1.docx", 2)]
        [InlineData("WC-1810", "RC/RC002-Image.docx", "RC/RC002-Image-After1.docx", 1)]
        [InlineData("WC-1820", "WC/WC039-Break-In-Row.docx", "WC/WC039-Break-In-Row-After1.docx", 1)]
        [InlineData("WC-1830", "WC/WC041-Table-5.docx", "WC/WC041-Table-5-Mod.docx", 2)]
        [InlineData("WC-1840", "WC/WC042-Table-5.docx", "WC/WC042-Table-5-Mod.docx", 2)]
        [InlineData("WC-1850", "WC/WC043-Nested-Table.docx", "WC/WC043-Nested-Table-Mod.docx", 2)]
        [InlineData("WC-1860", "WC/WC044-Text-Box.docx", "WC/WC044-Text-Box-Mod.docx", 2)]
        [InlineData("WC-1870", "WC/WC045-Text-Box.docx", "WC/WC045-Text-Box-Mod.docx", 2)]
        [InlineData("WC-1880", "WC/WC046-Two-Text-Box.docx", "WC/WC046-Two-Text-Box-Mod.docx", 2)]
        [InlineData("WC-1890", "WC/WC047-Two-Text-Box.docx", "WC/WC047-Two-Text-Box-Mod.docx", 2)]
        [InlineData("WC-1900", "WC/WC048-Text-Box-in-Cell.docx", "WC/WC048-Text-Box-in-Cell-Mod.docx", 6)]
        [InlineData("WC-1910", "WC/WC049-Text-Box-in-Cell.docx", "WC/WC049-Text-Box-in-Cell-Mod.docx", 5)]
        [InlineData("WC-1920", "WC/WC050-Table-in-Text-Box.docx", "WC/WC050-Table-in-Text-Box-Mod.docx", 8)]
        [InlineData("WC-1930", "WC/WC051-Table-in-Text-Box.docx", "WC/WC051-Table-in-Text-Box-Mod.docx", 9)]
        [InlineData("WC-1940", "WC/WC052-SmartArt-Same.docx", "WC/WC052-SmartArt-Same-Mod.docx", 2)]
        [InlineData("WC-1950", "WC/WC053-Text-in-Cell.docx", "WC/WC053-Text-in-Cell-Mod.docx", 2)]
        [InlineData("WC-1960", "WC/WC054-Text-in-Cell.docx", "WC/WC054-Text-in-Cell-Mod.docx", 0)]
        [InlineData("WC-1970", "WC/WC055-French.docx", "WC/WC055-French-Mod.docx", 2)]
        [InlineData("WC-1980", "WC/WC056-French.docx", "WC/WC056-French-Mod.docx", 2)]
        [InlineData("WC-1990", "WC/WC057-Table-Merged-Cell.docx", "WC/WC057-Table-Merged-Cell-Mod.docx", 4)]
        [InlineData("WC-2000", "WC/WC058-Table-Merged-Cell.docx", "WC/WC058-Table-Merged-Cell-Mod.docx", 6)]
        [InlineData("WC-2010", "WC/WC059-Footnote.docx", "WC/WC059-Footnote-Mod.docx", 5)]
        [InlineData("WC-2020", "WC/WC060-Endnote.docx", "WC/WC060-Endnote-Mod.docx", 3)]
        [InlineData("WC-2030", "WC/WC061-Style-Added.docx", "WC/WC061-Style-Added-Mod.docx", 1)]
        [InlineData("WC-2040", "WC/WC062-New-Char-Style-Added.docx", "WC/WC062-New-Char-Style-Added-Mod.docx", 2)]
        [InlineData("WC-2050", "WC/WC063-Footnote.docx", "WC/WC063-Footnote-Mod.docx", 1)]
        [InlineData("WC-2060", "WC/WC063-Footnote-Mod.docx", "WC/WC063-Footnote.docx", 1)]
        [InlineData("WC-2070", "WC/WC064-Footnote.docx", "WC/WC064-Footnote-Mod.docx", 0)]
        [InlineData("WC-2080", "WC/WC065-Textbox.docx", "WC/WC065-Textbox-Mod.docx", 2)]
        [InlineData("WC-2090", "WC/WC066-Textbox-Before-Ins.docx", "WC/WC066-Textbox-Before-Ins-Mod.docx", 1)]
        [InlineData("WC-2092", "WC/WC066-Textbox-Before-Ins-Mod.docx", "WC/WC066-Textbox-Before-Ins.docx", 1)]
        [InlineData("WC-2100", "WC/WC067-Textbox-Image.docx", "WC/WC067-Textbox-Image-Mod.docx", 2)]
        //[InlineData("WC-1000", "", "", 0)]
        //[InlineData("WC-1000", "", "", 0)]
        //[InlineData("WC-1000", "", "", 0)]
        //[InlineData("WC-1000", "", "", 0)]
        //[InlineData("WC-1000", "", "", 0)]
        //[InlineData("WC-1000", "", "", 0)]

        public void WC003_Compare(string testId, string name1, string name2, int revisionCount)
        {
            FileInfo source1Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
            FileInfo source2Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

            var rootTempDir = TestUtil.TempDir;
            var thisTestTempDir = new DirectoryInfo(Path.Combine(rootTempDir.FullName, testId));
            if (thisTestTempDir.Exists)
                Assert.True(false, "Duplicate test id???");
            else
                thisTestTempDir.Create();

            var source1CopiedToDestDocx = new FileInfo(Path.Combine(thisTestTempDir.FullName, source1Docx.Name));
            var source2CopiedToDestDocx = new FileInfo(Path.Combine(thisTestTempDir.FullName, source2Docx.Name));
            File.Copy(source1Docx.FullName, source1CopiedToDestDocx.FullName);
            File.Copy(source2Docx.FullName, source2CopiedToDestDocx.FullName);

            var before = source1CopiedToDestDocx.Name.Replace(".docx", "");
            var after = source2CopiedToDestDocx.Name.Replace(".docx", "");
            //var baselineDocxWithRevisionsFi = new FileInfo(Path.Combine(source1Docx.DirectoryName, before + "-COMPARE-" + after + ".docx"));
            var docxWithRevisionsFi = new FileInfo(Path.Combine(thisTestTempDir.FullName, before + "-COMPARE-" + after + ".docx"));

            /************************************************************************************************************************/

            if (s_OpenWord)
            {
                FileInfo source1DocxForWord = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
                FileInfo source2DocxForWord = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

                var source1CopiedToDestDocxForWord = new FileInfo(Path.Combine(thisTestTempDir.FullName, source1Docx.Name.Replace(".docx", "-For-Word.docx")));
                var source2CopiedToDestDocxForWord = new FileInfo(Path.Combine(thisTestTempDir.FullName, source2Docx.Name.Replace(".docx", "-For-Word.docx")));
                File.Copy(source1Docx.FullName, source1CopiedToDestDocxForWord.FullName);
                File.Copy(source2Docx.FullName, source2CopiedToDestDocxForWord.FullName);

                FileInfo wordExe = new FileInfo(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");
                var path = new DirectoryInfo(@"C:\Users\Eric\Documents\WindowsPowerShellModules\Open-Xml-PowerTools\TestFiles");
                WordRunner.RunWord(wordExe, source2CopiedToDestDocxForWord);
                WordRunner.RunWord(wordExe, source1CopiedToDestDocxForWord);
            }

            /************************************************************************************************************************/

            WmlDocument source1Wml = new WmlDocument(source1CopiedToDestDocx.FullName);
            WmlDocument source2Wml = new WmlDocument(source2CopiedToDestDocx.FullName);
            WmlComparerSettings settings = new WmlComparerSettings();
            settings.DebugTempFileDi = thisTestTempDir;
            WmlDocument comparedWml = WmlComparer.Compare(source1Wml, source2Wml, settings);
            comparedWml.SaveAs(docxWithRevisionsFi.FullName);

#if false
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // create batch file to copy newly generated ContentTypeXml and NarrDoc to the TestFiles directory.
            while (true)
            {
                try
                {
                    ////////// CODE TO REPEAT UNTIL SUCCESS //////////
                    var batchFileName = "Copy-Gen-Files-To-TestFiles.bat";
                    var batchFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, batchFileName));
                    var batch = "";
                    batch += "copy " + docxWithRevisionsFi.FullName + " " + source1Docx.DirectoryName + Environment.NewLine;
                    if (batchFi.Exists)
                        File.AppendAllText(batchFi.FullName, batch);
                    else
                        File.WriteAllText(batchFi.FullName, batch);
                    //////////////////////////////////////////////////
                    break;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                }
            }
#endif

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // validate generated document
            var validationErrors = "";
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(comparedWml.DocumentByteArray, 0, comparedWml.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    OpenXmlValidator validator = new OpenXmlValidator();
                    var errors = validator.Validate(wDoc).Where(e => !ExpectedErrors.Contains(e.Description));
                    if (errors.Count() > 0)
                    {

                        var ind = "  ";
                        var sb = new StringBuilder();
                        foreach (var err in errors)
                        {
#if true
                            sb.Append("Error" + Environment.NewLine);
                            sb.Append(ind + "ErrorType: " + err.ErrorType.ToString() + Environment.NewLine);
                            sb.Append(ind + "Description: " + err.Description + Environment.NewLine);
                            sb.Append(ind + "Part: " + err.Part.Uri.ToString() + Environment.NewLine);
                            sb.Append(ind + "XPath: " + err.Path.XPath + Environment.NewLine);
#else
                        sb.Append("            \"" + err.Description + "\"," + Environment.NewLine);
#endif
                        }
                        validationErrors = sb.ToString();
                    }
                }
            }

            /************************************************************************************************************************/

            if (s_OpenWord)
            {
                FileInfo wordExe = new FileInfo(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");
                WordRunner.RunWord(wordExe, docxWithRevisionsFi);
            }

            /************************************************************************************************************************/


            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Open Windows Explorer
            if (m_OpenTempDirInExplorer)
            {
                while (true)
                {
                    try
                    {
                        ////////// CODE TO REPEAT UNTIL SUCCESS //////////
                        var semaphorFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "z_ExplorerOpenedSemaphore.txt"));
                        if (!semaphorFi.Exists)
                        {
                            File.WriteAllText(semaphorFi.FullName, "");
                            TestUtil.Explorer(thisTestTempDir);
                        }
                        //////////////////////////////////////////////////
                        break;
                    }
                    catch (IOException)
                    {
                        System.Threading.Thread.Sleep(50);
                    }
                }
            }

            if (validationErrors != "")
            {
                Assert.True(false, validationErrors);
            }

            WmlComparerSettings settings2 = new WmlComparerSettings();

            WmlDocument revisionWml = new WmlDocument(docxWithRevisionsFi.FullName);
            var revisions = WmlComparer.GetRevisions(revisionWml, settings);
            Assert.Equal(revisionCount, revisions.Count());

            var afterRejectingWml = RevisionProcessor.RejectRevisions(revisionWml);

            var WRITE_TEMP_FILES = true;

            if (WRITE_TEMP_FILES)
            {
                var afterRejectingFi = new FileInfo(Path.Combine(thisTestTempDir.FullName, "AfterRejecting.docx"));
                afterRejectingWml.SaveAs(afterRejectingFi.FullName);
            }

            WmlDocument afterRejectingComparedWml = WmlComparer.Compare(source1Wml, afterRejectingWml, settings);
            var sanityCheck1 = WmlComparer.GetRevisions(afterRejectingComparedWml, settings);

            if (WRITE_TEMP_FILES)
            {
                var afterRejectingComparedFi = new FileInfo(Path.Combine(thisTestTempDir.FullName, "AfterRejectingCompared.docx"));
                afterRejectingComparedWml.SaveAs(afterRejectingComparedFi.FullName);
            }

            var afterAcceptingWml = RevisionProcessor.AcceptRevisions(revisionWml);

            if (WRITE_TEMP_FILES)
            {
                var afterAcceptingFi = new FileInfo(Path.Combine(thisTestTempDir.FullName, "AfterAccepting.docx"));
                afterAcceptingWml.SaveAs(afterAcceptingFi.FullName);
            }

            WmlDocument afterAcceptingComparedWml = WmlComparer.Compare(source2Wml, afterAcceptingWml, settings);
            var sanityCheck2 = WmlComparer.GetRevisions(afterAcceptingComparedWml, settings);

            if (WRITE_TEMP_FILES)
            {
                var afterAcceptingComparedFi = new FileInfo(Path.Combine(thisTestTempDir.FullName, "AfterAcceptingCompared.docx"));
                afterAcceptingComparedWml.SaveAs(afterAcceptingComparedFi.FullName);
            }

            if (sanityCheck1.Count() != 0)
                Assert.True(false, "Sanity Check #1 failed");
            if (sanityCheck2.Count() != 0)
                Assert.True(false, "Sanity Check #2 failed");
        }

#if false
        [Theory]
        [InlineData("WC/WC037-Textbox-Before.docx", "WC/WC037-Textbox-After1.docx", 2)]

        public void WC003_Throws(string name1, string name2, int revisionCount)
        {
            FileInfo source1Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
            FileInfo source2Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

            var source1CopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, source1Docx.Name));
            var source2CopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, source2Docx.Name));
            if (!source1CopiedToDestDocx.Exists)
                File.Copy(source1Docx.FullName, source1CopiedToDestDocx.FullName);
            if (!source2CopiedToDestDocx.Exists)
                File.Copy(source2Docx.FullName, source2CopiedToDestDocx.FullName);

            WmlDocument source1Wml = new WmlDocument(source1CopiedToDestDocx.FullName);
            WmlDocument source2Wml = new WmlDocument(source2CopiedToDestDocx.FullName);
            WmlComparerSettings settings = new WmlComparerSettings();
            Assert.Throws<OpenXmlPowerToolsException>(() =>
                {
                    WmlDocument comparedWml = WmlComparer.Compare(source1Wml, source2Wml, settings);
                });
        }
#endif

        [Theory]
        [InlineData("WCS-1000", "WC/WC001-Digits.docx")]
        [InlineData("WCS-1010", "WC/WC001-Digits-Deleted-Paragraph.docx")]
        [InlineData("WCS-1020", "WC/WC001-Digits-Mod.docx")]
        [InlineData("WCS-1030", "WC/WC002-DeleteAtBeginning.docx")]
        [InlineData("WCS-1040", "WC/WC002-DeleteAtEnd.docx")]
        [InlineData("WCS-1050", "WC/WC002-DeleteInMiddle.docx")]
        [InlineData("WCS-1060", "WC/WC002-DiffAtBeginning.docx")]
        [InlineData("WCS-1070", "WC/WC002-DiffInMiddle.docx")]
        [InlineData("WCS-1080", "WC/WC002-InsertAtBeginning.docx")]
        [InlineData("WCS-1090", "WC/WC002-InsertAtEnd.docx")]
        [InlineData("WCS-1100", "WC/WC002-InsertInMiddle.docx")]
        [InlineData("WCS-1110", "WC/WC002-Unmodified.docx")]
      //[InlineData("WCS-1120", "WC/WC004-Large.docx")]
      //[InlineData("WCS-1130", "WC/WC004-Large-Mod.docx")]
        [InlineData("WCS-1140", "WC/WC006-Table.docx")]
        [InlineData("WCS-1150", "WC/WC006-Table-Delete-Contests-of-Row.docx")]
        [InlineData("WCS-1160", "WC/WC006-Table-Delete-Row.docx")]
        [InlineData("WCS-1170", "WC/WC007-Deleted-at-Beginning-of-Para.docx")]
        [InlineData("WCS-1180", "WC/WC007-Longest-At-End.docx")]
        [InlineData("WCS-1190", "WC/WC007-Moved-into-Table.docx")]
        [InlineData("WCS-1200", "WC/WC007-Unmodified.docx")]
        [InlineData("WCS-1210", "WC/WC009-Table-Cell-1-1-Mod.docx")]
        [InlineData("WCS-1220", "WC/WC009-Table-Unmodified.docx")]
        [InlineData("WCS-1230", "WC/WC010-Para-Before-Table-Mod.docx")]
        [InlineData("WCS-1240", "WC/WC010-Para-Before-Table-Unmodified.docx")]
        [InlineData("WCS-1250", "WC/WC011-After.docx")]
        [InlineData("WCS-1260", "WC/WC011-Before.docx")]
        [InlineData("WCS-1270", "WC/WC012-Math-After.docx")]
        [InlineData("WCS-1280", "WC/WC012-Math-Before.docx")]
        [InlineData("WCS-1290", "WC/WC013-Image-After.docx")]
        [InlineData("WCS-1300", "WC/WC013-Image-After2.docx")]
        [InlineData("WCS-1310", "WC/WC013-Image-Before.docx")]
        [InlineData("WCS-1320", "WC/WC013-Image-Before2.docx")]
        [InlineData("WCS-1330", "WC/WC014-SmartArt-After.docx")]
        [InlineData("WCS-1340", "WC/WC014-SmartArt-Before.docx")]
        [InlineData("WCS-1350", "WC/WC014-SmartArt-With-Image-After.docx")]
        [InlineData("WCS-1360", "WC/WC014-SmartArt-With-Image-Before.docx")]
        [InlineData("WCS-1370", "WC/WC014-SmartArt-With-Image-Deleted-After.docx")]
        [InlineData("WCS-1380", "WC/WC014-SmartArt-With-Image-Deleted-After2.docx")]
        [InlineData("WCS-1390", "WC/WC015-Three-Paragraphs.docx")]
        [InlineData("WCS-1400", "WC/WC015-Three-Paragraphs-After.docx")]
        [InlineData("WCS-1410", "WC/WC016-Para-Image-Para.docx")]
        [InlineData("WCS-1420", "WC/WC016-Para-Image-Para-w-Deleted-Image.docx")]
        [InlineData("WCS-1430", "WC/WC017-Image.docx")]
        [InlineData("WCS-1440", "WC/WC017-Image-After.docx")]
        [InlineData("WCS-1450", "WC/WC018-Field-Simple-After-1.docx")]
        [InlineData("WCS-1460", "WC/WC018-Field-Simple-After-2.docx")]
        [InlineData("WCS-1470", "WC/WC018-Field-Simple-Before.docx")]
        [InlineData("WCS-1480", "WC/WC019-Hyperlink-After-1.docx")]
        [InlineData("WCS-1490", "WC/WC019-Hyperlink-After-2.docx")]
        [InlineData("WCS-1500", "WC/WC019-Hyperlink-Before.docx")]
        [InlineData("WCS-1510", "WC/WC020-FootNote-After-1.docx")]
        [InlineData("WCS-1520", "WC/WC020-FootNote-After-2.docx")]
        [InlineData("WCS-1530", "WC/WC020-FootNote-Before.docx")]
        [InlineData("WCS-1540", "WC/WC021-Math-After-1.docx")]
        [InlineData("WCS-1550", "WC/WC021-Math-Before-1.docx")]
        [InlineData("WCS-1560", "WC/WC022-Image-Math-Para-After.docx")]
        [InlineData("WCS-1570", "WC/WC022-Image-Math-Para-Before.docx")]
      //[InlineData("WCS-1580", "", "")]
      //[InlineData("WCS-1590", "", "")]
      //[InlineData("WCS-1600", "", "")]
      //[InlineData("WCS-1610", "", "")]

        public void WC004_Compare_To_Self(string testId, string name)
        {
            FileInfo sourceDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            var rootTempDir = TestUtil.TempDir;
            var thisTestTempDir = new DirectoryInfo(Path.Combine(rootTempDir.FullName, testId));
            if (thisTestTempDir.Exists)
                Assert.True(false, "Duplicate test id???");
            else
                thisTestTempDir.Create();

            var sourceCopiedToDestDocx = new FileInfo(Path.Combine(thisTestTempDir.FullName, sourceDocx.Name.Replace(".docx", "-Source.docx")));
            if (!sourceCopiedToDestDocx.Exists)
                File.Copy(sourceDocx.FullName, sourceCopiedToDestDocx.FullName);

            var before = sourceCopiedToDestDocx.Name.Replace(".docx", "");
            var docxComparedFi = new FileInfo(Path.Combine(thisTestTempDir.FullName, before + "-COMPARE" + ".docx"));
            var docxCompared2Fi = new FileInfo(Path.Combine(thisTestTempDir.FullName, before + "-COMPARE2" + ".docx"));

            WmlDocument source1Wml = new WmlDocument(sourceCopiedToDestDocx.FullName);
            WmlDocument source2Wml = new WmlDocument(sourceCopiedToDestDocx.FullName);
            WmlComparerSettings settings = new WmlComparerSettings();

            WmlDocument comparedWml = WmlComparer.Compare(source1Wml, source2Wml, settings);
            comparedWml.SaveAs(docxComparedFi.FullName);
            ValidateDocument(comparedWml);

            WmlDocument comparedWml2 = WmlComparer.Compare(comparedWml, source1Wml, settings);
            comparedWml2.SaveAs(docxCompared2Fi.FullName);
            ValidateDocument(comparedWml2);
        }

        [Theory]
        [InlineData("WCI-1000", "WC/WC040-Case-Before.docx", "WC/WC040-Case-After.docx", 2)]

        public void WC005_Compare_CaseInsensitive(string testId, string name1, string name2, int revisionCount)
        {
            FileInfo source1Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
            FileInfo source2Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

            var rootTempDir = TestUtil.TempDir;
            var thisTestTempDir = new DirectoryInfo(Path.Combine(rootTempDir.FullName, testId));
            if (thisTestTempDir.Exists)
                Assert.True(false, "Duplicate test id???");
            else
                thisTestTempDir.Create();

            var source1CopiedToDestDocx = new FileInfo(Path.Combine(thisTestTempDir.FullName, source1Docx.Name));
            var source2CopiedToDestDocx = new FileInfo(Path.Combine(thisTestTempDir.FullName, source2Docx.Name));
            if (!source1CopiedToDestDocx.Exists)
                File.Copy(source1Docx.FullName, source1CopiedToDestDocx.FullName);
            if (!source2CopiedToDestDocx.Exists)
                File.Copy(source2Docx.FullName, source2CopiedToDestDocx.FullName);

            /************************************************************************************************************************/

            if (s_OpenWord)
            {
                FileInfo source1DocxForWord = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
                FileInfo source2DocxForWord = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

                var source1CopiedToDestDocxForWord = new FileInfo(Path.Combine(thisTestTempDir.FullName, source1Docx.Name.Replace(".docx", "-For-Word.docx")));
                var source2CopiedToDestDocxForWord = new FileInfo(Path.Combine(thisTestTempDir.FullName, source2Docx.Name.Replace(".docx", "-For-Word.docx")));
                if (!source1CopiedToDestDocxForWord.Exists)
                    File.Copy(source1Docx.FullName, source1CopiedToDestDocxForWord.FullName);
                if (!source2CopiedToDestDocxForWord.Exists)
                    File.Copy(source2Docx.FullName, source2CopiedToDestDocxForWord.FullName);

                FileInfo wordExe = new FileInfo(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");
                var path = new DirectoryInfo(@"C:\Users\Eric\Documents\WindowsPowerShellModules\Open-Xml-PowerTools\TestFiles");
                WordRunner.RunWord(wordExe, source2CopiedToDestDocxForWord);
                WordRunner.RunWord(wordExe, source1CopiedToDestDocxForWord);
            }

            /************************************************************************************************************************/

            var before = source1CopiedToDestDocx.Name.Replace(".docx", "");
            var after = source2CopiedToDestDocx.Name.Replace(".docx", "");
            var docxWithRevisionsFi = new FileInfo(Path.Combine(thisTestTempDir.FullName, before + "-COMPARE-" + after + ".docx"));

            WmlDocument source1Wml = new WmlDocument(source1CopiedToDestDocx.FullName);
            WmlDocument source2Wml = new WmlDocument(source2CopiedToDestDocx.FullName);
            WmlComparerSettings settings = new WmlComparerSettings();
            settings.CaseInsensitive = true;
            settings.CultureInfo = System.Globalization.CultureInfo.CurrentCulture;
            WmlDocument comparedWml = WmlComparer.Compare(source1Wml, source2Wml, settings);
            comparedWml.SaveAs(docxWithRevisionsFi.FullName);

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(comparedWml.DocumentByteArray, 0, comparedWml.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    OpenXmlValidator validator = new OpenXmlValidator();
                    var errors = validator.Validate(wDoc).Where(e => !ExpectedErrors.Contains(e.Description));
                    if (errors.Count() > 0)
                    {

                        var ind = "  ";
                        var sb = new StringBuilder();
                        foreach (var err in errors)
                        {
#if true
                            sb.Append("Error" + Environment.NewLine);
                            sb.Append(ind + "ErrorType: " + err.ErrorType.ToString() + Environment.NewLine);
                            sb.Append(ind + "Description: " + err.Description + Environment.NewLine);
                            sb.Append(ind + "Part: " + err.Part.Uri.ToString() + Environment.NewLine);
                            sb.Append(ind + "XPath: " + err.Path.XPath + Environment.NewLine);
#else
                        sb.Append("            \"" + err.Description + "\"," + Environment.NewLine);
#endif
                        }
                        var sbs = sb.ToString();
                        Assert.Equal("", sbs);
                    }
                }
            }

            /************************************************************************************************************************/

            if (s_OpenWord)
            {
                FileInfo wordExe = new FileInfo(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");
                WordRunner.RunWord(wordExe, docxWithRevisionsFi);
            }

            /************************************************************************************************************************/

            WmlDocument revisionWml = new WmlDocument(docxWithRevisionsFi.FullName);
            var revisions = WmlComparer.GetRevisions(revisionWml, settings);
            Assert.Equal(revisionCount, revisions.Count());
        }

        private static void ValidateDocument(WmlDocument wmlToValidate)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(wmlToValidate.DocumentByteArray, 0, wmlToValidate.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    OpenXmlValidator validator = new OpenXmlValidator();
                    var errors = validator.Validate(wDoc).Where(e => !ExpectedErrors.Contains(e.Description));
                    if (errors.Count() != 0)
                    {
                        var ind = "  ";
                        var sb = new StringBuilder();
                        foreach (var err in errors)
                        {
#if true
                            sb.Append("Error" + Environment.NewLine);
                            sb.Append(ind + "ErrorType: " + err.ErrorType.ToString() + Environment.NewLine);
                            sb.Append(ind + "Description: " + err.Description + Environment.NewLine);
                            sb.Append(ind + "Part: " + err.Part.Uri.ToString() + Environment.NewLine);
                            sb.Append(ind + "XPath: " + err.Path.XPath + Environment.NewLine);
#else
                        sb.Append("            \"" + err.Description + "\"," + Environment.NewLine);
#endif

                        }
                        var sbs = sb.ToString();
                        Assert.Equal("", sbs);
                    }
                }
            }
        }

        public static string[] ExpectedErrors = new string[] {
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:allStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:customStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:latentStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:stylesInUse' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:headingStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:numberingStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:tableStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnRuns' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnParagraphs' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnNumbering' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnTables' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:clearFormatting' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:top3HeadingStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:visibleStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:alternateStyleNames' attribute is not declared.",
            "The attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:val' has invalid value '0'. The MinInclusive constraint failed. The value must be greater than or equal to 1.",
            "The attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:val' has invalid value '0'. The MinInclusive constraint failed. The value must be greater than or equal to 2.",
            "The 'urn:schemas-microsoft-com:office:office:gfxdata' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:fill' attribute is invalid - The value '0' is not valid according to any of the memberTypes of the union.",
        };

    }

    public class WordRunner
    {
        public static void RunWord(FileInfo executablePath, FileInfo docxPath)
        {
            if (executablePath.Exists)
            {
                using (Process proc = new Process())
                {
                    proc.StartInfo.FileName = executablePath.FullName;
                    proc.StartInfo.Arguments = docxPath.FullName;
                    proc.StartInfo.WorkingDirectory = docxPath.DirectoryName;
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.RedirectStandardOutput = true;
                    proc.StartInfo.RedirectStandardError = true;
                    proc.Start();
                }
            }
            else
            {
                throw new ArgumentException("Invalid executable path.", "executablePath");
            }
        }
    }
}

#endif
