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

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class WcTests2
    {
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static bool m_OpenWord = false;
        public static bool m_OpenTempDirInExplorer = false;
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        [Theory]
        [InlineData("CZ-1000", "CZ/CZ001-Plain.docx", "CZ/CZ001-Plain-Mod.docx", 1)]
        [InlineData("CZ-1010", "CZ/CZ002-Multi-Paragraphs.docx", "CZ/CZ002-Multi-Paragraphs-Mod.docx", 1)]
        [InlineData("CZ-1020", "CZ/CZ003-Multi-Paragraphs.docx", "CZ/CZ003-Multi-Paragraphs-Mod.docx", 1)]
        [InlineData("CZ-1030", "CZ/CZ004-Multi-Paragraphs-in-Cell.docx", "CZ/CZ004-Multi-Paragraphs-in-Cell-Mod.docx", 1)]
        public void CZ001_CompareTrackedInPrev(string testId, string name1, string name2, int revisionCount)
        {
            // TODO: Do we need to keep the revision count parameter?
            Assert.Equal(1, revisionCount);

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

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            if (m_OpenWord)
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
                WordRunner.RunWord(wordExe, source2CopiedToDestDocxForWord);
                WordRunner.RunWord(wordExe, source1CopiedToDestDocxForWord);
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            var before = source1CopiedToDestDocx.Name.Replace(".docx", "");
            var after = source2CopiedToDestDocx.Name.Replace(".docx", "");
            var docxWithRevisionsFi = new FileInfo(Path.Combine(thisTestTempDir.FullName, before + "-COMPARE-" + after + ".docx"));

            WmlDocument source1Wml = new WmlDocument(source1CopiedToDestDocx.FullName);
            WmlDocument source2Wml = new WmlDocument(source2CopiedToDestDocx.FullName);
            WmlComparerSettings settings = new WmlComparerSettings();
            settings.DebugTempFileDi = thisTestTempDir;
            WmlDocument comparedWml = WmlComparer.Compare(source1Wml, source2Wml, settings);

            ///////////////////////////
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
                            sb.Append("Error" + Environment.NewLine);
                            sb.Append(ind + "ErrorType: " + err.ErrorType.ToString() + Environment.NewLine);
                            sb.Append(ind + "Description: " + err.Description + Environment.NewLine);
                            sb.Append(ind + "Part: " + err.Part.Uri.ToString() + Environment.NewLine);
                            sb.Append(ind + "XPath: " + err.Path.XPath + Environment.NewLine);
                        }
                        var sbs = sb.ToString();
                        if (sbs != "")
                            Assert.True(false, sbs.ToString());
                    }
                }
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            if (m_OpenWord)
            {
                FileInfo wordExe = new FileInfo(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");
                WordRunner.RunWord(wordExe, docxWithRevisionsFi);
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
#if false
            WmlDocument revisionWml = new WmlDocument(docxWithRevisionsFi.FullName);
            var revisions = WmlComparer.GetRevisions(revisionWml, settings);
            Assert.Equal(revisionCount, revisions.Count());
#endif
        }

#if false
        [Theory]
        [InlineData("CZ-2000", "CA001-Plain.docx", "CA001-Plain-Mod.docx", 1)]
        [InlineData("CZ-2010", "WC001-Digits.docx", "WC001-Digits-Mod.docx", 4)]
        [InlineData("CZ-2020", "WC001-Digits.docx", "WC001-Digits-Deleted-Paragraph.docx", 1)]
        [InlineData("CZ-2030", "WC001-Digits-Deleted-Paragraph.docx", "WC001-Digits.docx", 1)]
        [InlineData("CZ-2040", "WC002-Unmodified.docx", "WC002-DiffInMiddle.docx", 2)]
        [InlineData("CZ-2050", "WC002-Unmodified.docx", "WC002-DiffAtBeginning.docx", 2)]
        [InlineData("CZ-2060", "WC002-Unmodified.docx", "WC002-DeleteAtBeginning.docx", 1)]
        [InlineData("CZ-2070", "WC002-Unmodified.docx", "WC002-InsertAtBeginning.docx", 1)]
        [InlineData("CZ-2080", "WC002-Unmodified.docx", "WC002-InsertAtEnd.docx", 1)]
        [InlineData("CZ-2080", "WC002-Unmodified.docx", "WC002-DeleteAtEnd.docx", 1)]
        [InlineData("CZ-2100", "WC002-Unmodified.docx", "WC002-DeleteInMiddle.docx", 1)]
        [InlineData("CZ-2110", "WC002-Unmodified.docx", "WC002-InsertInMiddle.docx", 1)]
        [InlineData("CZ-2120", "WC002-DeleteInMiddle.docx", "WC002-Unmodified.docx", 1)]
        //[InlineData("CZ-2130", "WC004-Large.docx", "WC004-Large-Mod.docx", 2)]
        [InlineData("CZ-2140", "WC006-Table.docx", "WC006-Table-Delete-Row.docx", 1)]
        [InlineData("CZ-2150", "WC006-Table-Delete-Row.docx", "WC006-Table.docx", 1)]
        [InlineData("CZ-2160", "WC006-Table.docx", "WC006-Table-Delete-Contests-of-Row.docx", 2)]
        [InlineData("CZ-2170", "WC007-Unmodified.docx", "WC007-Longest-At-End.docx", 2)]
        [InlineData("CZ-2180", "WC007-Unmodified.docx", "WC007-Deleted-at-Beginning-of-Para.docx", 2)]
        [InlineData("CZ-2200", "WC007-Unmodified.docx", "WC007-Moved-into-Table.docx", 2)]
        [InlineData("CZ-2210", "WC009-Table-Unmodified.docx", "WC009-Table-Cell-1-1-Mod.docx", 1)]
        [InlineData("CZ-2220", "WC010-Para-Before-Table-Unmodified.docx", "WC010-Para-Before-Table-Mod.docx", 3)]
        [InlineData("CZ-2230", "WC011-Before.docx", "WC011-After.docx", 2)]
        [InlineData("CZ-2240", "WC012-Math-Before.docx", "WC012-Math-After.docx", 2)]
        [InlineData("CZ-2250", "WC013-Image-Before.docx", "WC013-Image-After.docx", 2)]
        [InlineData("CZ-2260", "WC013-Image-Before.docx", "WC013-Image-After2.docx", 2)]
        [InlineData("CZ-2270", "WC013-Image-Before2.docx", "WC013-Image-After2.docx", 2)]
        [InlineData("CZ-2280", "WC014-SmartArt-Before.docx", "WC014-SmartArt-After.docx", 2)]
        [InlineData("CZ-2300", "WC014-SmartArt-With-Image-Before.docx", "WC014-SmartArt-With-Image-After.docx", 2)]
        [InlineData("CZ-2310", "WC014-SmartArt-With-Image-Before.docx", "WC014-SmartArt-With-Image-Deleted-After.docx", 3)]
        [InlineData("CZ-2320", "WC014-SmartArt-With-Image-Before.docx", "WC014-SmartArt-With-Image-Deleted-After2.docx", 1)]
        [InlineData("CZ-2330", "WC015-Three-Paragraphs.docx", "WC015-Three-Paragraphs-After.docx", 3)]
        [InlineData("CZ-2340", "WC016-Para-Image-Para.docx", "WC016-Para-Image-Para-w-Deleted-Image.docx", 1)]
        [InlineData("CZ-2350", "WC017-Image.docx", "WC017-Image-After.docx", 3)]
        [InlineData("CZ-2360", "WC018-Field-Simple-Before.docx", "WC018-Field-Simple-After-1.docx", 2)]
        [InlineData("CZ-2370", "WC018-Field-Simple-Before.docx", "WC018-Field-Simple-After-2.docx", 3)]
        [InlineData("CZ-2380", "WC019-Hyperlink-Before.docx", "WC019-Hyperlink-After-1.docx", 3)]
        [InlineData("CZ-2400", "WC019-Hyperlink-Before.docx", "WC019-Hyperlink-After-2.docx", 5)]
        [InlineData("CZ-2410", "WC020-FootNote-Before.docx", "WC020-FootNote-After-1.docx", 3)]
        [InlineData("CZ-2420", "WC020-FootNote-Before.docx", "WC020-FootNote-After-2.docx", 5)]
        [InlineData("CZ-2430", "WC021-Math-Before-1.docx", "WC021-Math-After-1.docx", 9)]
        [InlineData("CZ-2440", "WC021-Math-Before-2.docx", "WC021-Math-After-2.docx", 6)]
        [InlineData("CZ-2450", "WC022-Image-Math-Para-Before.docx", "WC022-Image-Math-Para-After.docx", 22)]
        [InlineData("CZ-2460", "WC023-Table-4-Row-Image-Before.docx", "WC023-Table-4-Row-Image-After-Delete-1-Row.docx", 9)]
        [InlineData("CZ-2470", "WC024-Table-Before.docx", "WC024-Table-After.docx", 1)]
        [InlineData("CZ-2480", "WC024-Table-Before.docx", "WC024-Table-After2.docx", 7)]
        [InlineData("CZ-2500", "WC025-Simple-Table-Before.docx", "WC025-Simple-Table-After.docx", 4)]
        [InlineData("CZ-2510", "WC026-Long-Table-Before.docx", "WC026-Long-Table-After-1.docx", 2)]
        [InlineData("CZ-2520", "WC027-Twenty-Paras-Before.docx", "WC027-Twenty-Paras-After-1.docx", 2)]
        [InlineData("CZ-2530", "WC027-Twenty-Paras-After-1.docx", "WC027-Twenty-Paras-Before.docx", 2)]
        [InlineData("CZ-2540", "WC027-Twenty-Paras-Before.docx", "WC027-Twenty-Paras-After-2.docx", 4)]
        [InlineData("CZ-2550", "WC030-Image-Math-Before.docx", "WC030-Image-Math-After.docx", 2)]
        [InlineData("CZ-2560", "WC031-Two-Maths-Before.docx", "WC031-Two-Maths-After.docx", 4)]
        [InlineData("CZ-2570", "WC032-Para-with-Para-Props.docx", "WC032-Para-with-Para-Props-After.docx", 3)]
        [InlineData("CZ-2580", "WC033-Merged-Cells-Before.docx", "WC033-Merged-Cells-After1.docx", 2)]
        [InlineData("CZ-2600", "WC033-Merged-Cells-Before.docx", "WC033-Merged-Cells-After2.docx", 4)]
        [InlineData("CZ-2610", "WC034-Footnotes-Before.docx", "WC034-Footnotes-After1.docx", 1)]
        [InlineData("CZ-2620", "WC034-Footnotes-Before.docx", "WC034-Footnotes-After2.docx", 6)]
        [InlineData("CZ-2630", "WC034-Footnotes-Before.docx", "WC034-Footnotes-After3.docx", 3)]
        [InlineData("CZ-2640", "WC034-Footnotes-After3.docx", "WC034-Footnotes-Before.docx", 3)]
        [InlineData("CZ-2650", "WC035-Footnote-Before.docx", "WC035-Footnote-After.docx", 2)]
        [InlineData("CZ-2660", "WC035-Footnote-After.docx", "WC035-Footnote-Before.docx", 2)]
        [InlineData("CZ-2670", "WC036-Footnote-With-Table-Before.docx", "WC036-Footnote-With-Table-After.docx", 5)]
        [InlineData("CZ-2680", "WC036-Footnote-With-Table-After.docx", "WC036-Footnote-With-Table-Before.docx", 5)]
        [InlineData("CZ-2700", "WC034-Endnotes-Before.docx", "WC034-Endnotes-After1.docx", 1)]
        [InlineData("CZ-2710", "WC034-Endnotes-Before.docx", "WC034-Endnotes-After2.docx", 6)]
        [InlineData("CZ-2720", "WC034-Endnotes-Before.docx", "WC034-Endnotes-After3.docx", 8)]
        [InlineData("CZ-2730", "WC034-Endnotes-After3.docx", "WC034-Endnotes-Before.docx", 8)]
        [InlineData("CZ-2740", "WC035-Endnote-Before.docx", "WC035-Endnote-After.docx", 2)]
        [InlineData("CZ-2750", "WC035-Endnote-After.docx", "WC035-Endnote-Before.docx", 2)]
        [InlineData("CZ-2760", "WC036-Endnote-With-Table-Before.docx", "WC036-Endnote-With-Table-After.docx", 6)]
        [InlineData("CZ-2770", "WC036-Endnote-With-Table-After.docx", "WC036-Endnote-With-Table-Before.docx", 6)]
        [InlineData("CZ-2780", "WC038-Document-With-BR-Before.docx", "WC038-Document-With-BR-After.docx", 2)]
        [InlineData("CZ-2790", "RC001-Before.docx", "RC001-After1.docx", 2)]
        [InlineData("CZ-2800", "RC002-Image.docx", "RC002-Image-After1.docx", 1)]
        [InlineData("CZ-2810", "WC039-Break-In-Row.docx", "WC039-Break-In-Row-After1.docx", 1)]
        //[InlineData("CZ-2820", "", "", 0)]
        //[InlineData("CZ-2830", "", "", 0)]
        //[InlineData("CZ-2840", "", "", 0)]
        //[InlineData("CZ-2850", "", "", 0)]
        //[InlineData("CZ-2860", "", "", 0)]
        //[InlineData("CZ-2870", "", "", 0)]
        //[InlineData("CZ-2880", "", "", 0)]
        //[InlineData("CZ-2890", "", "", 0)]
        public void CZ002_Compare(string testId, string name1, string name2, int revisionCount)
        {
            FileInfo source1Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
            FileInfo source2Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

            var rootTempDir = TestUtil.TempDir;
            var thisTestTempDir = new DirectoryInfo(Path.Combine(rootTempDir.FullName, testId));
            if (!thisTestTempDir.Exists)
                thisTestTempDir.Create();
            var source1CopiedToDestDocx = new FileInfo(Path.Combine(thisTestTempDir.FullName, source1Docx.Name));
            var source2CopiedToDestDocx = new FileInfo(Path.Combine(thisTestTempDir.FullName, source2Docx.Name));
            if (!source1CopiedToDestDocx.Exists)
                File.Copy(source1Docx.FullName, source1CopiedToDestDocx.FullName);
            if (!source2CopiedToDestDocx.Exists)
                File.Copy(source2Docx.FullName, source2CopiedToDestDocx.FullName);

            /************************************************************************************************************************/

            if (m_OpenWord)
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
            settings.DebugTempFileDi = TestUtil.TempDir;
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
                            sb.Append("Error" + Environment.NewLine);
                            sb.Append(ind + "ErrorType: " + err.ErrorType.ToString() + Environment.NewLine);
                            sb.Append(ind + "Description: " + err.Description + Environment.NewLine);
                            sb.Append(ind + "Part: " + err.Part.Uri.ToString() + Environment.NewLine);
                            sb.Append(ind + "XPath: " + err.Path.XPath + Environment.NewLine);
                        }
                        var sbs = sb.ToString();
                        if (sbs != "")
                            Assert.True(false, sbs.ToString());
                    }
                }
            }

            /************************************************************************************************************************/

            if (m_OpenWord)
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
#if false
            WmlDocument revisionWml = new WmlDocument(docxWithRevisionsFi.FullName);
            var revisions = WmlComparer.GetRevisions(revisionWml, settings);
            Assert.Equal(revisionCount, revisions.Count());
#endif
        }
#endif

#if false
        [Theory]
        [InlineData("RC001-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC001-After1.docx</DocName>
                <Color>LightYellow</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
              <RcInfo>
                <DocName>RC001-After2.docx</DocName>
                <Color>LightPink</Color>
                <Revisor>From Fred</Revisor>
              </RcInfo>
            </Root>")]
        [InlineData("RC002-Image.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC002-Image-After1.docx</DocName>
                <Color>LightBlue</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [InlineData("RC002-Image-After1.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC002-Image.docx</DocName>
                <Color>LightBlue</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [InlineData("WC027-Twenty-Paras-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>WC027-Twenty-Paras-After-1.docx</DocName>
                <Color>LightBlue</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [InlineData("WC027-Twenty-Paras-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>WC027-Twenty-Paras-After-3.docx</DocName>
                <Color>LightBlue</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [InlineData("RC003-Multi-Paras.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC003-Multi-Paras-After.docx</DocName>
                <Color>LightBlue</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [InlineData("RC004-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC004-After1.docx</DocName>
                <Color>LightYellow</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
              <RcInfo>
                <DocName>RC004-After2.docx</DocName>
                <Color>LightPink</Color>
                <Revisor>From Fred</Revisor>
              </RcInfo>
            </Root>")]

        public void WC001_Consolidate(string originalName, string revisedDocumentsXml)
        {
            FileInfo originalDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, originalName));

            var originalCopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, originalDocx.Name));
            if (!originalCopiedToDestDocx.Exists)
                File.Copy(originalDocx.FullName, originalCopiedToDestDocx.FullName);

            var revisedDocumentsXElement = XElement.Parse(revisedDocumentsXml);
            var revisedDocumentsArray = revisedDocumentsXElement
                .Elements()
                .Select(z =>
                {
                    FileInfo revisedDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, z.Element("DocName").Value));
                    var revisedCopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, revisedDocx.Name));
                    if (!revisedCopiedToDestDocx.Exists)
                        File.Copy(revisedDocx.FullName, revisedCopiedToDestDocx.FullName);
                    return new WmlRevisedDocumentInfo()
                    {
                        RevisedDocument = new WmlDocument(revisedCopiedToDestDocx.FullName),
                        Color = ColorParser.FromName(z.Element("Color").Value),
                        Revisor = z.Element("Revisor").Value,
                    };
                })
                .ToList();

            var consolidatedDocxName = originalCopiedToDestDocx.Name.Replace(".docx", "-Consolidated.docx");
            var consolidatedDocumentFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, consolidatedDocxName));

            WmlDocument source1Wml = new WmlDocument(originalCopiedToDestDocx.FullName);
            WmlComparerSettings settings = new WmlComparerSettings();
            WmlDocument consolidatedWml = WmlComparer.Consolidate(
                source1Wml,
                revisedDocumentsArray,
                settings);
            consolidatedWml.SaveAs(consolidatedDocumentFi.FullName);

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
                        var sbs = sb.ToString();
                        Assert.Equal("", sbs);
                    }
                }
            }

            /************************************************************************************************************************/

            if (s_OpenWord)
            {
                FileInfo wordExe = new FileInfo(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");
                WordRunner.RunWord(wordExe, consolidatedDocumentFi);
            }

            /************************************************************************************************************************/
        }

        [Theory]
        [InlineData("CA001-Plain.docx", "CA001-Plain-Mod.docx")]
        [InlineData("WC001-Digits.docx", "WC001-Digits-Mod.docx")]
        [InlineData("WC001-Digits.docx", "WC001-Digits-Deleted-Paragraph.docx")]
        [InlineData("WC001-Digits-Deleted-Paragraph.docx", "WC001-Digits.docx")]
        [InlineData("WC002-Unmodified.docx", "WC002-DiffInMiddle.docx")]
        [InlineData("WC002-Unmodified.docx", "WC002-DiffAtBeginning.docx")]
        [InlineData("WC002-Unmodified.docx", "WC002-DeleteAtBeginning.docx")]
        [InlineData("WC002-Unmodified.docx", "WC002-InsertAtBeginning.docx")]
        [InlineData("WC002-Unmodified.docx", "WC002-InsertAtEnd.docx")]
        [InlineData("WC002-Unmodified.docx", "WC002-DeleteAtEnd.docx")]
        [InlineData("WC002-Unmodified.docx", "WC002-DeleteInMiddle.docx")]
        [InlineData("WC002-Unmodified.docx", "WC002-InsertInMiddle.docx")]
        [InlineData("WC002-DeleteInMiddle.docx", "WC002-Unmodified.docx")]
        //[InlineData("WC004-Large.docx", "WC004-Large-Mod.docx")]
        [InlineData("WC006-Table.docx", "WC006-Table-Delete-Row.docx")]
        [InlineData("WC006-Table-Delete-Row.docx", "WC006-Table.docx")]
        [InlineData("WC006-Table.docx", "WC006-Table-Delete-Contests-of-Row.docx")]
        [InlineData("WC007-Unmodified.docx", "WC007-Longest-At-End.docx")]
        [InlineData("WC007-Unmodified.docx", "WC007-Deleted-at-Beginning-of-Para.docx")]
        [InlineData("WC007-Unmodified.docx", "WC007-Moved-into-Table.docx")]
        [InlineData("WC009-Table-Unmodified.docx", "WC009-Table-Cell-1-1-Mod.docx")]
        [InlineData("WC010-Para-Before-Table-Unmodified.docx", "WC010-Para-Before-Table-Mod.docx")]
        [InlineData("WC011-Before.docx", "WC011-After.docx")]
        [InlineData("WC012-Math-Before.docx", "WC012-Math-After.docx")]
        [InlineData("WC013-Image-Before.docx", "WC013-Image-After.docx")]
        [InlineData("WC013-Image-Before.docx", "WC013-Image-After2.docx")]
        [InlineData("WC013-Image-Before2.docx", "WC013-Image-After2.docx")]
        [InlineData("WC014-SmartArt-Before.docx", "WC014-SmartArt-After.docx")]
        [InlineData("WC014-SmartArt-With-Image-Before.docx", "WC014-SmartArt-With-Image-After.docx")]
        [InlineData("WC014-SmartArt-With-Image-Before.docx", "WC014-SmartArt-With-Image-Deleted-After.docx")]
        [InlineData("WC014-SmartArt-With-Image-Before.docx", "WC014-SmartArt-With-Image-Deleted-After2.docx")]
        [InlineData("WC015-Three-Paragraphs.docx", "WC015-Three-Paragraphs-After.docx")]
        [InlineData("WC016-Para-Image-Para.docx", "WC016-Para-Image-Para-w-Deleted-Image.docx")]
        [InlineData("WC017-Image.docx", "WC017-Image-After.docx")]
        [InlineData("WC018-Field-Simple-Before.docx", "WC018-Field-Simple-After-1.docx")]
        [InlineData("WC018-Field-Simple-Before.docx", "WC018-Field-Simple-After-2.docx")]
        [InlineData("WC019-Hyperlink-Before.docx", "WC019-Hyperlink-After-1.docx")]
        [InlineData("WC019-Hyperlink-Before.docx", "WC019-Hyperlink-After-2.docx")]
        [InlineData("WC020-FootNote-Before.docx", "WC020-FootNote-After-1.docx")]
        [InlineData("WC020-FootNote-Before.docx", "WC020-FootNote-After-2.docx")]
        [InlineData("WC021-Math-Before-1.docx", "WC021-Math-After-1.docx")]
        [InlineData("WC021-Math-Before-2.docx", "WC021-Math-After-2.docx")]
        [InlineData("WC022-Image-Math-Para-Before.docx", "WC022-Image-Math-Para-After.docx")]
        [InlineData("WC023-Table-4-Row-Image-Before.docx", "WC023-Table-4-Row-Image-After-Delete-1-Row.docx")]
        [InlineData("WC024-Table-Before.docx", "WC024-Table-After.docx")]
        [InlineData("WC024-Table-Before.docx", "WC024-Table-After2.docx")]
        [InlineData("WC025-Simple-Table-Before.docx", "WC025-Simple-Table-After.docx")]
        [InlineData("WC026-Long-Table-Before.docx", "WC026-Long-Table-After-1.docx")]
        [InlineData("WC027-Twenty-Paras-Before.docx", "WC027-Twenty-Paras-After-1.docx")]
        [InlineData("WC027-Twenty-Paras-After-1.docx", "WC027-Twenty-Paras-Before.docx")]
        [InlineData("WC027-Twenty-Paras-Before.docx", "WC027-Twenty-Paras-After-2.docx")]
        [InlineData("WC030-Image-Math-Before.docx", "WC030-Image-Math-After.docx")]
        [InlineData("WC031-Two-Maths-Before.docx", "WC031-Two-Maths-After.docx")]
        [InlineData("WC032-Para-with-Para-Props.docx", "WC032-Para-with-Para-Props-After.docx")]
        [InlineData("WC033-Merged-Cells-Before.docx", "WC033-Merged-Cells-After1.docx")]
        [InlineData("WC033-Merged-Cells-Before.docx", "WC033-Merged-Cells-After2.docx")]
        [InlineData("WC034-Footnotes-Before.docx", "WC034-Footnotes-After1.docx")]
        [InlineData("WC034-Footnotes-Before.docx", "WC034-Footnotes-After2.docx")]
        [InlineData("WC034-Footnotes-Before.docx", "WC034-Footnotes-After3.docx")]
        [InlineData("WC034-Footnotes-After3.docx", "WC034-Footnotes-Before.docx")]
        [InlineData("WC035-Footnote-Before.docx", "WC035-Footnote-After.docx")]
        [InlineData("WC035-Footnote-After.docx", "WC035-Footnote-Before.docx")]
        [InlineData("WC036-Footnote-With-Table-Before.docx", "WC036-Footnote-With-Table-After.docx")]
        [InlineData("WC036-Footnote-With-Table-After.docx", "WC036-Footnote-With-Table-Before.docx")]
        [InlineData("WC034-Endnotes-Before.docx", "WC034-Endnotes-After1.docx")]
        [InlineData("WC034-Endnotes-Before.docx", "WC034-Endnotes-After2.docx")]
        [InlineData("WC034-Endnotes-Before.docx", "WC034-Endnotes-After3.docx")]
        [InlineData("WC034-Endnotes-After3.docx", "WC034-Endnotes-Before.docx")]
        [InlineData("WC035-Endnote-Before.docx", "WC035-Endnote-After.docx")]
        [InlineData("WC035-Endnote-After.docx", "WC035-Endnote-Before.docx")]
        [InlineData("WC036-Endnote-With-Table-Before.docx", "WC036-Endnote-With-Table-After.docx")]
        [InlineData("WC036-Endnote-With-Table-After.docx", "WC036-Endnote-With-Table-Before.docx")]
        [InlineData("WC038-Document-With-BR-Before.docx", "WC038-Document-With-BR-After.docx")]
        [InlineData("RC001-Before.docx", "RC001-After1.docx")]
        [InlineData("RC002-Image.docx", "RC002-Image-After1.docx")]
        //[InlineData("", "")]
        //[InlineData("", "")]
        //[InlineData("", "")]
        //[InlineData("", "")]
        //[InlineData("", "")]
        //[InlineData("", "")]
        //[InlineData("", "")]
        //[InlineData("", "")]
        //[InlineData("", "")]


        public void WC002_Consolidate_Bulk_Test(string name1, string name2)
        {
            FileInfo source1Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
            FileInfo source2Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

            var source1CopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, source1Docx.Name));
            var source2CopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, source2Docx.Name));
            if (!source1CopiedToDestDocx.Exists)
                File.Copy(source1Docx.FullName, source1CopiedToDestDocx.FullName);
            if (!source2CopiedToDestDocx.Exists)
                File.Copy(source2Docx.FullName, source2CopiedToDestDocx.FullName);

            /************************************************************************************************************************/

            if (s_OpenWord)
            {
                FileInfo source1DocxForWord = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
                FileInfo source2DocxForWord = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

                var source1CopiedToDestDocxForWord = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, source1Docx.Name.Replace(".docx", "-For-Word.docx")));
                var source2CopiedToDestDocxForWord = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, source2Docx.Name.Replace(".docx", "-For-Word.docx")));
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
            var docxWithRevisionsFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, before + "-COMPARE-" + after + ".docx"));
            var docxConsolidatedFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, before + "-CONSOLIDATED-" + after + ".docx"));

            WmlDocument source1Wml = new WmlDocument(source1CopiedToDestDocx.FullName);
            WmlDocument source2Wml = new WmlDocument(source2CopiedToDestDocx.FullName);
            WmlComparerSettings settings = new WmlComparerSettings();
            WmlDocument comparedWml = WmlComparer.Compare(source1Wml, source2Wml, settings);
            comparedWml.SaveAs(docxWithRevisionsFi.FullName);

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
            consolidatedWml.SaveAs(docxConsolidatedFi.FullName);

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
                        var sbs = sb.ToString();
                        Assert.Equal("", sbs);
                    }
                }
            }

            /************************************************************************************************************************/

            if (s_OpenWord)
            {
                FileInfo wordExe = new FileInfo(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");
                WordRunner.RunWord(wordExe, docxConsolidatedFi);
            }

            /************************************************************************************************************************/
        }
#endif

#if false
        [Theory]
        [InlineData("WC037-Textbox-Before.docx", "WC037-Textbox-After1.docx", 2)]

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

        [Theory]
        [InlineData("WC001-Digits.docx")]
        [InlineData("WC001-Digits-Deleted-Paragraph.docx")]
        [InlineData("WC001-Digits-Mod.docx")]
        [InlineData("WC002-DeleteAtBeginning.docx")]
        [InlineData("WC002-DeleteAtEnd.docx")]
        [InlineData("WC002-DeleteInMiddle.docx")]
        [InlineData("WC002-DiffAtBeginning.docx")]
        [InlineData("WC002-DiffInMiddle.docx")]
        [InlineData("WC002-InsertAtBeginning.docx")]
        [InlineData("WC002-InsertAtEnd.docx")]
        [InlineData("WC002-InsertInMiddle.docx")]
        [InlineData("WC002-Unmodified.docx")]
        //[InlineData("WC004-Large.docx")]
        //[InlineData("WC004-Large-Mod.docx")]
        [InlineData("WC006-Table.docx")]
        [InlineData("WC006-Table-Delete-Contests-of-Row.docx")]
        [InlineData("WC006-Table-Delete-Row.docx")]
        [InlineData("WC007-Deleted-at-Beginning-of-Para.docx")]
        [InlineData("WC007-Longest-At-End.docx")]
        [InlineData("WC007-Moved-into-Table.docx")]
        [InlineData("WC007-Unmodified.docx")]
        [InlineData("WC009-Table-Cell-1-1-Mod.docx")]
        [InlineData("WC009-Table-Unmodified.docx")]
        [InlineData("WC010-Para-Before-Table-Mod.docx")]
        [InlineData("WC010-Para-Before-Table-Unmodified.docx")]
        [InlineData("WC011-After.docx")]
        [InlineData("WC011-Before.docx")]
        [InlineData("WC012-Math-After.docx")]
        [InlineData("WC012-Math-Before.docx")]
        [InlineData("WC013-Image-After.docx")]
        [InlineData("WC013-Image-After2.docx")]
        [InlineData("WC013-Image-Before.docx")]
        [InlineData("WC013-Image-Before2.docx")]
        [InlineData("WC014-SmartArt-After.docx")]
        [InlineData("WC014-SmartArt-Before.docx")]
        [InlineData("WC014-SmartArt-With-Image-After.docx")]
        [InlineData("WC014-SmartArt-With-Image-Before.docx")]
        [InlineData("WC014-SmartArt-With-Image-Deleted-After.docx")]
        [InlineData("WC014-SmartArt-With-Image-Deleted-After2.docx")]
        [InlineData("WC015-Three-Paragraphs.docx")]
        [InlineData("WC015-Three-Paragraphs-After.docx")]
        [InlineData("WC016-Para-Image-Para.docx")]
        [InlineData("WC016-Para-Image-Para-w-Deleted-Image.docx")]
        [InlineData("WC017-Image.docx")]
        [InlineData("WC017-Image-After.docx")]
        [InlineData("WC018-Field-Simple-After-1.docx")]
        [InlineData("WC018-Field-Simple-After-2.docx")]
        [InlineData("WC018-Field-Simple-Before.docx")]
        [InlineData("WC019-Hyperlink-After-1.docx")]
        [InlineData("WC019-Hyperlink-After-2.docx")]
        [InlineData("WC019-Hyperlink-Before.docx")]
        [InlineData("WC020-FootNote-After-1.docx")]
        [InlineData("WC020-FootNote-After-2.docx")]
        [InlineData("WC020-FootNote-Before.docx")]
        [InlineData("WC021-Math-After-1.docx")]
        [InlineData("WC021-Math-Before-1.docx")]
        [InlineData("WC022-Image-Math-Para-After.docx")]
        [InlineData("WC022-Image-Math-Para-Before.docx")]
        //[InlineData("", "")]
        //[InlineData("", "")]
        //[InlineData("", "")]
        //[InlineData("", "")]

        public void WC004_Compare_To_Self(string name)
        {
            FileInfo sourceDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            var sourceCopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-Source.docx")));
            if (!sourceCopiedToDestDocx.Exists)
                File.Copy(sourceDocx.FullName, sourceCopiedToDestDocx.FullName);

            var before = sourceCopiedToDestDocx.Name.Replace(".docx", "");
            var docxComparedFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, before + "-COMPARE" + ".docx"));
            var docxCompared2Fi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, before + "-COMPARE2" + ".docx"));

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
        [InlineData("WC040-Case-Before.docx", "WC040-Case-After.docx", 2)]
        //[InlineData("", "", 0)]
        //[InlineData("", "", 0)]
        //[InlineData("", "", 0)]
        //[InlineData("", "", 0)]
        //[InlineData("", "", 0)]
        //[InlineData("", "", 0)]
        //[InlineData("", "", 0)]
        //[InlineData("", "", 0)]

        public void WC005_Compare_CaseInsensitive(string name1, string name2, int revisionCount)
        {
            FileInfo source1Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
            FileInfo source2Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

            var source1CopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, source1Docx.Name));
            var source2CopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, source2Docx.Name));
            if (!source1CopiedToDestDocx.Exists)
                File.Copy(source1Docx.FullName, source1CopiedToDestDocx.FullName);
            if (!source2CopiedToDestDocx.Exists)
                File.Copy(source2Docx.FullName, source2CopiedToDestDocx.FullName);

            /************************************************************************************************************************/

            if (s_OpenWord)
            {
                FileInfo source1DocxForWord = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
                FileInfo source2DocxForWord = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

                var source1CopiedToDestDocxForWord = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, source1Docx.Name.Replace(".docx", "-For-Word.docx")));
                var source2CopiedToDestDocxForWord = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, source2Docx.Name.Replace(".docx", "-For-Word.docx")));
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
            var docxWithRevisionsFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, before + "-COMPARE-" + after + ".docx"));

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
                            sb.Append("Error" + Environment.NewLine);
                            sb.Append(ind + "ErrorType: " + err.ErrorType.ToString() + Environment.NewLine);
                            sb.Append(ind + "Description: " + err.Description + Environment.NewLine);
                            sb.Append(ind + "Part: " + err.Part.Uri.ToString() + Environment.NewLine);
                            sb.Append(ind + "XPath: " + err.Path.XPath + Environment.NewLine);
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
#endif

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
                            sb.Append("Error" + Environment.NewLine);
                            sb.Append(ind + "ErrorType: " + err.ErrorType.ToString() + Environment.NewLine);
                            sb.Append(ind + "Description: " + err.Description + Environment.NewLine);
                            sb.Append(ind + "Part: " + err.Part.Uri.ToString() + Environment.NewLine);
                            sb.Append(ind + "XPath: " + err.Path.XPath + Environment.NewLine);
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
        };

    }
#if false
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
#endif
}

#endif
