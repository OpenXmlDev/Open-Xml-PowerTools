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

#define COPY_FILES_FOR_DEBUGGING

// DO_CONVERSION_VIA_WORD is defined in the project OpenXmlPowerTools.Tests.OA.csproj, but not in the OpenXmlPowerTools.Tests.csproj

using System;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using Xunit;

#if DO_CONVERSION_VIA_WORD
using Word = Microsoft.Office.Interop.Word;
#endif

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public static class DhSettings
    {
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static bool m_OpenTempDirInExplorer = true;
        public static bool m_CopySourceFiles = true;
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    }

    public class Dh
    {
        public static bool s_CopySourceFiles = true;
        public static bool s_CopyFormattingAssembledDocx = true;
        public static bool s_ConvertUsingWord = true;

        // PowerShell oneliner that generates InlineData for all files in a directory
        // dir | % { '[InlineData("' + $_.Name + '")]' } | clip

        [Theory]
        [InlineData("DH-0010", "DH/DH001-5DayTourPlanTemplate.docx", false)]
        [InlineData("DH-0020", "DH/DH002-Hebrew-01.docx", false)]
        [InlineData("DH-0030", "DH/DH003-Hebrew-02.docx", false)]
        [InlineData("DH-0040", "DH/DH004-ResumeTemplate.docx", false)]
        [InlineData("DH-0050", "DH/DH005-TaskPlanTemplate.docx", false)]
        [InlineData("DH-0060", "DH/DH006-Test-01.docx", false)]
        [InlineData("DH-0070", "DH/DH007-Test-02.docx", false)]
        [InlineData("DH-0080", "DH/DH008-Test-03.docx", false)]
        [InlineData("DH-0090", "DH/DH009-Test-04.docx", false)]
        [InlineData("DH-0100", "DH/DH010-Test-05.docx", false)]
        [InlineData("DH-0110", "DH/DH011-Test-06.docx", false)]
        [InlineData("DH-0120", "DH/DH012-Test-07.docx", false)]
        [InlineData("DH-0130", "DH/DH013-Test-08.docx", false)]
        [InlineData("DH-0140", "DH/DH014-RTL-Table-01.docx", false)]
        [InlineData("DH-0150", "DH/DH015-Vertical-Spacing-atLeast.docx", false)]
        [InlineData("DH-0160", "DH/DH016-Horizontal-Spacing-firstLine.docx", false)]
        [InlineData("DH-0170", "DH/DH017-Vertical-Alignment-Cell-01.docx", false)]
        [InlineData("DH-0180", "DH/DH018-Vertical-Alignment-Para-01.docx", false)]
        [InlineData("DH-0190", "DH/DH019-Hidden-Run.docx", false)]
        [InlineData("DH-0200", "DH/DH020-Small-Caps.docx", false)]
        [InlineData("DH-0210", "DH/DH021-Symbols.docx", false)]
        [InlineData("DH-0220", "DH/DH022-Table-Of-Contents.docx", false)]
        [InlineData("DH-0230", "DH/DH023-Hyperlink.docx", false)]
        [InlineData("DH-0240", "DH/DH024-Tabs-01.docx", false)]
        [InlineData("DH-0250", "DH/DH025-Tabs-02.docx", false)]
        [InlineData("DH-0260", "DH/DH026-Tabs-03.docx", false)]
        [InlineData("DH-0270", "DH/DH027-Tabs-04.docx", false)]
        [InlineData("DH-0280", "DH/DH028-No-Break-Hyphen.docx", false)]
        [InlineData("DH-0290", "DH/DH029-Table-Merged-Cells.docx", false)]
        [InlineData("DH-0300", "DH/DH030-Content-Controls.docx", false)]
        [InlineData("DH-0310", "DH/DH031-Complicated-Document.docx", false)]
        [InlineData("DH-0320", "DH/DH032-Named-Color.docx", false)]
        [InlineData("DH-0330", "DH/DH033-Run-With-Border.docx", false)]
        [InlineData("DH-0340", "DH/DH034-Run-With-Position.docx", false)]
        [InlineData("DH-0350", "DH/DH035-Strike-Through.docx", false)]
        [InlineData("DH-0360", "DH/DH036-Super-Script.docx", false)]
        [InlineData("DH-0370", "DH/DH037-Sub-Script.docx", false)]
        [InlineData("DH-0380", "DH/DH038-Conflicting-Border-Weight.docx", false)]
        [InlineData("DH-0390", "DH/DH039-Bold.docx", false)]
        [InlineData("DH-0400", "DH/DH040-Hyperlink-Fieldcode-01.docx", false)]
        [InlineData("DH-0410", "DH/DH041-Hyperlink-Fieldcode-02.docx", false)]
        [InlineData("DH-0420", "DH/DH042-Image-Png.docx", false)]
        [InlineData("DH-0430", "DH/DH043-Chart.docx", false)]
        [InlineData("DH-0440", "DH/DH044-Embedded-Workbook.docx", false)]
        [InlineData("DH-0450", "DH/DH045-Italic.docx", false)]
        [InlineData("DH-0460", "DH/DH046-BoldAndItalic.docx", false)]
        [InlineData("DH-0470", "DH/DH047-No-Section.docx", false)]
        [InlineData("DH-0480", "DH/DH048-Excerpt.docx", false)]
        [InlineData("DH-0490", "DH/DH049-Borders.docx", false)]
        [InlineData("DH-0500", "DH/DH050-Shaded-Text-01.docx", false)]
        [InlineData("DH-0510", "DH/DH051-Shaded-Text-02.docx", false)]
        [InlineData("DH-0520", "DH/DH060-Image-with-Hyperlink.docx", false)]
        [InlineData("DH-0530", "DH/DH061-Hyperlink-in-Field.docx", false)]
        [InlineData("DH-0540", "DH/DH062-Deleted-Text.docx", true)]
        [InlineData("DH-0550", "DH/DH063-Inserted-Text.docx", true)]
        [InlineData("DH-0560", "DH/DH064-Deleted-Table-Row.docx", true)]
        [InlineData("DH-0570", "DH/DH065-inserted-Table-Row.docx", true)]
        [InlineData("DH-0580", "DH/DH066-MoveFrom-MoveTo.docx", true)]
        [InlineData("DH-0590", "DH/DH067-Table-Cell-Deletion.docx", true)]
        [InlineData("DH-0600", "DH/DH068-Table-Cell-Insertion.docx", true)]
        [InlineData("DH-0610", "DH/DH069-Table-Cell-Merge.docx", true)]
        [InlineData("DH-0620", "DH/DH070-Content-Control-Insertion-Deletion.docx", true)]
        [InlineData("DH-0630", "DH/DH071-Deleted-Field.docx", true)]
        [InlineData("DH-0640", "DH/DH072-Inserted-Numbering-Properties.docx", true)]
        [InlineData("DH-0650", "DH/DH073-Previous-Numbering-Field-Props.docx", true)]
        [InlineData("DH-0660", "DH/DH076-Run-Props-Change-on-Para-Mark.docx", true)]
        [InlineData("DH-0670", "DH/DH077-Section-Props-Change.docx", true)]
        [InlineData("DH-0680", "DH/DH078-Table-Grid-Change.docx", true)]
        [InlineData("DH-0690", "DH/DH079-Table-Properties-Change.docx", true)]
        [InlineData("DH-0700", "DH/DH080-Table-Cell-Props-Change.docx", true)]
        [InlineData("DH-0710", "DH/DH081-Table-Row-Props-Change.docx", true)]
        [InlineData("DH-0720", "DH/DH074-Paragraph-Props-Change.docx", true)]
        [InlineData("DH-0730", "DH/DH075-Run-Props-Change.docx", true)]
        [InlineData("DH-0740", "DH/DH082-Comment.docx", true)]
        [InlineData("DH-0750", "DH/DH083-Comments-Test.docx", true)]
        [InlineData("DH-0760", "DH/DH084-Comments-Larger-Doc.docx", true)]
        [InlineData("DH-0770", "DH/DH085-Mult-Comments.docx", true)]
        [InlineData("DH-0780", "DH/DH086-Footnote.docx", true)]
        [InlineData("DH-0790", "DH/DH087-Endnote.docx", true)]
        [InlineData("DH-0800", "DH/DH088-FootAndEndNote.docx", true)]
        
        public void DocxToHtml(string testId, string name, bool displayRevisionTracking)
        {
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create the dir for the test
            DirectoryInfo thisTestTempDir;
            string tempDirFullName = TestUtil.CreateTestDir(testId, out thisTestTempDir);

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Load DOCX
            FileInfo docxFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));
            var testSubDir = name.Split('/')[0];
            FileInfo testBaselineHtmlFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, testSubDir, "TestBaselineFiles", testId + "-3-OxPt.html"));
            var copiedTestBaselineHtmlFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, testId, docxFi.Name.Replace(".docx", "-2-OxPt.html")));

            WmlDocument wmlDocument = new WmlDocument(docxFi.FullName);

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Copy source files
            CopySourceFilesToTestDir(tempDirFullName, docxFi);
            File.Copy(testBaselineHtmlFi.FullName, copiedTestBaselineHtmlFi.FullName);

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Do the DOCX conversion
            var oxPtConvertedDestHtmlFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, testId, docxFi.Name.Replace(".docx", "-3-OxPt.html")));
            ConvertToHtml(docxFi, oxPtConvertedDestHtmlFi, displayRevisionTracking);

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // create batch file to copy newly generated documents to the TestFiles directory.
            TestUtil.AddToBatchFile(testBaselineHtmlFi, oxPtConvertedDestHtmlFi);

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Open Windows Explorer
            if (DhSettings.m_OpenTempDirInExplorer)
                TestUtil.OpenWindowsExplorer(thisTestTempDir);

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Do assertions
            CompareAgainstBaseline(testBaselineHtmlFi, oxPtConvertedDestHtmlFi);

#if false
            FileInfo sourceDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            var sourceCopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-1-Source.docx")));
            if (!sourceCopiedToDestDocx.Exists)
                File.Copy(sourceDocx.FullName, sourceCopiedToDestDocx.FullName);

            var assembledFormattingDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-2-FormattingAssembled.docx")));
            if (!assembledFormattingDestDocx.Exists)
                CopyFormattingAssembledDocx(sourceDocx, assembledFormattingDestDocx);

            var oxPtConvertedDestHtml = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-3-OxPt.html")));
            ConvertToHtml(sourceDocx, oxPtConvertedDestHtml);

#if DO_CONVERSION_VIA_WORD
            var wordConvertedDocHtml = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-4-Word.html")));
            ConvertToHtmlUsingWord(sourceDocx, wordConvertedDocHtml);
#endif
#endif
        }

        [Theory]
        [InlineData("DH-2000", "DH/DH006-Test-01.docx")]
        public void DocxToHtmlNoCssClasses(string testId, string name)
        {
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create the dir for the test
            DirectoryInfo thisTestTempDir;
            string tempDirFullName = TestUtil.CreateTestDir(testId, out thisTestTempDir);

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Load DOCX
            FileInfo docxFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));
            var testSubDir = name.Split('/')[0];
            FileInfo testBaselineHtmlFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, testSubDir, "TestBaselineFiles", testId + "-3-OxPt.html"));

            WmlDocument wmlDocument = new WmlDocument(docxFi.FullName);

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Copy source files
            CopySourceFilesToTestDir(tempDirFullName, docxFi);

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Do the DOCX conversion
            var oxPtConvertedDestHtmlFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, testId, docxFi.Name.Replace(".docx", "-3-OxPt.html")));
            ConvertToHtmlNoCssClasses(docxFi, oxPtConvertedDestHtmlFi);

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // create batch file to copy newly generated documents to the TestFiles directory.
            TestUtil.AddToBatchFile(testBaselineHtmlFi, oxPtConvertedDestHtmlFi);

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Open Windows Explorer
            if (DhSettings.m_OpenTempDirInExplorer)
                TestUtil.OpenWindowsExplorer(thisTestTempDir);

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Do assertions
            CompareAgainstBaseline(testBaselineHtmlFi, oxPtConvertedDestHtmlFi);

#if false
            FileInfo sourceDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            var oxPtConvertedDestHtml = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-5-OxPt-No-CSS-Classes.html")));
            ConvertToHtmlNoCssClasses(sourceDocx, oxPtConvertedDestHtml);
#endif
        }

        private void CompareAgainstBaseline(FileInfo testBaselineHtmlFi, FileInfo oxPtConvertedDestHtmlFi)
        {
            if (!testBaselineHtmlFi.Exists)
                Assert.True(false, "No baseline html file");

            XElement baseline = LocalNormalizeForComparison(XElement.Load(testBaselineHtmlFi.FullName));
            XElement converted = LocalNormalizeForComparison(XElement.Load(oxPtConvertedDestHtmlFi.FullName));
            var failed = !XNode.DeepEquals(baseline, converted);
            if (failed)
            {
                Assert.True(false, "DocxToHtml regression error");
            }
        }

        //static Regex s_DeleteWidthExpression = null;

        private XElement LocalNormalizeForComparison(XElement xElement)
        {
            xElement.Descendants(Xhtml.style).Remove();
            xElement.Descendants().Attributes("class").Remove();
            return xElement;

            //if (s_DeleteWidthExpression == null)
            //    s_DeleteWidthExpression = new Regex("width:[\\s0-9\\.]+in;");
            //var str = xElement.ToString();
            //var str2 = s_DeleteWidthExpression.Replace(str, "");
            //return XElement.Parse(str2);
        }


        private static void CopySourceFilesToTestDir(string tempDirFullName, FileInfo docxFi)
        {
            if (DhSettings.m_CopySourceFiles)
            {
                var docxCopiedFi = new FileInfo(Path.Combine(tempDirFullName, docxFi.Name));
                File.Copy(docxFi.FullName, docxCopiedFi.FullName);
            }
        }

        private static void CopyFormattingAssembledDocx(FileInfo source, FileInfo dest)
        {
            var ba = File.ReadAllBytes(source.FullName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(ba, 0, ba.Length);
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(ms, true))
                {

                    RevisionAccepter.AcceptRevisions(wordDoc);
                    SimplifyMarkupSettings simplifyMarkupSettings = new SimplifyMarkupSettings
                    {
                        RemoveComments = true,
                        RemoveContentControls = true,
                        RemoveEndAndFootNotes = true,
                        RemoveFieldCodes = false,
                        RemoveLastRenderedPageBreak = true,
                        RemovePermissions = true,
                        RemoveProof = true,
                        RemoveRsidInfo = true,
                        RemoveSmartTags = true,
                        RemoveSoftHyphens = true,
                        RemoveGoBackBookmark = true,
                        ReplaceTabsWithSpaces = false,
                    };
                    MarkupSimplifier.SimplifyMarkup(wordDoc, simplifyMarkupSettings);

                    FormattingAssemblerSettings formattingAssemblerSettings = new FormattingAssemblerSettings
                    {
                        RemoveStyleNamesFromParagraphAndRunProperties = false,
                        ClearStyles = false,
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        CreateHtmlConverterAnnotationAttributes = true,
                        OrderElementsPerStandard = false,
                        ListItemRetrieverSettings =
                            new ListItemRetrieverSettings()
                            {
                                ListItemTextImplementations = ListItemRetrieverSettings.DefaultListItemTextImplementations,
                            },
                    };

                    FormattingAssembler.AssembleFormatting(wordDoc, formattingAssemblerSettings);
                }
                var newBa = ms.ToArray();
                File.WriteAllBytes(dest.FullName, newBa);
            }
        }

        private static void ConvertToHtml(FileInfo sourceDocx, FileInfo destFileName, bool displayRevisionTracking)
        {
            byte[] byteArray = File.ReadAllBytes(sourceDocx.FullName);
            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(memoryStream, true))
                {
                    var outputDirectory = destFileName.Directory;
                    destFileName = new FileInfo(Path.Combine(outputDirectory.FullName, destFileName.Name));
                    var imageDirectoryName = destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";
                    int imageCounter = 0;
                    var pageTitle = (string)wDoc.CoreFilePropertiesPart.GetXDocument().Descendants(DC.title).FirstOrDefault();
                    if (pageTitle == null)
                        pageTitle = sourceDocx.Name;

                    DocxToHtmlSettings settings = new DocxToHtmlSettings()
                    {
                        PageTitle = pageTitle,
                        DisplayRevisionTracking = displayRevisionTracking,
                        FabricateCssClasses = true,
                        CssClassPrefix = "pt-",
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        ImageHandler = imageInfo =>
                        {
                            DirectoryInfo localDirInfo = new DirectoryInfo(imageDirectoryName);
                            if (!localDirInfo.Exists)
                                localDirInfo.Create();
                            ++imageCounter;
                            string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                            ImageFormat imageFormat = null;
                            if (extension == "png")
                            {
                                // Convert png to jpeg.
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "gif")
                                imageFormat = ImageFormat.Gif;
                            else if (extension == "bmp")
                                imageFormat = ImageFormat.Bmp;
                            else if (extension == "jpeg")
                                imageFormat = ImageFormat.Jpeg;
                            else if (extension == "tiff")
                            {
                                // Convert tiff to gif.
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "x-wmf")
                            {
                                extension = "wmf";
                                imageFormat = ImageFormat.Wmf;
                            }

                            // If the image format isn't one that we expect, ignore it,
                            // and don't return markup for the link.
                            if (imageFormat == null)
                                return null;

                            string imageSaveLocation = imageDirectoryName + "/image" + imageCounter.ToString() + "." + extension;
                            string imageRelativeFileName = localDirInfo.Name + "/image" + imageCounter.ToString() + "." + extension;
                            try
                            {
                                imageInfo.Bitmap.Save(imageSaveLocation, imageFormat);
                            }
                            catch (System.Runtime.InteropServices.ExternalException)
                            {
                                return null;
                            }
                            XElement img = new XElement(Xhtml.img,
                                new XAttribute(NoNamespace.src, imageRelativeFileName),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null ?
                                    new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                            return img;
                        }
                    };
                    XElement html = OpenXmlPowerTools.DocxToHtml.ConvertToHtml(wDoc, settings);

                    // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
                    // XEntity.  PtOpenXmlUtil.cs define the XEntity class.  See
                    // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
                    // for detailed explanation.
                    //
                    // If you further transform the XML tree returned by ConvertToHtmlTransform, you
                    // must do it correctly, or entities will not be serialized properly.

                    var htmlString = html.ToString(SaveOptions.DisableFormatting);
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                }
            }
        }

        private static void ConvertToHtmlNoCssClasses(FileInfo sourceDocx, FileInfo destFileName)
        {
            byte[] byteArray = File.ReadAllBytes(sourceDocx.FullName);
            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(memoryStream, true))
                {
                    var outputDirectory = destFileName.Directory;
                    destFileName = new FileInfo(Path.Combine(outputDirectory.FullName, destFileName.Name));
                    var imageDirectoryName = destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";
                    int imageCounter = 0;
                    var pageTitle = (string)wDoc.CoreFilePropertiesPart.GetXDocument().Descendants(DC.title).FirstOrDefault();
                    if (pageTitle == null)
                        pageTitle = sourceDocx.FullName;

                    DocxToHtmlSettings settings = new DocxToHtmlSettings()
                    {
                        PageTitle = pageTitle,
                        FabricateCssClasses = false,
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        ImageHandler = imageInfo =>
                        {
                            DirectoryInfo localDirInfo = new DirectoryInfo(imageDirectoryName);
                            if (!localDirInfo.Exists)
                                localDirInfo.Create();
                            ++imageCounter;
                            string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                            ImageFormat imageFormat = null;
                            if (extension == "png")
                            {
                                // Convert png to jpeg.
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "gif")
                                imageFormat = ImageFormat.Gif;
                            else if (extension == "bmp")
                                imageFormat = ImageFormat.Bmp;
                            else if (extension == "jpeg")
                                imageFormat = ImageFormat.Jpeg;
                            else if (extension == "tiff")
                            {
                                // Convert tiff to gif.
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "x-wmf")
                            {
                                extension = "wmf";
                                imageFormat = ImageFormat.Wmf;
                            }

                            // If the image format isn't one that we expect, ignore it,
                            // and don't return markup for the link.
                            if (imageFormat == null)
                                return null;

                            string imageFileName = imageDirectoryName + "/image" +
                                imageCounter.ToString() + "." + extension;
                            try
                            {
                                imageInfo.Bitmap.Save(imageFileName, imageFormat);
                            }
                            catch (System.Runtime.InteropServices.ExternalException)
                            {
                                return null;
                            }
                            XElement img = new XElement(Xhtml.img,
                                new XAttribute(NoNamespace.src, imageFileName),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null ?
                                    new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                            return img;
                        }
                    };
                    XElement html = OpenXmlPowerTools.DocxToHtml.ConvertToHtml(wDoc, settings);

                    // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
                    // XEntity.  PtOpenXmlUtil.cs define the XEntity class.  See
                    // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
                    // for detailed explanation.
                    //
                    // If you further transform the XML tree returned by ConvertToHtmlTransform, you
                    // must do it correctly, or entities will not be serialized properly.

                    var htmlString = html.ToString(SaveOptions.DisableFormatting);
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                }
            }
        }

#if DO_CONVERSION_VIA_WORD
        public static void ConvertToHtmlUsingWord(FileInfo sourceFileName, FileInfo destFileName)
        {
            Word.Application app = new Word.Application();
            app.Visible = false;
            try
            {
                Word.Document doc = app.Documents.Open(sourceFileName.FullName);
                doc.SaveAs2(destFileName.FullName, Word.WdSaveFormat.wdFormatFilteredHTML);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Console.WriteLine("Caught unexpected COM exception.");
                ((Microsoft.Office.Interop.Word._Application)app).Quit();
                Environment.Exit(0);
            }
            ((Microsoft.Office.Interop.Word._Application)app).Quit();
        }
#endif
    }
}

#endif
