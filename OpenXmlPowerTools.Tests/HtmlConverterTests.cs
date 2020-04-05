// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#define COPY_FILES_FOR_DEBUGGING

using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class HcTests
    {
        // PowerShell oneliner that generates InlineData for all files in a directory
        // dir | % { '[InlineData("' + $_.Name + '")]' } | clip

        [Theory]
        [InlineData("HC001-5DayTourPlanTemplate.docx")]
        [InlineData("HC002-Hebrew-01.docx")]
        [InlineData("HC003-Hebrew-02.docx")]
        [InlineData("HC004-ResumeTemplate.docx")]
        [InlineData("HC005-TaskPlanTemplate.docx")]
        [InlineData("HC006-Test-01.docx")]
        [InlineData("HC007-Test-02.docx")]
        [InlineData("HC008-Test-03.docx")]
        [InlineData("HC009-Test-04.docx")]
        [InlineData("HC010-Test-05.docx")]
        [InlineData("HC011-Test-06.docx")]
        [InlineData("HC012-Test-07.docx")]
        [InlineData("HC013-Test-08.docx")]
        [InlineData("HC014-RTL-Table-01.docx")]
        [InlineData("HC015-Vertical-Spacing-atLeast.docx")]
        [InlineData("HC016-Horizontal-Spacing-firstLine.docx")]
        [InlineData("HC017-Vertical-Alignment-Cell-01.docx")]
        [InlineData("HC018-Vertical-Alignment-Para-01.docx")]
        [InlineData("HC019-Hidden-Run.docx")]
        [InlineData("HC020-Small-Caps.docx")]
        [InlineData("HC021-Symbols.docx")]
        [InlineData("HC022-Table-Of-Contents.docx")]
        [InlineData("HC023-Hyperlink.docx")]
        [InlineData("HC024-Tabs-01.docx")]
        [InlineData("HC025-Tabs-02.docx")]
        [InlineData("HC026-Tabs-03.docx")]
        [InlineData("HC027-Tabs-04.docx")]
        [InlineData("HC028-No-Break-Hyphen.docx")]
        [InlineData("HC029-Table-Merged-Cells.docx")]
        [InlineData("HC030-Content-Controls.docx")]
        [InlineData("HC031-Complicated-Document.docx")]
        [InlineData("HC032-Named-Color.docx")]
        [InlineData("HC033-Run-With-Border.docx")]
        [InlineData("HC034-Run-With-Position.docx")]
        [InlineData("HC035-Strike-Through.docx")]
        [InlineData("HC036-Super-Script.docx")]
        [InlineData("HC037-Sub-Script.docx")]
        [InlineData("HC038-Conflicting-Border-Weight.docx")]
        [InlineData("HC039-Bold.docx")]
        [InlineData("HC040-Hyperlink-Fieldcode-01.docx")]
        [InlineData("HC041-Hyperlink-Fieldcode-02.docx")]
        [InlineData("HC042-Image-Png.docx")]
        [InlineData("HC043-Chart.docx")]
        [InlineData("HC044-Embedded-Workbook.docx")]
        [InlineData("HC045-Italic.docx")]
        [InlineData("HC046-BoldAndItalic.docx")]
        [InlineData("HC047-No-Section.docx")]
        [InlineData("HC048-Excerpt.docx")]
        [InlineData("HC049-Borders.docx")]
        [InlineData("HC050-Shaded-Text-01.docx")]
        [InlineData("HC051-Shaded-Text-02.docx")]
        [InlineData("HC060-Image-with-Hyperlink.docx")]
        [InlineData("HC061-Hyperlink-in-Field.docx")]
        public void HC001(string name)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));

#if COPY_FILES_FOR_DEBUGGING
            var sourceCopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-1-Source.docx")));
            if (!sourceCopiedToDestDocx.Exists)
            {
                File.Copy(sourceDocx.FullName, sourceCopiedToDestDocx.FullName);
            }

            var assembledFormattingDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-2-FormattingAssembled.docx")));
            if (!assembledFormattingDestDocx.Exists)
            {
                CopyFormattingAssembledDocx(sourceDocx, assembledFormattingDestDocx);
            }
#endif

            var oxPtConvertedDestHtml = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-3-OxPt.html")));
            ConvertToHtml(sourceDocx, oxPtConvertedDestHtml);
        }

        [Theory]
        [InlineData("HC006-Test-01.docx")]
        public void HC002_NoCssClasses(string name)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));

            var oxPtConvertedDestHtml = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-5-OxPt-No-CSS-Classes.html")));
            ConvertToHtmlNoCssClasses(sourceDocx, oxPtConvertedDestHtml);
        }

        private static void CopyFormattingAssembledDocx(FileInfo source, FileInfo dest)
        {
            var ba = File.ReadAllBytes(source.FullName);
            using (var ms = new MemoryStream())
            {
                ms.Write(ba, 0, ba.Length);
                using (var wordDoc = WordprocessingDocument.Open(ms, true))
                {
                    RevisionAccepter.AcceptRevisions(wordDoc);
                    var simplifyMarkupSettings = new SimplifyMarkupSettings
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

                    var formattingAssemblerSettings = new FormattingAssemblerSettings
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

        private static void ConvertToHtml(FileInfo sourceDocx, FileInfo destFileName)
        {
            var byteArray = File.ReadAllBytes(sourceDocx.FullName);
            using (var memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (var wDoc = WordprocessingDocument.Open(memoryStream, true))
                {
                    var outputDirectory = destFileName.Directory;
                    destFileName = new FileInfo(Path.Combine(outputDirectory.FullName, destFileName.Name));
                    var imageDirectoryName = destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";
                    var imageCounter = 0;
                    var pageTitle = (string)wDoc.CoreFilePropertiesPart.GetXDocument().Descendants(DC.title).FirstOrDefault();
                    if (pageTitle == null)
                    {
                        pageTitle = sourceDocx.FullName;
                    }

                    var settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = pageTitle,
                        FabricateCssClasses = true,
                        CssClassPrefix = "pt-",
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        ImageHandler = imageInfo =>
                        {
                            var localDirInfo = new DirectoryInfo(imageDirectoryName);
                            if (!localDirInfo.Exists)
                            {
                                localDirInfo.Create();
                            }

                            ++imageCounter;
                            var extension = imageInfo.ContentType.Split('/')[1].ToUpperInvariant();
                            ImageFormat imageFormat = null;
                            if (extension == "PNG")
                            {
                                // Convert png to jpeg.
                                extension = "GIF";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "GIF")
                            {
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "BMP")
                            {
                                imageFormat = ImageFormat.Bmp;
                            }
                            else if (extension == "JPEG")
                            {
                                imageFormat = ImageFormat.Jpeg;
                            }
                            else if (extension == "TIFF")
                            {
                                // Convert tiff to gif.
                                extension = "GIF";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "X-WMF")
                            {
                                extension = "WMF";
                                imageFormat = ImageFormat.Wmf;
                            }

                            // If the image format isn't one that we expect, ignore it, and don't return markup for the link.
                            if (imageFormat == null)
                            {
                                return null;
                            }

                            var imageFileName = imageDirectoryName + "/image" + imageCounter.ToString(CultureInfo.InvariantCulture) + "." + extension;
                            try
                            {
                                imageInfo.Bitmap.Save(imageFileName, imageFormat);
                            }
                            catch (System.Runtime.InteropServices.ExternalException)
                            {
                                return null;
                            }
                            var img = new XElement(Xhtml.img,
                                new XAttribute(NoNamespace.src, imageFileName),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null ?
                                    new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                            return img;
                        }
                    };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

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
            var byteArray = File.ReadAllBytes(sourceDocx.FullName);
            using (var memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (var wDoc = WordprocessingDocument.Open(memoryStream, true))
                {
                    var outputDirectory = destFileName.Directory;
                    destFileName = new FileInfo(Path.Combine(outputDirectory.FullName, destFileName.Name));
                    var imageDirectoryName = destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";
                    var imageCounter = 0;
                    var pageTitle = (string)wDoc.CoreFilePropertiesPart.GetXDocument().Descendants(DC.title).FirstOrDefault();
                    if (pageTitle == null)
                    {
                        pageTitle = sourceDocx.FullName;
                    }

                    var settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = pageTitle,
                        FabricateCssClasses = false,
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        ImageHandler = imageInfo =>
                        {
                            var localDirInfo = new DirectoryInfo(imageDirectoryName);
                            if (!localDirInfo.Exists)
                            {
                                localDirInfo.Create();
                            }

                            ++imageCounter;
                            var extension = imageInfo.ContentType.Split('/')[1].ToUpperInvariant();
                            ImageFormat imageFormat = null;
                            if (extension == "PNG")
                            {
                                // Convert png to jpeg.
                                extension = "GIF";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "GIF")
                            {
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "BMP")
                            {
                                imageFormat = ImageFormat.Bmp;
                            }
                            else if (extension == "JPEG")
                            {
                                imageFormat = ImageFormat.Jpeg;
                            }
                            else if (extension == "TIFF")
                            {
                                // Convert tiff to gif.
                                extension = "GIF";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "X-WMF")
                            {
                                extension = "WMF";
                                imageFormat = ImageFormat.Wmf;
                            }

                            // If the image format isn't one that we expect, ignore it,
                            // and don't return markup for the link.
                            if (imageFormat == null)
                            {
                                return null;
                            }

                            var imageFileName = imageDirectoryName + "/image" + imageCounter.ToString(CultureInfo.InvariantCulture) + "." + extension;
                            try
                            {
                                imageInfo.Bitmap.Save(imageFileName, imageFormat);
                            }
                            catch (System.Runtime.InteropServices.ExternalException)
                            {
                                return null;
                            }
                            var img = new XElement(Xhtml.img,
                                new XAttribute(NoNamespace.src, imageFileName),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null ?
                                    new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                            return img;
                        }
                    };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

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
    }
}

#endif