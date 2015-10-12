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
using System.Xml;
using System.Xml.Linq;
using HtmlAgilityPack;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenXmlPowerTools;
using Xunit;

namespace OxPt
{
    public class LiTests
    {

        // PowerShell oneliner that generates InlineData for all files in a directory
        // dir | % { '[InlineData("' + $_.Name + '")]' } | clip

        [Theory]
        [InlineData("LIR001-en-US-ordinal.docx")]
        [InlineData("LIR002-en-US-ordinalText.docx")]
        [InlineData("LIR003-en-US-upperLetter.docx")]
        [InlineData("LIR004-en-US-upperRoman.docx")]
        [InlineData("LIR005-fr-FR-cardinalText.docx")]
        [InlineData("LIR006-fr-FR-ordinalText.docx")]
        // [InlineData("LIR007-ru-RU-ordinalText.docx")]  // todo this fails, the code needs updated.
        [InlineData("LIR008-zh-CH-chineseCountingThousand.docx")]
        [InlineData("LIR009-zh-CN-chineseCounting.docx")]
        [InlineData("LIR010-zh-CN-ideographTraditional.docx")]
        [InlineData("LIR011-en-US-00001.docx")]
        [InlineData("LIR012-en-US-0001.docx")]
        [InlineData("LIR013-en-US-001.docx")]
        [InlineData("LIR014-en-US-01.docx")]
        [InlineData("LIR015-en-US-cardinalText.docx")]
        [InlineData("LIR016-en-US-decimal.docx")]
        [InlineData("LIR017-en-US-decimalEnclosedCircle.docx")]
        [InlineData("LIR018-en-US-decimalZero.docx")]
        [InlineData("LIR019-en-US-lowerLetter.docx")]
        [InlineData("LIR020-en-US-lowerRoman.docx")]
        public void LIR001(string file)
        {
            FileInfo lirFile = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, file));
            WmlDocument wmlDoc = new WmlDocument(lirFile.FullName);

            var wordHtmlFile = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, lirFile.Name.Replace(".docx", "-Word.html")));
            WordAutomationUtilities.SaveAsHtmlUsingWord(lirFile, wordHtmlFile);

            var ptHtmlFile = ConvertToHtml(lirFile.FullName, TestUtil.TempDir.FullName);
            var fiPtXml = SaveHtmlAsXml(ptHtmlFile);

            // read and write to get the BOM on the file
            var wh = File.ReadAllText(wordHtmlFile.FullName, Encoding.Default);
            File.WriteAllText(wordHtmlFile.FullName, wh, Encoding.UTF8);

            var wordXml = SaveHtmlAsXml(wordHtmlFile);
            CompareNumbering(fiPtXml, wordXml);
        }

        private static void CompareNumbering(FileInfo fiPtXml, FileInfo wordXml)
        {
            char splitChar = '|';

            Console.WriteLine("Comparing {0} to {1}", fiPtXml.Name, wordXml.Name);
            var xdWord = XDocument.Load(wordXml.FullName);
            List<string> wordRawParagraphText = xdWord.Descendants("h2").Select(p => p.DescendantNodes().OfType<XText>().Select(t => t.Value).Aggregate((s, i) => s + i)).ToList();
            var wordTextToCompare = wordRawParagraphText.Where(p => p.Contains(splitChar.ToString())).Select(p => p.Split(splitChar)[0]).ToList();

            var xdPt = XDocument.Load(fiPtXml.FullName);
            XNamespace xhtml = "http://www.w3.org/1999/xhtml";
            List<string> ptRawParagraphText = xdPt.Descendants(xhtml + "h2").Select(p => p.DescendantNodes().OfType<XText>().Select(t => t.Value).Aggregate((s, i) => s + i)).ToList();
            var ptTextToCompare = ptRawParagraphText.Where(p => p.Contains(splitChar.ToString())).Select(p => new
            {
                ListItem = p.Split(splitChar)[0],
                ParaText = p.Split(splitChar)[1],
            }).ToList();

            if (!wordTextToCompare.Any())
            {
                throw new Exception("Internal error - no items selected");
            }
            if (wordTextToCompare.Count() != ptTextToCompare.Count())
            {
                Assert.True(false);
                //Console.WriteLine("Error, differing line counts");
                //Console.WriteLine("Word line count: {0}", wordTextToCompare.Count());
                //Console.WriteLine("Pt line count: {0}", ptTextToCompare.Count());
                return;
            }
            var zipped = wordTextToCompare.Zip(ptTextToCompare, (w, p) => new
            {
                WordText = w,
                PtText = p.ListItem,
                ParagraphText = p.ParaText,
            });
            var mismatchList = zipped.Where(z =>
            {
                var w = z.WordText.Replace('\n', ' ').Trim();
                var p = z.PtText.Replace('\n', ' ').Trim();
                var match = w == p;
                if (match)
                    return false;
                return true;
            }).Select(z => z.ParagraphText).ToList();
            if (mismatchList.Any())
            {
                Assert.True(false);
                //Console.WriteLine("Mismatches: {0}", mismatchList.Count());
                //foreach (var item in mismatchList.Take(20))
                //{
                //    Console.WriteLine(item);
                //}
            }
        }

        public static FileInfo ConvertToHtml(string file, string outputDirectory)
        {
            var fi = new FileInfo(file);
            byte[] byteArray = File.ReadAllBytes(fi.FullName);
            FileInfo destFileName;
            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(memoryStream, true))
                {
                    destFileName = new FileInfo(fi.Name.Replace(".docx", ".html"));
                    if (outputDirectory != null && outputDirectory != string.Empty)
                    {
                        DirectoryInfo di = new DirectoryInfo(outputDirectory);
                        if (!di.Exists)
                        {
                            throw new OpenXmlPowerToolsException("Output directory does not exist");
                        }
                        destFileName = new FileInfo(Path.Combine(di.FullName, destFileName.Name));
                    }
                    var imageDirectoryName = destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";
                    int imageCounter = 0;
                    var pageTitle = (string)wDoc.CoreFilePropertiesPart.GetXDocument().Descendants(DC.title).FirstOrDefault();
                    if (pageTitle == null)
                        pageTitle = fi.FullName;

                    HtmlConverterSettings settings = new HtmlConverterSettings()
                    {
                        PageTitle = pageTitle,
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
                    XElement html = HtmlConverter.ConvertToHtml(wDoc, settings);

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
            return destFileName;
        }

        public static FileInfo SaveHtmlAsXml(FileInfo htmlFileName)
        {
            string baseName = htmlFileName.Name.Substring(0, htmlFileName.Name.Length - htmlFileName.Extension.Length);
            FileInfo destFile = new FileInfo(Path.Combine(htmlFileName.DirectoryName, baseName + ".xml"));

            HtmlDocument hdoc = new HtmlDocument();
            hdoc.Load(htmlFileName.FullName, Encoding.UTF8);
            hdoc.OptionOutputAsXml = true;
            hdoc.Save(destFile.FullName, Encoding.UTF8);
            StringBuilder sb = new StringBuilder(File.ReadAllText(destFile.FullName, Encoding.Default));
            sb.Replace("è", "&egrave;");
            sb.Replace("&amp;", "&");
            sb.Replace("&nbsp;", "\xA0");
            sb.Replace("&quot;", "\"");
            sb.Replace("&lt;", "~lt;");
            sb.Replace("&gt;", "~gt;");
            sb.Replace("&#", "~#");
            sb.Replace("&", "&amp;");
            sb.Replace("~lt;", "&lt;");
            sb.Replace("~gt;", "&gt;");
            sb.Replace("~#", "&#");
            File.WriteAllText(destFile.FullName, sb.ToString(), Encoding.UTF8);
            return destFile;
        }
    }

}
