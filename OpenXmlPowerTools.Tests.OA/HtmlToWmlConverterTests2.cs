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
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenXmlPowerTools;
using Xunit;
using HtmlAgilityPack;
using System.Text.RegularExpressions;

/*******************************************************************************************
 * HtmlToWmlConverter expects the HTML to be passed as an XElement, i.e. as XML.  While the HTML test files that
 * are included in Open-Xml-PowerTools are able to be read as XML, most HTML is not able to be read as XML.
 * The best solution is to use the HtmlAgilityPack, which can parse HTML and save as XML.  The HtmlAgilityPack
 * is licensed under the Ms-PL (same as Open-Xml-PowerTools) so it is convenient to include it in your solution,
 * and thereby you can convert HTML to XML that can be processed by the HtmlToWmlConverter.
 * 
 * A convenient way to get the DLL that has been checked out with HtmlToWmlConverter is to clone the repo at
 * https://github.com/EricWhiteDev/HtmlAgilityPack
 * 
 * That repo contains only the DLL that has been checked out with HtmlToWmlConverter.
 * 
 * Of course, you can also get the HtmlAgilityPack source and compile it to get the DLL.  You can find it at
 * http://codeplex.com/HtmlAgilityPack
 * 
 * We don't include the HtmlAgilityPack in Open-Xml-PowerTools, to simplify installation.  The XUnit tests in
 * this module do not require the HtmlAgilityPack to run.
*******************************************************************************************/

#if DO_CONVERSION_VIA_WORD
using Word = Microsoft.Office.Interop.Word;
#endif

/***************************************************************************************************
 * The XUnit tests in this module are not included in the standard Open-Xml-PowerTools tests because
 * they use either Word automation, the HtmlAgilityPack, or both.
***************************************************************************************************/ 

namespace OxPt
{
    public class HwTests2
    {
        static bool s_ProduceAnnotatedHtml = true;

        // PowerShell oneliner that generates InlineData for all files in a directory
        // dir | % { '[InlineData("' + $_.Name + '")]' } | clip

        [Theory]
        [InlineData("HW002-Table01.docx")]
        [InlineData("HW002-Table02.docx")]
        [InlineData("HW002-Table03.docx")]
        [InlineData("HW002-Table04.docx")]
        [InlineData("HW002-Table05.docx")]
        [InlineData("HW002-Table06.docx")]
        [InlineData("HW002-Table07.docx")]
        [InlineData("HW002-Table08.docx")]
        [InlineData("HW002-Table09.docx")]
        [InlineData("HW002-Table10.docx")]
        [InlineData("HW002-Table11.docx")]
        [InlineData("HW002-Table12.docx")]
        [InlineData("HW002-Table13.docx")]
        [InlineData("HW002-Table14.docx")]
        [InlineData("HW002-Table15.docx")]
        [InlineData("HW002-Table16.docx")]
        [InlineData("HW002-Table17.docx")]
        [InlineData("HW002-Table18.docx")]
        public void HW002(string name)
        {
            var sourceDocxFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            var sourceCopiedToDestDocxFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-2-Source.docx")));
            var sourceCopiedToDestHtmlFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-2-Source.html")));
            var destCssFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-3.css")));
            var destDocxFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-4-ConvertedByHtmlToWml.docx")));
            var annotatedHtmlFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-5-Annotated.txt")));

            File.Copy(sourceDocxFi.FullName, sourceCopiedToDestDocxFi.FullName);

            WordAutomationUtilities.SaveAsHtmlUsingWord(sourceDocxFi, sourceCopiedToDestHtmlFi);
            XElement html = null;
            int cnt = 0;
            while (true)
            {
                try
                {
                    html = HtmlToWmlReadAsXElement.ReadAsXElement(sourceCopiedToDestHtmlFi);
                    break;
                }
                catch (XmlException e)
                {
                    throw e;
                }
                catch (IOException i)
                {
                    if (++cnt == 20)
                        throw i;
                    System.Threading.Thread.Sleep(50);
                    continue;
                }
            }

            string usedAuthorCss = HtmlToWmlConverter.CleanUpCss((string)html.Descendants().FirstOrDefault(d => d.Name.LocalName.ToLower() == "style"));
            File.WriteAllText(destCssFi.FullName, usedAuthorCss);

            HtmlToWmlConverterSettings settings = HtmlToWmlConverter.GetDefaultSettings();
            // image references in HTML files contain the path to the subdir that contains the images, so base URI is the name of the directory
            // that contains the HTML files
            settings.BaseUriForImages = Path.Combine(TestUtil.TempDir.FullName);

            WmlDocument doc = HtmlToWmlConverter.ConvertHtmlToWml(
                defaultCss,
                usedAuthorCss,
                userCss,
                html,
                settings,
                null,  // use the default EmptyDocument
                s_ProduceAnnotatedHtml ? annotatedHtmlFi.FullName : null);

            Assert.NotNull(doc);

            if (doc != null)
                SaveValidateAndFormatMainDocPart(destDocxFi, doc);

#if DO_CONVERSION_VIA_WORD
            var newAltChunkBeforeFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-5-AltChunkBefore.docx")));
            var newAltChunkAfterFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-6-ConvertedViaWord.docx")));
            WordAutomationUtilities.DoConversionViaWord(newAltChunkBeforeFi, newAltChunkAfterFi, html);
#endif
        }

        [Theory]
        [InlineData("T0015.html")]
        public void HW003(string name)
        {
            string testDocPrefix = "HW003_";
            var sourceHtmlFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            var sourceCopiedToDestHtmlFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, (testDocPrefix + sourceHtmlFi.Name).Replace(".html", "-1-Source.html")));
            var destCssFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, (testDocPrefix + sourceHtmlFi.Name).Replace(".html", "-2.css")));
            var destDocxFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, (testDocPrefix + sourceHtmlFi.Name).Replace(".html", "-3-ConvertedByHtmlToWml.docx")));
            var annotatedHtmlFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, (testDocPrefix + sourceHtmlFi.Name).Replace(".html", "-4-Annotated.txt")));

            File.Copy(sourceHtmlFi.FullName, sourceCopiedToDestHtmlFi.FullName);
            XElement html = HtmlToWmlReadAsXElement.ReadAsXElement(sourceCopiedToDestHtmlFi);

            string usedAuthorCss = HtmlToWmlConverter.CleanUpCss((string)html.Descendants().FirstOrDefault(d => d.Name.LocalName.ToLower() == "style"));
            File.WriteAllText(destCssFi.FullName, usedAuthorCss);

            HtmlToWmlConverterSettings settings = HtmlToWmlConverter.GetDefaultSettings();
            settings.BaseUriForImages = Path.Combine(TestUtil.TempDir.FullName);
            settings.DefaultBlockContentMargin = "36pt";

            WmlDocument doc = HtmlToWmlConverter.ConvertHtmlToWml(defaultCss, usedAuthorCss, userCss, html, settings, null, s_ProduceAnnotatedHtml ? annotatedHtmlFi.FullName : null);
            Assert.NotNull(doc);
            if (doc != null)
                SaveValidateAndFormatMainDocPart(destDocxFi, doc);
        }

        [Theory]
        [InlineData("HW010-Symbols01.docx")]
        [InlineData("HW010-Symbols02.docx")]
        [InlineData("HW010-TableWithEmptyRows.docx")]
        [InlineData("HW010-TableWithThreeEmptyRows.docx")]
        [InlineData("HW010-TableWithImage.docx")]
        [InlineData("HW010-SpanWithSingleSpace.docx")]
        [InlineData("HW010-Tab01.docx")]
        
        public void HW010(string name)
        {
            var sourceDocxFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            var sourceCopiedToDestDocxFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-2-Source.docx")));
            var sourceCopiedToDestHtmlFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-2-Source.html")));
            var destCssFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-3.css")));
            var destDocxFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-4-ConvertedByHtmlToWml.docx")));
            var annotatedHtmlFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-5-Annotated.txt")));

            File.Copy(sourceDocxFi.FullName, sourceCopiedToDestDocxFi.FullName);

            SaveAsHtmlUsingHtmlConverter(sourceCopiedToDestDocxFi.FullName, sourceCopiedToDestDocxFi.DirectoryName);
            XElement html = HtmlToWmlReadAsXElement.ReadAsXElement(sourceCopiedToDestHtmlFi);

            string usedAuthorCss = HtmlToWmlConverter.CleanUpCss((string)html.Descendants().FirstOrDefault(d => d.Name.LocalName.ToLower() == "style"));
            File.WriteAllText(destCssFi.FullName, usedAuthorCss);

            var settingsWmlDocument = new WmlDocument(sourceCopiedToDestDocxFi.FullName);
            HtmlToWmlConverterSettings settings = HtmlToWmlConverter.GetDefaultSettings(settingsWmlDocument);
            // image references in HTML files contain the path to the subdir that contains the images, so base URI is the name of the directory
            // that contains the HTML files
            settings.BaseUriForImages = Path.Combine(TestUtil.TempDir.FullName);

            WmlDocument doc = HtmlToWmlConverter.ConvertHtmlToWml(
                defaultCss,
                usedAuthorCss,
                userCss,
                html,
                settings,
                null,  // use the default EmptyDocument
                s_ProduceAnnotatedHtml ? annotatedHtmlFi.FullName : null);

            Assert.NotNull(doc);

            if (doc != null)
                SaveValidateAndFormatMainDocPart(destDocxFi, doc);

#if DO_CONVERSION_VIA_WORD
            var newAltChunkBeforeFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-5-AltChunkBefore.docx")));
            var newAltChunkAfterFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-6-ConvertedViaWord.docx")));
            WordAutomationUtilities.DoConversionViaWord(newAltChunkBeforeFi, newAltChunkAfterFi, html);
#endif
        }

        private static void SaveAsHtmlUsingHtmlConverter(string file, string outputDirectory)
        {
            var fi = new FileInfo(file);
            Console.WriteLine(fi.Name);
            byte[] byteArray = File.ReadAllBytes(fi.FullName);
            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(memoryStream, true))
                {
                    var destFileName = new FileInfo(fi.Name.Replace(".docx", ".html"));
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

                    var pageTitle = fi.FullName;
                    var part = wDoc.CoreFilePropertiesPart;
                    if (part != null)
                    {
                        pageTitle = (string)part.GetXDocument().Descendants(DC.title).FirstOrDefault() ?? fi.FullName;
                    }

                    // TODO: Determine max-width from size of content area.
                    HtmlConverterSettings settings = new HtmlConverterSettings()
                    {
                        AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
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
                                imageFormat = ImageFormat.Png;
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
                            string imageSource = localDirInfo.Name + "/image" +
                                imageCounter.ToString() + "." + extension;

                            XElement img = new XElement(Xhtml.img,
                                new XAttribute(NoNamespace.src, imageSource),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null ?
                                    new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                            return img;
                        }
                    };
                    XElement htmlElement = HtmlConverter.ConvertToHtml(wDoc, settings);

                    // Produce HTML document with <!DOCTYPE html > declaration to tell the browser
                    // we are using HTML5.
                    var html = new XDocument(
                        new XDocumentType("html", null, null, null),
                        htmlElement);

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

        private static void SaveValidateAndFormatMainDocPart(FileInfo destDocxFi, WmlDocument doc)
        {
            WmlDocument formattedDoc;

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
                using (WordprocessingDocument document = WordprocessingDocument.Open(ms, true))
                {
                    XDocument xDoc = document.MainDocumentPart.GetXDocument();
                    document.MainDocumentPart.PutXDocumentWithFormatting();
                    OpenXmlValidator validator = new OpenXmlValidator();
                    var errors = validator.Validate(document);
                    var errorsString = errors
                        .Select(e => e.Description + Environment.NewLine)
                        .StringConcatenate();

                    // Assert that there were no errors in the generated document.
                    Assert.Equal("", errorsString);
                }
                formattedDoc = new WmlDocument(destDocxFi.FullName, ms.ToArray());
            }
            formattedDoc.SaveAs(destDocxFi.FullName);
        }


        /*
            * display property:
            * - inline
            * - block
            * - list-item
            * - inline-block
            * - table
            * - inline-table
            * - table-row-group
            * - table-header-group
            * - table-footer-group
            * - table-row
            * - table-column-group
            * - table-column
            * - table-cell
            * - table-caption
            * - none
            * - inherit
            * 
            * position property:
            * - static
            * - relative
            * - absolute
            * - fixed
            * - inherit
            * 
            * top, left, bottom, right properties:
            * (only apply if position property is not static)
        */

        static string defaultCss =
            @"html, address,
blockquote,
body, dd, div,
dl, dt, fieldset, form,
frame, frameset,
h1, h2, h3, h4,
h5, h6, noframes,
ol, p, ul, center,
dir, hr, menu, pre { display: block; unicode-bidi: embed }
li { display: list-item }
head { display: none }
table { display: table }
tr { display: table-row }
thead { display: table-header-group }
tbody { display: table-row-group }
tfoot { display: table-footer-group }
col { display: table-column }
colgroup { display: table-column-group }
td, th { display: table-cell }
caption { display: table-caption }
th { font-weight: bolder; text-align: center }
caption { text-align: center }
body { margin: auto; }
h1 { font-size: 2em; margin: auto; }
h2 { font-size: 1.5em; margin: auto; }
h3 { font-size: 1.17em; margin: auto; }
h4, p,
blockquote, ul,
fieldset, form,
ol, dl, dir,
menu { margin: auto }
a { color: blue; }
h5 { font-size: .83em; margin: auto }
h6 { font-size: .75em; margin: auto }
h1, h2, h3, h4,
h5, h6, b,
strong { font-weight: bolder }
blockquote { margin-left: 40px; margin-right: 40px }
i, cite, em,
var, address { font-style: italic }
pre, tt, code,
kbd, samp { font-family: monospace }
pre { white-space: pre }
button, textarea,
input, select { display: inline-block }
big { font-size: 1.17em }
small, sub, sup { font-size: .83em }
sub { vertical-align: sub }
sup { vertical-align: super }
table { border-spacing: 2px; }
thead, tbody,
tfoot { vertical-align: middle }
td, th, tr { vertical-align: inherit }
s, strike, del { text-decoration: line-through }
hr { border: 1px inset }
ol, ul, dir,
menu, dd { margin-left: 40px }
ol { list-style-type: decimal }
ol ul, ul ol,
ul ul, ol ol { margin-top: 0; margin-bottom: 0 }
u, ins { text-decoration: underline }
br:before { content: ""\A""; white-space: pre-line }
center { text-align: center }
:link, :visited { text-decoration: underline }
:focus { outline: thin dotted invert }
/* Begin bidirectionality settings (do not change) */
BDO[DIR=""ltr""] { direction: ltr; unicode-bidi: bidi-override }
BDO[DIR=""rtl""] { direction: rtl; unicode-bidi: bidi-override }
*[DIR=""ltr""] { direction: ltr; unicode-bidi: embed }
*[DIR=""rtl""] { direction: rtl; unicode-bidi: embed }

";

        // original defaultCss
        //    static string defaultCss =
        //        @"html, address,
        //blockquote,
        //body, dd, div,
        //dl, dt, fieldset, form,
        //frame, frameset,
        //h1, h2, h3, h4,
        //h5, h6, noframes,
        //ol, p, ul, center,
        //dir, hr, menu, pre { display: block; unicode-bidi: embed }
        //li { display: list-item }
        //head { display: none }
        //table { display: table }
        //tr { display: table-row }
        //thead { display: table-header-group }
        //tbody { display: table-row-group }
        //tfoot { display: table-footer-group }
        //col { display: table-column }
        //colgroup { display: table-column-group }
        //td, th { display: table-cell }
        //caption { display: table-caption }
        //th { font-weight: bolder; text-align: center }
        //caption { text-align: center }
        //body { margin: 8px; }
        //h1 { font-size: 2em; margin: .67em 0em; }
        //h2 { font-size: 1.5em; margin: .75em 0em; }
        //h3 { font-size: 1.17em; margin: .83em 0em; }
        //h4, p,
        //blockquote, ul,
        //fieldset, form,
        //ol, dl, dir,
        //menu { margin: 1.12em 0 }
        //a { color: blue; }
        //h5 { font-size: .83em; margin: 1.5em 0 }
        //h6 { font-size: .75em; margin: 1.67em 0 }
        //h1, h2, h3, h4,
        //h5, h6, b,
        //strong { font-weight: bolder }
        //blockquote { margin-left: 40px; margin-right: 40px }
        //i, cite, em,
        //var, address { font-style: italic }
        //pre, tt, code,
        //kbd, samp { font-family: monospace }
        //pre { white-space: pre }
        //button, textarea,
        //input, select { display: inline-block }
        //big { font-size: 1.17em }
        //small, sub, sup { font-size: .83em }
        //sub { vertical-align: sub }
        //sup { vertical-align: super }
        //table { border-spacing: 2px; }
        //thead, tbody,
        //tfoot { vertical-align: middle }
        //td, th, tr { vertical-align: inherit }
        //s, strike, del { text-decoration: line-through }
        //hr { border: 1px inset }
        //ol, ul, dir,
        //menu, dd { margin-left: 40px }
        //ol { list-style-type: decimal }
        //ol ul, ul ol,
        //ul ul, ol ol { margin-top: 0; margin-bottom: 0 }
        //u, ins { text-decoration: underline }
        //br:before { content: ""\A""; white-space: pre-line }
        //center { text-align: center }
        //:link, :visited { text-decoration: underline }
        //:focus { outline: thin dotted invert }
        ///* Begin bidirectionality settings (do not change) */
        //BDO[DIR=""ltr""] { direction: ltr; unicode-bidi: bidi-override }
        //BDO[DIR=""rtl""] { direction: rtl; unicode-bidi: bidi-override }
        //*[DIR=""ltr""] { direction: ltr; unicode-bidi: embed }
        //*[DIR=""rtl""] { direction: rtl; unicode-bidi: embed }";

        static string userCss = @"";
    }
}
