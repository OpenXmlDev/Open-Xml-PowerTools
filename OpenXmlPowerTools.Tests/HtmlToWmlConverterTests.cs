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

namespace OxPt
{
    public class HwTests
    {
        static bool s_ProduceAnnotatedHtml = true;

        // PowerShell oneliner that generates InlineData for all files in a directory
        // dir | % { '[InlineData("' + $_.Name + '")]' } | clip

        [Theory]
        [InlineData("T0010.html")]
        [InlineData("T0011.html")]
        [InlineData("T0012.html")]
        [InlineData("T0013.html")]
        [InlineData("T0014.html")]
        [InlineData("T0015.html")]
        [InlineData("T0020.html")]
        [InlineData("T0030.html")]
        [InlineData("T0040.html")]
        [InlineData("T0050.html")]
        [InlineData("T0060.html")]
        [InlineData("T0070.html")]
        [InlineData("T0080.html")]
        [InlineData("T0090.html")]
        [InlineData("T0100.html")]
        [InlineData("T0110.html")]
        [InlineData("T0111.html")]
        [InlineData("T0112.html")]
        [InlineData("T0120.html")]
        [InlineData("T0130.html")]
        [InlineData("T0140.html")]
        [InlineData("T0150.html")]
        [InlineData("T0160.html")]
        [InlineData("T0170.html")]
        [InlineData("T0180.html")]
        [InlineData("T0190.html")]
        [InlineData("T0200.html")]
        [InlineData("T0210.html")]
        [InlineData("T0220.html")]
        [InlineData("T0230.html")]
        [InlineData("T0240.html")]
        [InlineData("T0250.html")]
        [InlineData("T0251.html")]
        [InlineData("T0260.html")]
        [InlineData("T0270.html")]
        [InlineData("T0280.html")]
        [InlineData("T0290.html")]
        [InlineData("T0300.html")]
        [InlineData("T0310.html")]
        [InlineData("T0320.html")]
        [InlineData("T0330.html")]
        [InlineData("T0340.html")]
        [InlineData("T0350.html")]
        [InlineData("T0360.html")]
        [InlineData("T0370.html")]
        [InlineData("T0380.html")]
        [InlineData("T0390.html")]
        [InlineData("T0400.html")]
        [InlineData("T0410.html")]
        [InlineData("T0420.html")]
        [InlineData("T0430.html")]
        [InlineData("T0440.html")]
        [InlineData("T0450.html")]
        [InlineData("T0460.html")]
        [InlineData("T0470.html")]
        [InlineData("T0480.html")]
        [InlineData("T0490.html")]
        [InlineData("T0500.html")]
        [InlineData("T0510.html")]
        [InlineData("T0520.html")]
        [InlineData("T0530.html")]
        [InlineData("T0540.html")]
        [InlineData("T0550.html")]
        [InlineData("T0560.html")]
        [InlineData("T0570.html")]
        [InlineData("T0580.html")]
        [InlineData("T0590.html")]
        [InlineData("T0600.html")]
        [InlineData("T0610.html")]
        [InlineData("T0620.html")]
        [InlineData("T0622.html")]
        [InlineData("T0630.html")]
        [InlineData("T0640.html")]
        [InlineData("T0650.html")]
        [InlineData("T0651.html")]
        [InlineData("T0660.html")]
        [InlineData("T0670.html")]
        [InlineData("T0680.html")]
        [InlineData("T0690.html")]
        [InlineData("T0691.html")]
        [InlineData("T0692.html")]
        [InlineData("T0700.html")]
        [InlineData("T0710.html")]
        [InlineData("T0720.html")]
        [InlineData("T0730.html")]
        [InlineData("T0740.html")]
        [InlineData("T0742.html")]
        [InlineData("T0745.html")]
        [InlineData("T0750.html")]
        [InlineData("T0760.html")]
        [InlineData("T0770.html")]
        [InlineData("T0780.html")]
        [InlineData("T0790.html")]
        [InlineData("T0791.html")]
        [InlineData("T0792.html")]
        [InlineData("T0793.html")]
        [InlineData("T0794.html")]
        [InlineData("T0795.html")]
        [InlineData("T0802.html")]
        [InlineData("T0804.html")]
        [InlineData("T0805.html")]
        [InlineData("T0810.html")]
        [InlineData("T0812.html")]
        [InlineData("T0814.html")]
        [InlineData("T0820.html")]
        [InlineData("T0821.html")]
        [InlineData("T0830.html")]
        [InlineData("T0840.html")]
        [InlineData("T0850.html")]
        [InlineData("T0851.html")]
        [InlineData("T0860.html")]
        [InlineData("T0870.html")]
        [InlineData("T0880.html")]
        [InlineData("T0890.html")]
        [InlineData("T0900.html")]
        [InlineData("T0910.html")]
        [InlineData("T0920.html")]
        [InlineData("T0921.html")]
        [InlineData("T0922.html")]
        [InlineData("T0923.html")]
        [InlineData("T0924.html")]
        [InlineData("T0925.html")]
        [InlineData("T0926.html")]
        [InlineData("T0927.html")]
        [InlineData("T0928.html")]
        [InlineData("T0929.html")]
        [InlineData("T0930.html")]
        [InlineData("T0931.html")]
        [InlineData("T0932.html")]
        [InlineData("T0933.html")]
        [InlineData("T0934.html")]
        [InlineData("T0935.html")]
        [InlineData("T0936.html")]
        [InlineData("T0940.html")]
        [InlineData("T0945.html")]
        [InlineData("T0948.html")]
        [InlineData("T0950.html")]
        [InlineData("T0955.html")]
        [InlineData("T0960.html")]
        [InlineData("T0968.html")]
        [InlineData("T0970.html")]
        [InlineData("T0980.html")]
        [InlineData("T0990.html")]
        [InlineData("T1000.html")]
        [InlineData("T1010.html")]
        [InlineData("T1020.html")]
        [InlineData("T1030.html")]
        [InlineData("T1040.html")]
        [InlineData("T1050.html")]
        [InlineData("T1060.html")]
        [InlineData("T1070.html")]
        [InlineData("T1080.html")]
        [InlineData("T1100.html")]
        [InlineData("T1110.html")]
        [InlineData("T1111.html")]
        [InlineData("T1112.html")]
        [InlineData("T1120.html")]
        [InlineData("T1130.html")]
        [InlineData("T1131.html")]
        [InlineData("T1132.html")]
        [InlineData("T1140.html")]
        [InlineData("T1150.html")]
        [InlineData("T1160.html")]
        [InlineData("T1170.html")]
        [InlineData("T1180.html")]
        [InlineData("T1190.html")]
        [InlineData("T1200.html")]
        [InlineData("T1201.html")]
        [InlineData("T1210.html")]
        [InlineData("T1220.html")]
        [InlineData("T1230.html")]
        [InlineData("T1240.html")]
        [InlineData("T1241.html")]
        [InlineData("T1242.html")]
        [InlineData("T1250.html")]
        [InlineData("T1251.html")]
        [InlineData("T1260.html")]
        [InlineData("T1270.html")]
        [InlineData("T1280.html")]
        [InlineData("T1290.html")]
        [InlineData("T1300.html")]
        [InlineData("T1310.html")]
        [InlineData("T1320.html")]
        [InlineData("T1330.html")]
        [InlineData("T1340.html")]
        [InlineData("T1350.html")]
        [InlineData("T1360.html")]
        [InlineData("T1370.html")]
        [InlineData("T1380.html")]
        [InlineData("T1390.html")]
        [InlineData("T1400.html")]
        [InlineData("T1410.html")]
        [InlineData("T1420.html")]
        [InlineData("T1430.html")]
        [InlineData("T1440.html")]
        [InlineData("T1450.html")]
        [InlineData("T1460.html")]
        [InlineData("T1470.html")]
        [InlineData("T1480.html")]
        [InlineData("T1490.html")]
        [InlineData("T1500.html")]
        [InlineData("T1510.html")]
        [InlineData("T1520.html")]
        [InlineData("T1530.html")]
        [InlineData("T1540.html")]
        [InlineData("T1550.html")]
        [InlineData("T1560.html")]
        [InlineData("T1570.html")]
        [InlineData("T1580.html")]
        [InlineData("T1590.html")]
        [InlineData("T1591.html")]
        [InlineData("T1610.html")]
        [InlineData("T1620.html")]
        [InlineData("T1630.html")]
        [InlineData("T1640.html")]
        [InlineData("T1650.html")]
        [InlineData("T1660.html")]
        [InlineData("T1670.html")]
        [InlineData("T1680.html")]
        [InlineData("T1690.html")]
        [InlineData("T1700.html")]
        [InlineData("T1710.html")]
        public void HW001(string name)
        {
#if false
            string[] cssFilter = new[] {
                "text-indent",
                "margin-left",
                "margin-right",
                "padding-left",
                "padding-right",
            };
#else
            string[] cssFilter = null;
#endif

#if false
            string[] htmlFilter = new[] {
                "img",
            };
#else
            string[] htmlFilter = null;
#endif

            var sourceHtmlFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));
            var sourceImageDi = new DirectoryInfo(Path.Combine(TestUtil.SourceDir.FullName, sourceHtmlFi.Name.Replace(".html", "_files")));

            var destImageDi = new DirectoryInfo(Path.Combine(TestUtil.TempDir.FullName, sourceImageDi.Name));
            var sourceCopiedToDestHtmlFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceHtmlFi.Name.Replace(".html", "-1-Source.html")));
            var destCssFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceHtmlFi.Name.Replace(".html", "-2.css")));
            var destDocxFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceHtmlFi.Name.Replace(".html", "-3-ConvertedByHtmlToWml.docx")));
            var annotatedHtmlFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceHtmlFi.Name.Replace(".html", "-4-Annotated.txt")));

            File.Copy(sourceHtmlFi.FullName, sourceCopiedToDestHtmlFi.FullName);
            XElement html = HtmlToWmlReadAsXElement.ReadAsXElement(sourceCopiedToDestHtmlFi);

            string htmlString = html.ToString();
            if (htmlFilter != null && htmlFilter.Any())
            {
                bool found = false;
                foreach (var item in htmlFilter)
                {
                    if (htmlString.Contains(item))
                    {
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    sourceCopiedToDestHtmlFi.Delete();
                    return;
                }
            }

            string usedAuthorCss = HtmlToWmlConverter.CleanUpCss((string)html.Descendants().FirstOrDefault(d => d.Name.LocalName.ToLower() == "style"));
            File.WriteAllText(destCssFi.FullName, usedAuthorCss);

            if (cssFilter != null && cssFilter.Any())
            {
                bool found = false;
                foreach (var item in cssFilter)
                {
                    if (usedAuthorCss.Contains(item))
                    {
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    sourceCopiedToDestHtmlFi.Delete();
                    destCssFi.Delete();
                    return;
                }
            }

            if (sourceImageDi.Exists)
            {
                destImageDi.Create();
                foreach (var file in sourceImageDi.GetFiles())
                {
                    File.Copy(file.FullName, destImageDi.FullName + "/" + file.Name);
                }
            }
            FileInfo img1Fi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "img.png"));
            FileInfo img2Fi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "img2.png"));
            if (!img1Fi.Exists)
                File.Copy(Path.Combine(TestUtil.SourceDir.FullName, "img.png"), img1Fi.FullName);
            if (!img2Fi.Exists)
                File.Copy(Path.Combine(TestUtil.SourceDir.FullName, "img2.png"), img2Fi.FullName);

            HtmlToWmlConverterSettings settings = HtmlToWmlConverter.GetDefaultSettings();
            // image references in HTML files contain the path to the subdir that contains the images, so base URI is the name of the directory
            // that contains the HTML files
            settings.BaseUriForImages = Path.Combine(TestUtil.TempDir.FullName);

            WmlDocument doc = HtmlToWmlConverter.ConvertHtmlToWml(defaultCss, usedAuthorCss, userCss, html, settings, null, s_ProduceAnnotatedHtml ? annotatedHtmlFi.FullName : null);
            Assert.NotNull(doc);
            if (doc != null)
                SaveValidateAndFormatMainDocPart(destDocxFi, doc);

#if DO_CONVERSION_VIA_WORD
            var newAltChunkBeforeFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".html", "-5-AltChunkBefore.docx")));
            var newAltChunkAfterFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".html", "-6-ConvertedViaWord.docx")));
            WordAutomationUtilities.DoConversionViaWord(newAltChunkBeforeFi, newAltChunkAfterFi, html);
#endif
        }

        [Theory]
        [InlineData("E0010.html")]
        [InlineData("E0020.html")]
        public void HW004(string name)
        {

            var sourceHtmlFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));
            var sourceImageDi = new DirectoryInfo(Path.Combine(TestUtil.SourceDir.FullName, sourceHtmlFi.Name.Replace(".html", "_files")));

            var destImageDi = new DirectoryInfo(Path.Combine(TestUtil.TempDir.FullName, sourceImageDi.Name));
            var sourceCopiedToDestHtmlFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceHtmlFi.Name.Replace(".html", "-1-Source.html")));
            var destCssFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceHtmlFi.Name.Replace(".html", "-2.css")));
            var destDocxFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceHtmlFi.Name.Replace(".html", "-3-ConvertedByHtmlToWml.docx")));
            var annotatedHtmlFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceHtmlFi.Name.Replace(".html", "-4-Annotated.txt")));

            File.Copy(sourceHtmlFi.FullName, sourceCopiedToDestHtmlFi.FullName);
            XElement html = HtmlToWmlReadAsXElement.ReadAsXElement(sourceCopiedToDestHtmlFi);

            string usedAuthorCss = HtmlToWmlConverter.CleanUpCss((string)html.Descendants().FirstOrDefault(d => d.Name.LocalName.ToLower() == "style"));
            File.WriteAllText(destCssFi.FullName, usedAuthorCss);

            HtmlToWmlConverterSettings settings = HtmlToWmlConverter.GetDefaultSettings();
            settings.BaseUriForImages = Path.Combine(TestUtil.TempDir.FullName);

            Assert.Throws<OpenXmlPowerToolsException>(() => HtmlToWmlConverter.ConvertHtmlToWml(defaultCss, usedAuthorCss, userCss, html, settings, null, s_ProduceAnnotatedHtml ? annotatedHtmlFi.FullName : null));
        }

        private static void SaveValidateAndFormatMainDocPart(FileInfo destDocxFi, WmlDocument doc)
        {
            WmlDocument formattedDoc;

            doc.SaveAs(destDocxFi.FullName);
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

        static string userCss = @"";
    }
}
