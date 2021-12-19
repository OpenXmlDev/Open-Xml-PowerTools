// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OpenXmlPowerTools;

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
 * We don't include the HtmlAgilityPack in Open-Xml-PowerTools, to simplify installation.  The example files
 * in this module do not require HtmlAgilityPack to run.
*******************************************************************************************/

internal static class HtmlToWmlConverterExample
{
    private const string DefaultCss = @"html, address,
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

    private const string UserCss = @"";

    private static void Main()
    {
        DateTime n = DateTime.Now;

        var tempDi = new DirectoryInfo(
            $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");

        tempDi.Create();

        foreach (string file in Directory.GetFiles("../../../", "*.html") /* .Where(f => f.Contains("Test-01")) */)
        {
            ConvertToDocx(file, tempDi.FullName);
        }
    }

    private static void ConvertToDocx(string file, string destinationDir)
    {
        var sourceHtmlFile = new FileInfo(file);
        Console.WriteLine("Converting " + sourceHtmlFile.Name);

        var destCssFi = new FileInfo(Path.Combine(destinationDir, sourceHtmlFile.Name.Replace(".html", "-2.css")));

        var destDocxFi =
            new FileInfo(Path.Combine(destinationDir, sourceHtmlFile.Name.Replace(".html", "-3-ConvertedByHtmlToWml.docx")));

        var annotatedHtmlFi =
            new FileInfo(Path.Combine(destinationDir, sourceHtmlFile.Name.Replace(".html", "-4-Annotated.txt")));

        XElement html = HtmlToWmlReadAsXElement.ReadAsXElement(sourceHtmlFile);

        string usedAuthorCss =
            HtmlToWmlConverter.CleanUpCss((string) html.Descendants().FirstOrDefault(d => d.Name.LocalName.ToLower() == "style"));

        File.WriteAllText(destCssFi.FullName, usedAuthorCss);

        HtmlToWmlConverterSettings settings = HtmlToWmlConverter.GetDefaultSettings();

        // image references in HTML files contain the path to the subdir that contains the images,
        // so base URI is the name of the directory that contains the HTML files
        settings.BaseUriForImages = sourceHtmlFile.DirectoryName;

        WmlDocument doc = HtmlToWmlConverter.ConvertHtmlToWml(DefaultCss, usedAuthorCss, UserCss, html, settings, null,
            annotatedHtmlFi.FullName);

        doc.SaveAs(destDocxFi.FullName);
    }

    private static class HtmlToWmlReadAsXElement
    {
        public static XElement ReadAsXElement(FileInfo sourceHtmlFi)
        {
            string htmlString = File.ReadAllText(sourceHtmlFi.FullName);

#if USE_HTMLAGILITYPACK
            XElement html;

            try
            {
                html = XElement.Parse(htmlString);
            }
            catch (XmlException)
            {
                HtmlDocument hdoc = new HtmlDocument();
                hdoc.Load(sourceHtmlFi.FullName, Encoding.Default);
                hdoc.OptionOutputAsXml = true;
                hdoc.Save(sourceHtmlFi.FullName, Encoding.Default);
                StringBuilder sb = new StringBuilder(File.ReadAllText(sourceHtmlFi.FullName, Encoding.Default));
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
                File.WriteAllText(sourceHtmlFi.FullName, sb.ToString(), Encoding.Default);
                html = XElement.Parse(sb.ToString());
            }
#else
            XElement html = XElement.Parse(htmlString);
#endif
            // HtmlToWmlConverter expects the HTML elements to be in no namespace, so convert all elements to no namespace.
            html = (XElement) ConvertToNoNamespace(html);
            return html;
        }

        private static object ConvertToNoNamespace(XNode node)
        {
            return node is XElement element
                ? new XElement(element.Name.LocalName,
                    element.Attributes().Where(a => !a.IsNamespaceDeclaration),
                    element.Nodes().Select(ConvertToNoNamespace))
                : node;
        }
    }
}
