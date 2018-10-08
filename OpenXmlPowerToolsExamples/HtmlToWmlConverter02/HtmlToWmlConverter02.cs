// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    internal class Program
    {
        private static readonly string[] ProductNames =
        {
            "Unicycle",
            "Bicycle",
            "Tricycle",
            "Skateboard",
            "Roller Blades",
            "Hang Glider"
        };

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

            var templateDoc = new FileInfo("../../../TemplateDocument.docx");
            var dataFile = new FileInfo(Path.Combine(tempDi.FullName, "Data.xml"));

            // The following method generates a large data file with random data.
            // In a real world scenario, this is where you would query your data source and produce XML that will drive your document generation process.
            const int numberOfDocumentsToGenerate = 100;
            XElement data = GenerateDataFromDataSource(dataFile, numberOfDocumentsToGenerate);

            var wmlDoc = new WmlDocument(templateDoc.FullName);
            var count = 1;

            foreach (XElement customer in data.Elements("Customer"))
            {
                var assembledDoc = new FileInfo(Path.Combine(tempDi.FullName, $"Letter-{count++:0000}.docx"));
                Console.WriteLine("Generating {0}", assembledDoc.Name);
                WmlDocument wmlAssembledDoc = DocumentAssembler.AssembleDocument(wmlDoc, customer, out bool templateError);
                if (templateError)
                {
                    Console.WriteLine("Errors in template.");
                    Console.WriteLine("See {0} to determine the errors in the template.", assembledDoc.Name);
                }

                wmlAssembledDoc.SaveAs(assembledDoc.FullName);

                Console.WriteLine("Converting to HTML {0}", assembledDoc.Name);
                FileInfo htmlFileName = ConvertToHtml(assembledDoc.FullName, tempDi.FullName);

                Console.WriteLine("Converting back to DOCX {0}", htmlFileName.Name);
                ConvertToDocx(htmlFileName.FullName, tempDi.FullName);
            }
        }

        private static XElement GenerateDataFromDataSource(FileInfo dataFi, int numberOfDocumentsToGenerate)
        {
            var customers = new XElement("Customers");
            var r = new Random();

            for (var i = 0; i < numberOfDocumentsToGenerate; ++i)
            {
                var customer = new XElement("Customer",
                    new XElement("CustomerID", i + 1),
                    new XElement("Name", "Eric White"),
                    new XElement("HighValueCustomer", r.Next(2) == 0 ? "True" : "False"),
                    new XElement("Orders"));

                XElement orders = customer.Elements("Orders").First();
                int numberOfOrders = r.Next(10) + 1;

                for (var j = 0; j < numberOfOrders; j++)
                {
                    var order = new XElement("Order",
                        new XAttribute("Number", j + 1),
                        new XElement("ProductDescription", ProductNames[r.Next(ProductNames.Length)]),
                        new XElement("Quantity", r.Next(10)),
                        new XElement("OrderDate", "September 26, 2015"));

                    orders.Add(order);
                }

                customers.Add(customer);
            }

            customers.Save(dataFi.FullName);
            return customers;
        }

        public static FileInfo ConvertToHtml(string file, string outputDirectory)
        {
            var fi = new FileInfo(file);
            byte[] byteArray = File.ReadAllBytes(fi.FullName);

            using var memoryStream = new MemoryStream();
            memoryStream.Write(byteArray, 0, byteArray.Length);

            using WordprocessingDocument wDoc = WordprocessingDocument.Open(memoryStream, true);

            var destFileName = new FileInfo(fi.Name.Replace(".docx", ".html"));
            if (!string.IsNullOrEmpty(outputDirectory))
            {
                var di = new DirectoryInfo(outputDirectory);
                if (!di.Exists)
                {
                    throw new OpenXmlPowerToolsException("Output directory does not exist");
                }

                destFileName = new FileInfo(Path.Combine(di.FullName, destFileName.Name));
            }

            string imageDirectoryName = destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";
            var imageCounter = 0;

            string pageTitle = fi.FullName;
            CoreFilePropertiesPart part = wDoc.CoreFilePropertiesPart;
            if (part != null)
            {
                pageTitle = (string)part.GetXDocument().Descendants(DC.title).FirstOrDefault() ?? fi.FullName;
            }

            // TODO: Determine max-width from size of content area.
            var settings = new WmlToHtmlConverterSettings
            {
                AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
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
                    string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                    ImageFormat imageFormat = null;

                    if (extension == "png")
                    {
                        imageFormat = ImageFormat.Png;
                    }
                    else if (extension == "gif")
                    {
                        imageFormat = ImageFormat.Gif;
                    }
                    else if (extension == "bmp")
                    {
                        imageFormat = ImageFormat.Bmp;
                    }
                    else if (extension == "jpeg")
                    {
                        imageFormat = ImageFormat.Jpeg;
                    }
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
                    {
                        return null;
                    }

                    string imageFileName = imageDirectoryName + "/image" + imageCounter + "." + extension;
                    try
                    {
                        imageInfo.Bitmap.Save(imageFileName, imageFormat);
                    }
                    catch (ExternalException)
                    {
                        return null;
                    }

                    string imageSource = localDirInfo.Name + "/image" +
                                         imageCounter + "." + extension;

                    var img = new XElement(Xhtml.img,
                        new XAttribute(NoNamespace.src, imageSource),
                        imageInfo.ImgStyleAttribute,
                        imageInfo.AltText != null ? new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);

                    return img;
                }
            };

            XElement htmlElement = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

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

            return destFileName;
        }

        private static void ConvertToDocx(string file, string destinationDir)
        {
            var sourceHtmlFi = new FileInfo(file);
            var destDocxFi =
                new FileInfo(Path.Combine(destinationDir, sourceHtmlFi.Name.Replace(".html", "-ConvertedByHtmlToWml.docx")));

            XElement html = HtmlToWmlReadAsXElement.ReadAsXElement(sourceHtmlFi);

            string usedAuthorCss =
                HtmlToWmlConverter.CleanUpCss((string)html.Descendants()
                    .FirstOrDefault(d => d.Name.LocalName.ToLower() == "style"));

            HtmlToWmlConverterSettings settings = HtmlToWmlConverter.GetDefaultSettings();

            // image references in HTML files contain the path to the subdir that contains the images, so base URI is the name of the directory
            // that contains the HTML files
            settings.BaseUriForImages = sourceHtmlFi.DirectoryName;

            WmlDocument doc = HtmlToWmlConverter.ConvertHtmlToWml(DefaultCss, usedAuthorCss, UserCss, html, settings);
            doc.SaveAs(destDocxFi.FullName);
        }

        public class HtmlToWmlReadAsXElement
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
                html = (XElement)ConvertToNoNamespace(html);
                return html;
            }

            private static object ConvertToNoNamespace(XNode node)
            {
                if (node is XElement element)
                {
                    return new XElement(element.Name.LocalName,
                        element.Attributes().Where(a => !a.IsNamespaceDeclaration),
                        element.Nodes().Select(ConvertToNoNamespace));
                }

                return node;
            }
        }
    }
}
