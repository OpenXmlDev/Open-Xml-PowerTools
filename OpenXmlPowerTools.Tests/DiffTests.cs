using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using AngleSharp.Xhtml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PreMailer.Net;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Xunit;

namespace OpenXmlPowerTools.Tests
{
    public class DiffTests
    {
        private const string CUSTOMCSS = @"
ins, del{
  text-decoration: none;
  padding: 0 .3em;
  border-radius: .3em;
  text-indent: 0;
  display: inline-block;
}
del img, ins img {
  opacity: 0.5;
}
img {
  max-width: 100%;
  max-height: 100%;
}
ins {
  background: #83d5a8;
  -webkit-box-decoration-break: clone;
  -o-box-decoration-break: clone;
  box-decoration-break: clone;
}
del {
  background: rgba(231, 76, 60,.5);
}";

        [Fact]
        public void CompleteDiffTests()
        {

            List<XElement> diffElements = new List<XElement>();

            Stream oldFile = File.Open("./TestData/Document1.docx", FileMode.Open);
            Stream newFile = File.Open("./TestData/Document2.docx", FileMode.Open);

            using WordprocessingDocument oldDocument = WordprocessingDocument.Open(oldFile, false);
            using WordprocessingDocument newDocument = WordprocessingDocument.Open(newFile, false);

            using MemoryStream newDocumentMS = new MemoryStream();
            newDocument.WriteTo(newDocumentMS);

            // header diff
            IEnumerable<XElement>? headerDiff = GetDiffFromTwoParts(oldDocument.MainDocumentPart.HeaderParts, newDocument.MainDocumentPart.HeaderParts, newDocumentMS);
            if (headerDiff.Any())
            {
                diffElements.AddRange(headerDiff);
                diffElements.Add(new XElement(XhtmlNoNamespace.hr));
            }

            // content diff
            WmlDocument result = GetRevisionResult(oldDocument, newDocument);
            XElement? htmlResult = GetHtml(result);
            diffElements.AddRange(GetDiffElementsFromHtmlElement(htmlResult));

            // footer diff
            IEnumerable<XElement>? footerDiff = GetDiffFromTwoParts(oldDocument.MainDocumentPart.FooterParts, newDocument.MainDocumentPart.FooterParts, newDocumentMS);
            if (headerDiff.Any())
            {
                diffElements.AddRange(footerDiff);
                diffElements.Add(new XElement(XhtmlNoNamespace.hr));
            }

            StringBuilder sb = new StringBuilder();

            foreach (XElement item in diffElements)
            {
                sb.Append(item.ToString());
            }

            string resultHtml = sb.ToString();

            Assert.NotNull(resultHtml);
        }

        public static IEnumerable<XElement> GetDiffElementsFromHtmlElement(XElement htmlXElement)
        {
            // inline the complete css (from //html/head/styles)
            InlineResult preMailerResult = PreMailer.Net.PreMailer.MoveCssInline(htmlXElement.ToString(), false, "meta", CUSTOMCSS, true, true);

            if (preMailerResult.Warnings.Count > 0)
            {
                Debug.WriteLine(preMailerResult.Warnings);
            }

            // Premailer returns html (which allows <meta attr="value"> without a fronstlash at the end) and we have to get xhtml
            // convert to xhtml using AngleSharp
            HtmlParser parser = new HtmlParser();
            IHtmlDocument doc = parser.ParseDocument(preMailerResult.Html);

            using StringWriter sw = new StringWriter();
            doc.ToHtml(sw, XhtmlMarkupFormatter.Instance);

            // CUSTOM_CSS formats ins and del elements; since they add a background color to indicate addition/deletion we have to remove
            // the background-stlye for all descendants of ins-, or del-elements
            XElement htmlWithInlinedCss = XElement.Parse(sw.ToString());

            IEnumerable<XElement> allInsElements = htmlWithInlinedCss.Descendants().Elements(Xhtml.ins);
            IEnumerable<XElement> allDelElements = htmlWithInlinedCss.Descendants().Elements(Xhtml.del);

            // delete background from all descendants of inserted/deleted elements (ensures that the according background from this project is shown)
            RemoveAllBackgroundStyles(allInsElements.Concat(allDelElements));

            // remove all font-family declarations from paragraphs (spans below them inherited from the paragraph)
            // TODO: should get fixed in Open-Xml-PowerTools (del/ins element does not get the font-family added, because it's not a span/div...)
            RemoveAllFontFamiliesFromParagraphs(allInsElements.Concat(allDelElements));

            return htmlWithInlinedCss.Element(Xhtml.body)!.Elements();
        }

        private static void RemoveAllFontFamiliesFromParagraphs(IEnumerable<XElement> elements)
        {
            foreach (XElement el in elements)
            {
                if (el.Parent!.Name == Xhtml.p || el.Parent.Name == XhtmlNoNamespace.p)
                {
                    XElement parent = el.Parent; // paragraph element
                    XAttribute attr;
                    if ((attr = parent.Attribute(XhtmlNoNamespace.style)!) != null)
                    {
                        attr.Value = Regex.Replace(attr.Value, @"font-family:?.Symbol;", string.Empty);
                    }

                    if ((attr = parent.Attribute(Xhtml.style)!) != null)
                    {
                        attr.Value = Regex.Replace(attr.Value, @"font-family:?.Symbol;", string.Empty);
                    }
                }
            }
        }

        private static void RemoveAllBackgroundStyles(IEnumerable<XElement> elements)
        {
            foreach (XElement el in elements)
            {
                el.Descendants().ToList().ForEach(e =>
                {
                    XAttribute attr;
                    if ((attr = e.Attribute(XhtmlNoNamespace.style)!) != null)
                    {
                        attr.Value = Regex.Replace(attr.Value, @"background:[^;]*;", string.Empty);
                    }

                    if ((attr = e.Attribute(Xhtml.style)!) != null)
                    {
                        attr.Value = Regex.Replace(attr.Value, @"background:[^;]*;", string.Empty);
                    }
                });
            }
        }

        public static XElement GetHtml(WmlDocument document)
        {
            WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings() { AcceptRevisions = false }; // IMPORTANT: do not accept revisions

            static XElement ImageHandler(ImageInfo imageInfo)
            {
                try
                {
                    MemoryStream ms = new MemoryStream();

                    imageInfo.Bitmap.Save(ms, imageInfo.Bitmap.RawFormat);

                    byte[] byteImage = ms.ToArray();

                    string base64String = Convert.ToBase64String(byteImage);

                    return new XElement(
                        Xhtml.img,
                        new XAttribute(NoNamespace.src, "data:" + imageInfo.ContentType + ";base64," + base64String),
                        new XAttribute(NoNamespace.alt, imageInfo.AltText));
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    return null;
                }

            }

            settings.ImageHandler = ImageHandler;

            return WmlToHtmlConverter.ConvertToHtml(document, settings);
        }

        private static IEnumerable<XElement> GetDiffFromTwoParts(IEnumerable<OpenXmlPart> oldParts, IEnumerable<OpenXmlPart> newParts, MemoryStream compareContainerDocument)
        {
            WmlDocument? comparerResult = GetDiffResult(compareContainerDocument, oldParts, newParts);

            if (comparerResult == null)
            {
                return new List<XElement>();
            }

            XElement? headerResult = GetHtml(comparerResult);

            return GetDiffElementsFromHtmlElement(headerResult);
        }

        public static WordprocessingDocument GetDocumentWithContent(Stream templateStream, IEnumerable<XElement> newContentElement)
        {
            WordprocessingDocument tempDoc = WordprocessingDocument.Open(templateStream, true);
            tempDoc.MainDocumentPart.Document.Body.RemoveAllChildren();

            Body body = new Body();
            foreach (XElement element in newContentElement)
            {
                if (element != null)
                {
                    body.Append(ToOpenXmlElement(element));
                }
            }

            // other replacing methods somehow do not remove all children. pay ATTENTION when changing something in this method
            tempDoc.MainDocumentPart.RootElement.ReplaceChild<Body>(body, tempDoc.MainDocumentPart.Document.Body);

            tempDoc.Save();

            return tempDoc;
        }

        public static OpenXmlElement ToOpenXmlElement(XElement xe)
        {
            using StreamWriter sw = new StreamWriter(new MemoryStream());
            sw.Write(xe.ToString());
            sw.Flush();
            sw.BaseStream.Seek(0, SeekOrigin.Begin);

            using TypedOpenXmlPartReader re = new TypedOpenXmlPartReader(sw.BaseStream);

            re.Read();
            OpenXmlElement oxe = re.LoadCurrentElement();
            re.Close();

            return oxe;
        }

        private static WmlDocument? GetDiffResult(MemoryStream compareContainerDocument, IEnumerable<OpenXmlPart> oldParts, IEnumerable<OpenXmlPart> newParts)
        {

#pragma warning disable CA2000 // Dispose objects before losing scope
            WordprocessingDocument? oldDocHeadersDocument = GetDocumentWithContent(compareContainerDocument, GetPartsAsXDocument(oldParts));
            WordprocessingDocument? newDocHeadersDocument = GetDocumentWithContent(compareContainerDocument, GetPartsAsXDocument(newParts));
#pragma warning restore CA2000 // Dispose objects before losing scope

            // return empty list if both documents contain no children
            if (!oldDocHeadersDocument.MainDocumentPart.Document.Body.HasChildren && !newDocHeadersDocument.MainDocumentPart.Document.Body.HasChildren)
            {
                return null;
            }

            if (!oldDocHeadersDocument.MainDocumentPart.Document.Body.HasChildren)
            {
                oldDocHeadersDocument.MainDocumentPart.Document.Body.AppendChild(new Paragraph());
            }

            if (!newDocHeadersDocument.MainDocumentPart.Document.Body.HasChildren)
            {
                oldDocHeadersDocument.MainDocumentPart.Document.Body.AppendChild(new Paragraph());
            }

            return GetRevisionResult(oldDocHeadersDocument, newDocHeadersDocument);
        }

        private static IEnumerable<XElement> GetPartsAsXDocument(IEnumerable<OpenXmlPart> openXmlPackages)
        {
            return openXmlPackages.Select(el =>
            {
                return el.GetXDocument().Root;
            });
        }

        private static WmlDocument GetRevisionResult(WordprocessingDocument oldFile, WordprocessingDocument newFile)
        {
            WmlComparerSettings settings = new WmlComparerSettings()
            {
                DetailThreshold = 0.25,
            };

            using MemoryStream oldStream = GetMemoryStream(oldFile);
            WmlDocument? oldDocument = new WmlDocument("old.docx", oldStream);

            using MemoryStream newStream = GetMemoryStream(newFile);
            WmlDocument? newDocument = new WmlDocument("new.docx", newStream);

            return WmlComparer.Compare(oldDocument, newDocument, settings);
        }

        public static MemoryStream GetMemoryStream(WordprocessingDocument wordprocessingDocument)
        {
            MemoryStream ms = new MemoryStream();
            wordprocessingDocument.WriteTo(ms);

            return ms;
        }

        
    }

    public static class OpenXmlPackageExtensions
    {

        /// <summary>
        /// If a document is opened using an <see cref="MemoryStream"/> not all changes to the document
        /// will be stored back to that stream. For example, if an image is removed, the imange itself
        /// will be still a part of the <see cref="MemoryStream"/>. For that reason we have to
        /// save the document to a different stream using the <see cref="OpenXmlPackage.Clone(Stream, bool)"/>
        /// method. This is what this method is doing. Unfortunately, it is not possible to
        /// write asynchronously to that stream.
        /// </summary>
        /// <param name="openXmlPackage">The <see cref="OpenXmlPackage"/>.</param>
        /// <param name="stream">The <see cref="Stream"/> to which the document should be stored to.</param>
        public static void WriteTo(this OpenXmlPackage openXmlPackage, Stream stream)
        {
            using (openXmlPackage.Clone(stream, false))
            {
            }
        }

    }
}
