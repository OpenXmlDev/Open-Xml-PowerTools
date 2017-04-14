using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OpenXmlPowerTools.Tests
{
    public class PtOpenXmlExtensionsTests
    {
        private const WordprocessingDocumentType DocumentType = WordprocessingDocumentType.Document;

        [Fact]
        public void MustReloadToSeeContentAddedThroughPowerTools()
        {
            // Do not automatically reload the root element on put.
            PtOpenXmlExtensions.ReloadRootElementOnPut = false;

            using (var stream = new MemoryStream())
            {
                CreateEmptyWordprocessingDocument(stream);

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, true))
                {
                    MainDocumentPart part = wordDocument.MainDocumentPart;

                    // Add a paragraph through the SDK.
                    Body body = part.Document.Body;
                    body.AppendChild(new Paragraph(new Run(new Text("Added through SDK"))));

                    // Save any changes made through the strongly typed SDK classes to the parts
                    // of the WordprocessingDocument. This can be done by invoking the Save method
                    // on the WordprocessingDocument, which will save all parts that had changes,
                    // or by invoking part.RootElement.Save() for the one part that was changed.
                    wordDocument.Save();

                    // Add a paragraph through the PowerTools.
                    XDocument content = part.GetXDocument();
                    XElement bodyElement = content.Descendants(W.body).First();
                    bodyElement.Add(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Added through PowerTools"))));
                    part.PutXDocument();

                    // Get the part's content through the SDK. However, we will only see what we
                    // added through the SDK, not what we added through the PowerTools functionality.
                    body = part.Document.Body;
                    List<Paragraph> paragraphs = body.Elements<Paragraph>().ToList();
                    Assert.Equal(1, paragraphs.Count);
                    Assert.Equal("Added through SDK", paragraphs[0].InnerText);

                    // Now, let's reload the root element of this one part.
                    // Reloading those root elements this way is fine if you know exactly which
                    // parts had their content changed by the Open XML PowerTools.
                    part.RootElement.Reload();

                    // Get the part's content through the SDK. Having reloaded the root element,
                    // we should now see both paragraphs.
                    body = part.Document.Body;
                    paragraphs = body.Elements<Paragraph>().ToList();
                    Assert.Equal(2, paragraphs.Count);
                    Assert.Equal("Added through SDK", paragraphs[0].InnerText);
                    Assert.Equal("Added through PowerTools", paragraphs[1].InnerText);
                }
            }
        }

        [Fact]
        public void CanReloadAutomaticallyOnPut()
        {
            // Automatically reload the root element on put.
            PtOpenXmlExtensions.ReloadRootElementOnPut = true;

            using (var stream = new MemoryStream())
            {
                CreateEmptyWordprocessingDocument(stream);

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, true))
                {
                    MainDocumentPart part = wordDocument.MainDocumentPart;

                    // Add a paragraph through the SDK.
                    Body body = part.Document.Body;
                    body.AppendChild(new Paragraph(new Run(new Text("Added through SDK"))));

                    // This demonstrates the use of the Save method invoked on the part's root
                    // element. In this case, we could have used part.Document.Save() as well.
                    // These methods are fine where you know exactly which parts were changed.
                    part.RootElement.Save();

                    // Add a paragraph through the PowerTools.
                    XDocument content = part.GetXDocument();
                    XElement bodyElement = content.Descendants(W.body).First();
                    bodyElement.Add(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Added through PowerTools"))));
                    part.PutXDocument();

                    // Get the part's content through the SDK. With automatic reload turned on, the
                    // we should see both paragraphs.
                    body = part.Document.Body;
                    List<Paragraph> paragraphs = body.Elements<Paragraph>().ToList();
                    Assert.Equal(2, paragraphs.Count);
                    Assert.Equal("Added through SDK", paragraphs[0].InnerText);
                    Assert.Equal("Added through PowerTools", paragraphs[1].InnerText);
                }
            }
        }

        [Fact]
        public void MustRemovePowerToolsAnnotationsToGetCorrectContent()
        {
            // Automatically reload the root element on put.
            PtOpenXmlExtensions.ReloadRootElementOnPut = true;

            using (var stream = new MemoryStream())
            {
                CreateEmptyWordprocessingDocument(stream);

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, true))
                {
                    MainDocumentPart part = wordDocument.MainDocumentPart;

                    // Add a first paragraph through the SDK.
                    Body body = part.Document.Body;
                    body.AppendChild(new Paragraph(new Run(new Text("First"))));

                    // This demonstrates the use of the Save method invoked on the MainDocumentPart's
                    // Document property, i.e., this part's RootElement.
                    part.Document.Save();

                    // Get content through the PowerTools. We will see the one paragraph added
                    // by using the strongly typed SDK classes.
                    XDocument content = part.GetXDocument();
                    List<XElement> paragraphElements = content.Descendants(W.p).ToList();
                    Assert.Equal(1, paragraphElements.Count);
                    Assert.Equal("First", paragraphElements[0].Value);

                    // Add a second paragraph through the SDK in the exact same way as above.
                    body = part.Document.Body;
                    body.AppendChild(new Paragraph(new Run(new Text("Second"))));
                    part.Document.Save();

                    // Get content through the PowerTools in the exact same way as above.
                    // What we will see, though, is that we still only get the first paragraph.
                    // This is caused by the GetXDocument method using the cached XDocument
                    // rather reading the part's stream again.
                    content = part.GetXDocument();
                    paragraphElements = content.Descendants(W.p).ToList();
                    Assert.Equal(1, paragraphElements.Count);
                    Assert.Equal("First", paragraphElements[0].Value);

                    // To make the GetXDocument read the parts' streams, we need to remove
                    // the annotations from the parts. In this case, we could have removed
                    // one annotation from one part. However, this doesn't work for all
                    // utilities provided by the PowerTools, as those will change multiple
                    // parts instead of just one.
                    wordDocument.RemovePowerToolsAnnotations();

                    // Get content through the PowerTools in the exact same way as above.
                    // We should now see both paragraphs.
                    content = part.GetXDocument();
                    paragraphElements = content.Descendants(W.p).ToList();
                    Assert.Equal(2, paragraphElements.Count);
                    Assert.Equal("First", paragraphElements[0].Value);
                    Assert.Equal("Second", paragraphElements[1].Value);
                }
            }
        }

        private static void CreateEmptyWordprocessingDocument(Stream stream)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.Document = new Document(new Body());
            }
        }
    }
}
