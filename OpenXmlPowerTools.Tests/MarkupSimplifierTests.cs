// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OpenXmlPowerTools.Tests
{
    public class MarkupSimplifierTests
    {
        private const WordprocessingDocumentType DocumentType = WordprocessingDocumentType.Document;

        private const string SmartTagDocumentTextValue = "The countries include Algeria, Botswana, and Sri Lanka.";
        private const string SmartTagDocumentXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p >
      <w:r>
        <w:t xml:space=""preserve"">The countries include </w:t>
      </w:r>
      <w:smartTag w:uri=""urn:schemas-microsoft-com:office:smarttags"" w:element=""country-region"">
        <w:r>
          <w:t>Algeria</w:t>
        </w:r>
      </w:smartTag>
      <w:r>
        <w:t xml:space=""preserve"">, </w:t>
      </w:r>
      <w:smartTag w:uri=""urn:schemas-microsoft-com:office:smarttags"" w:element=""country-region"">
        <w:r>
          <w:t>Botswana</w:t>
        </w:r>
      </w:smartTag>
      <w:r>
        <w:t xml:space=""preserve"">, and </w:t>
      </w:r>
      <w:smartTag w:uri=""urn:schemas-microsoft-com:office:smarttags"" w:element=""country-region"">
        <w:smartTag w:uri=""urn:schemas-microsoft-com:office:smarttags"" w:element=""place"">
          <w:r>
            <w:t>Sri Lanka</w:t>
          </w:r>
        </w:smartTag>
      </w:smartTag>
      <w:r>
        <w:t>.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>
";

        private const string SdtDocumentXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:sdt>
      <w:sdtPr>
        <w:text/>
      </w:sdtPr>
      <w:sdtContent>
        <w:p>
          <w:r>
            <w:t>Hello World!</w:t>
          </w:r>
        </w:p>
      </w:sdtContent>
    </w:sdt>
  </w:body>
</w:document>";

        private const string GoBackBookmarkDocumentXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id=""0"" w:name=""_GoBack""/>
      <w:bookmarkEnd w:id=""0""/>
    </w:p>
  </w:body>
</w:document>";

        [Fact]
        public void CanRemoveSmartTags()
        {
            XDocument partDocument = XDocument.Parse(SmartTagDocumentXmlString);
            Assert.True(partDocument.Descendants(W.smartTag).Any());

            using (var stream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(partDocument);

                var settings = new SimplifyMarkupSettings { RemoveSmartTags = true };
                MarkupSimplifier.SimplifyMarkup(wordDocument, settings);

                partDocument = part.GetXDocument();
                XElement t = partDocument.Descendants(W.t).First();

                Assert.False(partDocument.Descendants(W.smartTag).Any());
                Assert.Equal(SmartTagDocumentTextValue, t.Value);
            }
        }

        [Fact]
        public void CanRemoveContentControls()
        {
            XDocument partDocument = XDocument.Parse(SdtDocumentXmlString);
            Assert.True(partDocument.Descendants(W.sdt).Any());

            using (var stream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(partDocument);

                var settings = new SimplifyMarkupSettings { RemoveContentControls = true };
                MarkupSimplifier.SimplifyMarkup(wordDocument, settings);

                partDocument = part.GetXDocument();
                XElement element = partDocument
                    .Descendants(W.body)
                    .Descendants()
                    .First();

                Assert.False(partDocument.Descendants(W.sdt).Any());
                Assert.Equal(W.p, element.Name);
            }
        }

        [Fact]
        public void CanRemoveGoBackBookmarks()
        {
            XDocument partDocument = XDocument.Parse(GoBackBookmarkDocumentXmlString);
            Assert.Contains(partDocument
                .Descendants(W.bookmarkStart)
, e => e.Attribute(W.name).Value == "_GoBack" && e.Attribute(W.id).Value == "0");
            Assert.Contains(partDocument
                .Descendants(W.bookmarkEnd)
, e => e.Attribute(W.id).Value == "0");

            using (var stream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(partDocument);

                var settings = new SimplifyMarkupSettings { RemoveGoBackBookmark = true };
                MarkupSimplifier.SimplifyMarkup(wordDocument, settings);

                partDocument = part.GetXDocument();
                Assert.False(partDocument.Descendants(W.bookmarkStart).Any());
                Assert.False(partDocument.Descendants(W.bookmarkEnd).Any());
            }
        }
    }
}

#endif
