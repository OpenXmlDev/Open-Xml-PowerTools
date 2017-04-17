using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace OpenXmlPowerTools.Tests
{
    public class OpenXmlRegexTests
    {
        private const WordprocessingDocumentType DocumentType = WordprocessingDocumentType.Document;

        private const string LeftDoubleQuotationMarks = @"[\u0022“„«»”]";
        private const string Words = @"[\w\-&/]+(?:\s[\w\-&/]+)*";
        private const string RightDoubleQuotationMarks = @"[\u0022”‟»«“]";

        private const string QuotationMarksDocumentXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:r>
        <w:t xml:space=""preserve"">Text can be enclosed in “normal double quotes” and in </w:t>
      </w:r>
      <w:r>
        <w:t>«</w:t>
      </w:r>
      <w:r>
        <w:t>double angle quotation marks</w:t>
      </w:r>
      <w:r>
        <w:t>»</w:t>
      </w:r>
      <w:r>
        <w:t>.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>";

        private const string QuotationMarksAndTrackedChangesDocumentXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:r>
        <w:t xml:space=""preserve"">Text can be enclosed in “normal </w:t>
      </w:r>
      <w:ins w:id=""8"" w:author=""Thomas Barnekow"" w:date=""2016-12-03T15:54:00Z"">
        <w:r>
          <w:t xml:space=""preserve"">double </w:t>
        </w:r>
      </w:ins>
      <w:r>
        <w:t xml:space=""preserve"">quotes” </w:t>
      </w:r>
      <w:del w:id=""9"" w:author=""Thomas Barnekow"" w:date=""2016-12-03T15:55:00Z"">
        <w:r>
          <w:delText xml:space=""preserve"">or </w:delText>
        </w:r>
      </w:del>
      <w:ins w:id=""10"" w:author=""Thomas Barnekow"" w:date=""2016-12-03T15:55:00Z"">
        <w:r>
          <w:t xml:space=""preserve"">and </w:t>
        </w:r>
      </w:ins>
      <w:r>
        <w:t xml:space=""preserve"">in </w:t>
      </w:r>
      <w:r>
        <w:t>«</w:t>
      </w:r>
      <w:r>
        <w:t xml:space=""preserve"">double </w:t>
      </w:r>
      <w:ins w:id=""11"" w:author=""Thomas Barnekow"" w:date=""2016-12-03T15:54:00Z"">
        <w:r>
          <w:t xml:space=""preserve"">angle </w:t>
        </w:r>
      </w:ins>
      <w:r>
        <w:t>quotation marks</w:t>
      </w:r>
      <w:r>
        <w:t>»</w:t>
      </w:r>
      <w:r>
        <w:t>.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>";

        private const string SymbolsAndTrackedChangesDocumentXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:r>
        <w:t xml:space=""preserve"">We can also use symbols such as </w:t>
      </w:r>
      <w:del w:id=""4"" w:author=""Thomas Barnekow"" w:date=""2017-04-16T12:31:00Z"">
        <w:r>
          <w:sym w:font=""Wingdings"" w:char=""F028""/>
        </w:r>
        <w:r>
          <w:delText xml:space=""preserve"">, </w:delText>
        </w:r>
      </w:del>
      <w:r>
        <w:sym w:font=""Wingdings"" w:char=""F021""/>
      </w:r>
      <w:r>
        <w:t xml:space=""preserve""> or </w:t>
      </w:r>
      <w:r>
        <w:sym w:font=""Wingdings"" w:char=""F028""/>
      </w:r>
      <w:r>
        <w:t>.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>";

        private static string InnerText(XContainer e)
        {
            return e.Descendants(W.r)
                .Where(r => r.Parent.Name != W.del)
                .Select(UnicodeMapper.RunToString)
                .StringConcatenate();
        }

        private static string InnerDelText(XContainer e)
        {
            return e.Descendants(W.delText)
                .Select(delText => delText.Value)
                .StringConcatenate();
        }

        [Fact]
        public void CanReplaceTextWithQuotationMarks()
        {
            XDocument partDocument = XDocument.Parse(QuotationMarksDocumentXmlString);
            XElement p = partDocument.Descendants(W.p).First();
            string innerText = InnerText(p);

            Assert.Equal(
                "Text can be enclosed in “normal double quotes” and in «double angle quotation marks».",
                innerText);

            using (var stream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(partDocument);

                IEnumerable<XElement> content = partDocument.Descendants(W.p);
                var regex = new Regex(string.Format("{0}(?<words>{1}){2}", LeftDoubleQuotationMarks, Words,
                    RightDoubleQuotationMarks));
                int count = OpenXmlRegex.Replace(content, regex, "‘changed ${words}’", null);

                p = partDocument.Descendants(W.p).First();
                innerText = InnerText(p);

                Assert.Equal(2, count);
                Assert.Equal(
                    "Text can be enclosed in ‘changed normal double quotes’ and in ‘changed double angle quotation marks’.",
                    innerText);
            }
        }

        [Fact]
        public void CanReplaceTextWithQuotationMarksAndAddTrackedChangesWhenReplacing()
        {
            XDocument partDocument = XDocument.Parse(QuotationMarksDocumentXmlString);
            XElement p = partDocument.Descendants(W.p).First();
            string innerText = InnerText(p);

            Assert.Equal(
                "Text can be enclosed in “normal double quotes” and in «double angle quotation marks».",
                innerText);

            using (var stream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(partDocument);

                IEnumerable<XElement> content = partDocument.Descendants(W.p);
                var regex = new Regex(string.Format("{0}(?<words>{1}){2}", LeftDoubleQuotationMarks, Words,
                    RightDoubleQuotationMarks));
                int count = OpenXmlRegex.Replace(content, regex, "‘changed ${words}’", null, true, "John Doe");

                p = partDocument.Descendants(W.p).First();
                innerText = InnerText(p);

                Assert.Equal(2, count);
                Assert.Equal(
                    "Text can be enclosed in ‘changed normal double quotes’ and in ‘changed double angle quotation marks’.",
                    innerText);

                Assert.True(p.Elements(W.ins).Any(e => InnerText(e) == "‘changed normal double quotes’"));
                Assert.True(p.Elements(W.ins).Any(e => InnerText(e) == "‘changed double angle quotation marks’"));

                Assert.True(p.Elements(W.del).Any(e => InnerDelText(e) == "“normal double quotes”"));
                Assert.True(p.Elements(W.del).Any(e => InnerDelText(e) == "«double angle quotation marks»"));
            }
        }

        [Fact]
        public void CanReplaceTextWithQuotationMarksAndTrackedChanges()
        {
            XDocument partDocument = XDocument.Parse(QuotationMarksAndTrackedChangesDocumentXmlString);
            XElement p = partDocument.Descendants(W.p).First();
            string innerText = InnerText(p);

            Assert.Equal(
                "Text can be enclosed in “normal double quotes” and in «double angle quotation marks».",
                innerText);

            using (var stream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(partDocument);

                IEnumerable<XElement> content = partDocument.Descendants(W.p);
                var regex = new Regex(string.Format("{0}(?<words>{1}){2}", LeftDoubleQuotationMarks, Words,
                    RightDoubleQuotationMarks));
                int count = OpenXmlRegex.Replace(content, regex, "‘changed ${words}’", null, true, "John Doe");

                p = partDocument.Descendants(W.p).First();
                innerText = InnerText(p);

                Assert.Equal(2, count);
                Assert.Equal(
                    "Text can be enclosed in ‘changed normal double quotes’ and in ‘changed double angle quotation marks’.",
                    innerText);

                Assert.True(p.Elements(W.ins).Any(e => InnerText(e) == "‘changed normal double quotes’"));
                Assert.True(p.Elements(W.ins).Any(e => InnerText(e) == "‘changed double angle quotation marks’"));
            }
        }

        [Fact]
        public void CanReplaceTextWithSymbolsAndTrackedChanges()
        {
            XDocument partDocument = XDocument.Parse(SymbolsAndTrackedChangesDocumentXmlString);
            XElement p = partDocument.Descendants(W.p).First();
            string innerText = InnerText(p);

            Assert.Equal("We can also use symbols such as \uF021 or \uF028.", innerText);

            using (var stream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(partDocument);

                IEnumerable<XElement> content = partDocument.Descendants(W.p);
                var regex = new Regex(@"[\uF021]");
                int count = OpenXmlRegex.Replace(content, regex, "\uF028", null, true, "John Doe");

                p = partDocument.Descendants(W.p).First();
                innerText = InnerText(p);

                Assert.Equal(1, count);
                Assert.Equal("We can also use symbols such as \uF028 or \uF028.", innerText);

                Assert.True(p.Descendants(W.ins).Any(
                    ins => ins.Descendants(W.sym).Any(
                        sym => sym.Attribute(W.font).Value == "Wingdings" && 
                               sym.Attribute(W._char).Value == "F028")));
            }
        }
    }
}
