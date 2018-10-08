//
// Copyright 2017-2018 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas@barnekow.info
//

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OpenXmlPowerTools.Tests
{
    /// <summary>
    /// Base class for unit tests providing utility methods.
    /// </summary>
    public class TestsBase
    {
        private const WordprocessingDocumentType DocumentType = WordprocessingDocumentType.Document;

        private const string StylesXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<w:styles xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:style w:type=""paragraph"" w:default=""1"" w:styleId=""Normal"">
    <w:name w:val=""Normal""/>
    <w:qFormat/>
  </w:style>
  <w:style w:type=""character"" w:default=""1"" w:styleId=""DefaultParagraphFont"">
    <w:name w:val=""Default Paragraph Font""/>
    <w:uiPriority w:val=""1""/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
  </w:style>
  <w:style w:type=""table"" w:default=""1"" w:styleId=""TableNormal"">
    <w:name w:val=""Normal Table""/>
    <w:uiPriority w:val=""99""/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:tblPr>
      <w:tblInd w:w=""0"" w:type=""dxa""/>
      <w:tblCellMar>
        <w:top w:w=""0"" w:type=""dxa""/>
        <w:left w:w=""108"" w:type=""dxa""/>
        <w:bottom w:w=""0"" w:type=""dxa""/>
        <w:right w:w=""108"" w:type=""dxa""/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>
  <w:style w:type=""numbering"" w:default=""1"" w:styleId=""NoList"">
    <w:name w:val=""No List""/>
    <w:uiPriority w:val=""99""/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
  </w:style>
  <w:style w:type=""table"" w:styleId=""TableGrid"">
    <w:name w:val=""Table Grid""/>
    <w:basedOn w:val=""TableNormal""/>
    <w:uiPriority w:val=""39""/>
    <w:pPr>
      <w:spacing w:after=""0"" w:line=""240"" w:lineRule=""auto""/>
    </w:pPr>
    <w:tblPr>
      <w:tblBorders>
        <w:top w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""auto""/>
        <w:left w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""auto""/>
        <w:bottom w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""auto""/>
        <w:right w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""auto""/>
        <w:insideH w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""auto""/>
        <w:insideV w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""auto""/>
      </w:tblBorders>
    </w:tblPr>
  </w:style>
</w:styles>";

        private static readonly string[] ExpectedValidationErrors =
        {
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:allStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:alternateStyleNames' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:clearFormatting' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:customStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnRuns' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnParagraphs' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnNumbering' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnTables' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowFirstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowLastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:headingStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowFirstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowLastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:latentStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:numberingStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:stylesInUse' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:tableStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:top3HeadingStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:visibleStyles' attribute is not declared.",
            "The 'urn:schemas-microsoft-com:office:office:gfxdata' attribute is not declared.",
            "The attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:name' has invalid value 'useWord2013TrackBottomHyphenation'. The Enumeration constraint failed.",
            "The attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:val' has invalid value '0'. The MinInclusive constraint failed. The value must be greater than or equal to 1.",
            "The attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:val' has invalid value '0'. The MinInclusive constraint failed. The value must be greater than or equal to 2.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:fill' attribute is invalid - The value '0' is not valid according to any of the memberTypes of the union."
        };

        /// <summary>
        /// Creates an empty <see cref="WordprocessingDocument" /> with a <see cref="MainDocumentPart" />
        /// on the given <see cref="Stream" />.
        /// </summary>
        /// <param name="stream">The stream.</param>
        protected static void CreateEmptyWordprocessingDocument(Stream stream)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.Document = new Document(new Body());
            }
        }

        /// <summary>
        /// Creates a new <see cref="WordprocessingDocument" /> on a <see cref="MemoryStream" />,
        /// having the given <see cref="XDocument" /> as the main document part's content and
        /// adding a default style definitions part.
        /// </summary>
        /// <param name="document">The main document part's content.</param>
        /// <returns>The new <see cref="WordprocessingDocument" />.</returns>
        protected static WordprocessingDocument CreateWordprocessingDocument(XDocument document)
        {
            return CreateWordprocessingDocument(new MemoryStream(), document);
        }

        /// <summary>
        /// Creates a new <see cref="WordprocessingDocument" /> on the given <see cref="Stream" />,
        /// having the given <see cref="XDocument" /> as the main document part's content and
        /// adding a default style definitions part.
        /// </summary>
        /// <param name="stream">The <see cref="Stream" />.</param>
        /// <param name="document">The <see cref="XDocument" />.</param>
        /// <returns>The new <see cref="WordprocessingDocument" />.</returns>
        protected static WordprocessingDocument CreateWordprocessingDocument(Stream stream, XDocument document)
        {
            WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType);

            MainDocumentPart mainDocumentPart = wordDocument.AddMainDocumentPart();
            mainDocumentPart.PutXDocument(document);

            var styleDefinitionsPart = mainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            styleDefinitionsPart.PutXDocument(XDocument.Parse(StylesXml));

            return wordDocument;
        }

        /// <summary>
        /// Creates a <see cref="WmlDocument" /> that will have (1) a <see cref="MainDocumentPart" />
        /// with the given <paramref name="document" /> as its <see cref="XDocument" /> and (2) a
        /// <see cref="StyleDefinitionsPart" /> with an empty w:styles root element.
        /// </summary>
        /// <param name="fileName">The <see cref="WmlDocument" />'s file name.</param>
        /// <param name="document">The <see cref="MainDocumentPart" />'s <see cref="XDocument" />.</param>
        /// <returns>The new <see cref="WmlDocument" />.</returns>
        protected static WmlDocument CreateWmlDocument(string fileName, XDocument document)
        {
            var stream = new MemoryStream();
            WordprocessingDocument wordDocument = CreateWordprocessingDocument(stream, document);
            wordDocument.Close();

            return new WmlDocument(fileName, stream);
        }

        /// <summary>
        /// Validate the given <see cref="WmlDocument" />, using the <see cref="OpenXmlValidator" />.
        /// </summary>
        /// <param name="wmlDocument">The <see cref="WmlDocument" />.</param>
        protected static void Validate(WmlDocument wmlDocument)
        {
            using (var memoryStreamDocument = new OpenXmlMemoryStreamDocument(wmlDocument))
            using (WordprocessingDocument wordDocument = memoryStreamDocument.GetWordprocessingDocument())
            {
                Validate(wordDocument);
            }
        }

        /// <summary>
        /// Validate the given <see cref="WordprocessingDocument" />, using the <see cref="OpenXmlValidator" />.
        /// </summary>
        /// <param name="wordDocument">The <see cref="WordprocessingDocument" />.</param>
        protected static void Validate(WordprocessingDocument wordDocument)
        {
            var validator = new OpenXmlValidator();
            IList<ValidationErrorInfo> errorInfos = validator
                .Validate(wordDocument)
                .Where(e => !ExpectedValidationErrors.Contains(e.Description))
                .ToList();

            if (errorInfos.Any())
            {
                const string ind = "  ";
                var sb = new StringBuilder();

                foreach (ValidationErrorInfo errorInfo in errorInfos)
                {
                    sb.Append("Error" + Environment.NewLine);
                    sb.Append(ind + "ErrorType: " + errorInfo.ErrorType + Environment.NewLine);
                    sb.Append(ind + "Description: " + errorInfo.Description + Environment.NewLine);
                    sb.Append(ind + "Part: " + errorInfo.Part.Uri + Environment.NewLine);
                    sb.Append(ind + "XPath: " + errorInfo.Path.XPath + Environment.NewLine);
                }

                string validationErrors = sb.ToString();
                if (validationErrors != "")
                {
                    Assert.True(false, validationErrors);
                }
            }
        }
    }
}
