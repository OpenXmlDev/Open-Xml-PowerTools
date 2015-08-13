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

namespace OpenXmlPowerTools.Tests
{
    public class DocumentAssemblerTests
    {
        [Theory]
        [InlineData("DA001-TemplateDocument.docx", "DA-Data.xml", false)]
        [InlineData("DA002-TemplateDocument.docx", "DA-DataNotHighValueCust.xml", false)]
        [InlineData("DA003-Select-XPathFindsNoData.docx", "DA-Data.xml", true)]
        [InlineData("DA004-Select-XPathFindsNoDataOptional.docx", "DA-Data.xml", false)]
        [InlineData("DA005-SelectRowData-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA006-SelectTestValue-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA007-SelectRepeatingData-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA008-TableElementWithNoTable.docx", "DA-Data.xml", true)]
        [InlineData("DA009-InvalidXPath.docx", "DA-Data.xml", true)]
        [InlineData("DA010-InvalidXml.docx", "DA-Data.xml", true)]
        [InlineData("DA011-SchemaError.docx", "DA-Data.xml", true)]
        [InlineData("DA012-OtherMarkupTypes.docx", "DA-Data.xml", true)]
        [InlineData("DA013-Runs.docx", "DA-Data.xml", false)]
        [InlineData("DA014-TwoRuns-NoValuesSelected.docx", "DA-Data.xml", true)]
        [InlineData("DA015-TwoRunsXmlExceptionInFirst.docx", "DA-Data.xml", true)]
        [InlineData("DA016-TwoRunsSchemaErrorInSecond.docx", "DA-Data.xml", true)]
        [InlineData("DA017-FiveRuns.docx", "DA-Data.xml", true)]
        [InlineData("DA018-SmartQuotes.docx", "DA-Data.xml", false)]
        [InlineData("DA019-RunIsEntireParagraph.docx", "DA-Data.xml", false)]
        [InlineData("DA020-TwoRunsAndNoOtherContent.docx", "DA-Data.xml", true)]
        [InlineData("DA021-NestedRepeat.docx", "DA-DataNestedRepeat.xml", false)]
        [InlineData("DA022-InvalidXPath.docx", "DA-Data.xml", true)]
        [InlineData("DA023-RepeatWOEndRepeat.docx", "DA-Data.xml", true)]
        [InlineData("DA026-InvalidRootXmlElement.docx", "DA-Data.xml", true)]
        [InlineData("DA027-XPathErrorInPara.docx", "DA-Data.xml", true)]
        [InlineData("DA028-NoPrototypeRow.docx", "DA-Data.xml", true)]
        [InlineData("DA029-NoDataForCell.docx", "DA-Data.xml", true)]
        [InlineData("DA030-TooMuchDataForCell.docx", "DA-TooMuchDataForCell.xml", true)]
        [InlineData("DA031-CellDataInAttributes.docx", "DA-CellDataInAttributes.xml", true)]
        [InlineData("DA032-TooMuchDataForConditional.docx", "DA-TooMuchDataForConditional.xml", true)]
        [InlineData("DA033-ConditionalOnAttribute.docx", "DA-ConditionalOnAttribute.xml", false)]
        [InlineData("DA034-HeaderFooter.docx", "DA-Data.xml", false)]
        [InlineData("DA100-TemplateDocument.docx", "DA-Data.xml", false)]
        [InlineData("DA101-TemplateDocument.docx", "DA-Data.xml", true)]
        [InlineData("DA102-TemplateDocument.docx", "DA-Data.xml", true)]
        
        public void DA101_DocumentAssembler(string templateDocumentName, string dataFileName, bool templateError)
        {
            FileInfo templateDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, templateDocumentName));
            FileInfo dataFile = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, dataFileName));

            WmlDocument wmlTemplate = new WmlDocument(templateDocx.FullName);
            XElement data = XElement.Load(dataFile.FullName);

            bool returnedTemplateError;
            WmlDocument afterAssembling = DocumentAssembler.AssembleDocument(wmlTemplate, data, out returnedTemplateError);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, templateDocx.Name.Replace(".docx", "-processed-by-DocumentAssembler.docx")));
            afterAssembling.SaveAs(assembledDocx.FullName);

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(afterAssembling.DocumentByteArray, 0, afterAssembling.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    OpenXmlValidator v = new OpenXmlValidator();
                    var valErrors = v.Validate(wDoc).Where(ve => !s_ExpectedErrors.Contains(ve.Description));

#if true
                    StringBuilder sb = new StringBuilder();
                    foreach (var item in valErrors.Select(r => r.Description).OrderBy(t => t).Distinct())
	                {
		                sb.Append(item).Append(Environment.NewLine);
	                }
                    string z = sb.ToString();
                    Console.WriteLine(z);
#endif

                    Assert.Equal(0, valErrors.Count());
                }
            }

            Assert.Equal(templateError, returnedTemplateError);
        }

        [Theory]
        [InlineData("DA024-TrackedRevisions.docx", "DA-Data.xml", true)]
        public void DA102_Throws(string templateDocumentName, string dataFileName, bool templateError)
        {
            FileInfo templateDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, templateDocumentName));
            FileInfo dataFile = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, dataFileName));

            WmlDocument wmlTemplate = new WmlDocument(templateDocx.FullName);
            XElement data = XElement.Load(dataFile.FullName);

            bool returnedTemplateError;
            WmlDocument afterAssembling;
            Assert.Throws<OpenXmlPowerToolsException>(() =>
                {
                    afterAssembling = DocumentAssembler.AssembleDocument(wmlTemplate, data, out returnedTemplateError);
                });
        }

        [Theory]
        [InlineData("DA025-TemplateDocument.docx", "DA-Data.xml", false)]
        public void DA103_UseXmlDocument(string templateDocumentName, string dataFileName, bool templateError)
        {
            FileInfo templateDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, templateDocumentName));
            FileInfo dataFile = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, dataFileName));

            WmlDocument wmlTemplate = new WmlDocument(templateDocx.FullName);
            XmlDocument data = new XmlDocument();
            data.Load(dataFile.FullName);

            bool returnedTemplateError;
            WmlDocument afterAssembling = DocumentAssembler.AssembleDocument(wmlTemplate, data, out returnedTemplateError);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, templateDocx.Name.Replace(".docx", "-processed-by-DocumentAssembler.docx")));
            afterAssembling.SaveAs(assembledDocx.FullName);

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(afterAssembling.DocumentByteArray, 0, afterAssembling.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    OpenXmlValidator v = new OpenXmlValidator();
                    var valErrors = v.Validate(wDoc).Where(ve => !s_ExpectedErrors.Contains(ve.Description));
                    Assert.Equal(0, valErrors.Count());
                }
            }

            Assert.Equal(templateError, returnedTemplateError);
        }

        private static List<string> s_ExpectedErrors = new List<string>()
        {
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowFirstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowLastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowFirstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowLastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddVBand' attribute is not declared.",
        };
    }
}
