using OpenXmlPowerTools;
using System.IO;
using System.Xml.Linq;
using Xunit;

namespace OxPt
{
    public class MgTests
    {
        [Theory]
        [InlineData("Presentation.pptx")]
        [InlineData("Spreadsheet.xlsx")]
        [InlineData("DA001-TemplateDocument.docx")]
        [InlineData("DA002-TemplateDocument.docx")]
        [InlineData("DA003-Select-XPathFindsNoData.docx")]
        [InlineData("DA004-Select-XPathFindsNoDataOptional.docx")]
        [InlineData("DA005-SelectRowData-NoData.docx")]
        [InlineData("DA006-SelectTestValue-NoData.docx")]
        public void MG001(string name)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var fi = new FileInfo(Path.Combine(sourceDir.FullName, name));

            var settings = new MetricsGetterSettings()
            {
                IncludeTextInContentControls = false,
                IncludeXlsxTableCellData = false,
                RetrieveNamespaceList = true,
                RetrieveContentTypeList = true,
            };

            var extension = fi.Extension.ToLower();
            XElement metrics = null;
            if (Util.IsWordprocessingML(extension))
            {
                var wmlDocument = new WmlDocument(fi.FullName);
                metrics = MetricsGetter.GetDocxMetrics(wmlDocument, settings);
            }
            else if (Util.IsSpreadsheetML(extension))
            {
                var smlDocument = new SmlDocument(fi.FullName);
                metrics = MetricsGetter.GetXlsxMetrics(smlDocument, settings);
            }
            else if (Util.IsPresentationML(extension))
            {
                var pmlDocument = new PmlDocument(fi.FullName);
                metrics = MetricsGetter.GetPptxMetrics(pmlDocument, settings);
            }

            Assert.NotNull(metrics);
        }
    }
}