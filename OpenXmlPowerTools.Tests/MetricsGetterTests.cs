using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OpenXmlPowerTools;
using Xunit;

#if X64
namespace OpenXmlPowerTools.Tests.X64
#else
namespace OpenXmlPowerTools.Tests
#endif
{
    public class MetricsGetterTests
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
        public void MG001_MetricsGetter(string documentName)
        {
            FileInfo fi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, documentName));

            MetricsGetterSettings settings = new MetricsGetterSettings()
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
                WmlDocument wmlDocument = new WmlDocument(fi.FullName);
                metrics = MetricsGetter.GetDocxMetrics(wmlDocument, settings);
            }
            else if (Util.IsSpreadsheetML(extension))
            {
                SmlDocument smlDocument = new SmlDocument(fi.FullName);
                metrics = MetricsGetter.GetXlsxMetrics(smlDocument, settings);
            }
            else if (Util.IsPresentationML(extension))
            {
                PmlDocument pmlDocument = new PmlDocument(fi.FullName);
                metrics = MetricsGetter.GetPptxMetrics(pmlDocument, settings);
            }

            Assert.NotEqual(null, metrics);
        }
    }
}
