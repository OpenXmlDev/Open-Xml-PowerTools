/***************************************************************************

Copyright (c) Microsoft Corporation 2012-2015.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

***************************************************************************/

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OpenXmlPowerTools;
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
            FileInfo fi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

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
