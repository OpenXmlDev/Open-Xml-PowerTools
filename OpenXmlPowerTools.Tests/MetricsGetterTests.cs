// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Xml.Linq;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OpenXmlPowerTools.Tests
{
    public sealed class MetricsGetterTests
    {
        [Theory]
        [InlineData("DA001-TemplateDocument.docx")]
        [InlineData("Presentation.pptx")]
        [InlineData("Spreadsheet.xlsx")]
        public void GetMetrics_OfficeDocuments_MetricsReturned(string fileName)
        {
            // Arrange
            var testFilesDirectory = new DirectoryInfo("../../../../TestFiles/");
            var file = new FileInfo(Path.Combine(testFilesDirectory.FullName, fileName));

            var settings = new MetricsGetterSettings
            {
                IncludeTextInContentControls = false,
                IncludeXlsxTableCellData = false,
                RetrieveNamespaceList = true,
                RetrieveContentTypeList = true,
            };

            // Act
            XElement metrics = MetricsGetter.GetMetrics(file.FullName, settings);

            // Assert
            Assert.NotNull(metrics);
        }

        [Theory]
        [InlineData("DA001-TemplateDocument.docx")]
        [InlineData("DA002-TemplateDocument.docx")]
        [InlineData("DA003-Select-XPathFindsNoData.docx")]
        [InlineData("DA004-Select-XPathFindsNoDataOptional.docx")]
        [InlineData("DA005-SelectRowData-NoData.docx")]
        [InlineData("DA006-SelectTestValue-NoData.docx")]
        public void GetDocxMetrics_WordDocument_MetricsReturned(string fileName)
        {
            // Arrange
            var testFilesDirectory = new DirectoryInfo("../../../../TestFiles/");
            var file = new FileInfo(Path.Combine(testFilesDirectory.FullName, fileName));
            var wmlDocument = new WmlDocument(file.FullName);

            var settings = new MetricsGetterSettings
            {
                IncludeTextInContentControls = false,
                IncludeXlsxTableCellData = false,
                RetrieveNamespaceList = true,
                RetrieveContentTypeList = true,
            };

            // Act
            XElement metrics = MetricsGetter.GetDocxMetrics(wmlDocument, settings);

            // Assert
            Assert.NotNull(metrics);
        }

        [Theory]
        [InlineData("Presentation.pptx")]
        public void GetPptxMetrics_PowerPointPresentation_MetricsReturned(string fileName)
        {
            // Arrange
            var testFilesDirectory = new DirectoryInfo("../../../../TestFiles/");
            var file = new FileInfo(Path.Combine(testFilesDirectory.FullName, fileName));
            var pmlDocument = new PmlDocument(file.FullName);

            var settings = new MetricsGetterSettings
            {
                IncludeTextInContentControls = false,
                IncludeXlsxTableCellData = false,
                RetrieveNamespaceList = true,
                RetrieveContentTypeList = true,
            };

            // Act
            XElement metrics = MetricsGetter.GetPptxMetrics(pmlDocument, settings);

            // Assert
            Assert.NotNull(metrics);
        }

        [Theory]
        [InlineData("Spreadsheet.xlsx")]
        public void GetXlsxMetrics_ExcelWorkbook_MetricsReturned(string fileName)
        {
            // Arrange
            var testFilesDirectory = new DirectoryInfo("../../../../TestFiles/");
            var file = new FileInfo(Path.Combine(testFilesDirectory.FullName, fileName));
            var smlDocument = new SmlDocument(file.FullName);

            var settings = new MetricsGetterSettings
            {
                IncludeTextInContentControls = false,
                IncludeXlsxTableCellData = false,
                RetrieveNamespaceList = true,
                RetrieveContentTypeList = true,
            };

            // Act
            XElement metrics = MetricsGetter.GetXlsxMetrics(smlDocument, settings);

            // Assert
            Assert.NotNull(metrics);
        }
    }
}

#endif
