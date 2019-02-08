// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class ShTests
    {
        // PowerShell oneliner that generates InlineData for all files in a directory
        // dir | % { '[InlineData("' + $_.Name + '")]' } | clip

        [Theory]
        [InlineData("SH101-SimpleFormats.xlsx", "Sheet1")]
        [InlineData("SH102-9-x-9.xlsx", "Sheet1")]
        [InlineData("SH103-No-SharedString.xlsx", "Sheet1")]
        [InlineData("SH104-With-SharedString.xlsx", "Sheet1")]
        [InlineData("SH105-No-SharedString.xlsx", "Sheet1")]
        [InlineData("SH106-9-x-9-Formatted.xlsx", "Sheet1")]
        [InlineData("SH108-SimpleFormattedCell.xlsx", "Sheet1")]
        [InlineData("SH109-CellWithBorder.xlsx", "Sheet1")]
        [InlineData("SH110-CellWithMasterStyle.xlsx", "Sheet1")]
        [InlineData("SH111-ChangedDefaultColumnWidth.xlsx", "Sheet1")]
        [InlineData("SH112-NotVertMergedCell.xlsx", "Sheet1")]
        [InlineData("SH113-VertMergedCell.xlsx", "Sheet1")]
        [InlineData("SH114-Centered-Cell.xlsx", "Sheet1")]
        [InlineData("SH115-DigitsToRight.xlsx", "Sheet1")]
        [InlineData("SH116-FmtNumId-1.xlsx", "Sheet1")]
        [InlineData("SH117-FmtNumId-2.xlsx", "Sheet1")]
        [InlineData("SH118-FmtNumId-3.xlsx", "Sheet1")]
        [InlineData("SH119-FmtNumId-4.xlsx", "Sheet1")]
        [InlineData("SH120-FmtNumId-9.xlsx", "Sheet1")]
        [InlineData("SH121-FmtNumId-11.xlsx", "Sheet1")]
        [InlineData("SH122-FmtNumId-12.xlsx", "Sheet1")]
        [InlineData("SH123-FmtNumId-14.xlsx", "Sheet1")]
        [InlineData("SH124-FmtNumId-15.xlsx", "Sheet1")]
        [InlineData("SH125-FmtNumId-16.xlsx", "Sheet1")]
        [InlineData("SH126-FmtNumId-17.xlsx", "Sheet1")]
        [InlineData("SH127-FmtNumId-18.xlsx", "Sheet1")]
        [InlineData("SH128-FmtNumId-19.xlsx", "Sheet1")]
        [InlineData("SH129-FmtNumId-20.xlsx", "Sheet1")]
        [InlineData("SH130-FmtNumId-21.xlsx", "Sheet1")]
        [InlineData("SH131-FmtNumId-22.xlsx", "Sheet1")]

        public void SH005_ConvertSheet(string name, string sheetName)
        {
            FileInfo sourceXlsx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            var sourceCopiedToDestXlsx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceXlsx.Name.Replace(".xlsx", "-1-Source.xlsx")));
            if (!sourceCopiedToDestXlsx.Exists)
                File.Copy(sourceXlsx.FullName, sourceCopiedToDestXlsx.FullName);

            var dataTemplateFileNameSuffix = "-2-Generated-XmlData-Entire-Sheet.xml";
            var dataXmlFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceXlsx.Name.Replace(".xlsx", dataTemplateFileNameSuffix)));
            using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(sourceXlsx.FullName, true))
            {
                var settings = new SmlToHtmlConverterSettings();
                var rangeXml = SmlDataRetriever.RetrieveSheet(sDoc, sheetName);
                rangeXml.Save(dataXmlFi.FullName);
            }
        }

        [Theory]
        [InlineData("SH101-SimpleFormats.xlsx", "Sheet1", "A1:B10")]
        [InlineData("SH101-SimpleFormats.xlsx", "Sheet1", "A4:B8")]
        [InlineData("SH102-9-x-9.xlsx", "Sheet1", "A1:A1")]
        [InlineData("SH102-9-x-9.xlsx", "Sheet1", "C2:C2")]
        [InlineData("SH102-9-x-9.xlsx", "Sheet1", "A9:A9")]
        [InlineData("SH102-9-x-9.xlsx", "Sheet1", "I1:I1")]
        [InlineData("SH102-9-x-9.xlsx", "Sheet1", "I9:I9")]
        [InlineData("SH102-9-x-9.xlsx", "Sheet1", "A1:I9")]
        [InlineData("SH102-9-x-9.xlsx", "Sheet1", "A2:D4")]
        [InlineData("SH102-9-x-9.xlsx", "Sheet1", "C5:G7")]
        [InlineData("SH103-No-SharedString.xlsx", "Sheet1", "A1:A1")]
        [InlineData("SH104-With-SharedString.xlsx", "Sheet1", "A4:A7")]
        [InlineData("SH105-No-SharedString.xlsx", "Sheet1", "A4:A7")]
        [InlineData("SH106-9-x-9-Formatted.xlsx", "Sheet1", "A1:I9")]
        [InlineData("SH108-SimpleFormattedCell.xlsx", "Sheet1", "A1:A1")]
        [InlineData("SH109-CellWithBorder.xlsx", "Sheet1", "A1:A1")]
        [InlineData("SH110-CellWithMasterStyle.xlsx", "Sheet1", "A1:A1")]
        [InlineData("SH111-ChangedDefaultColumnWidth.xlsx", "Sheet1", "A1:A1")]
        [InlineData("SH112-NotVertMergedCell.xlsx", "Sheet1", "A1:A1")]
        [InlineData("SH113-VertMergedCell.xlsx", "Sheet1", "A1:A1")]
        [InlineData("SH114-Centered-Cell.xlsx", "Sheet1", "A1:A1")]
        [InlineData("SH115-DigitsToRight.xlsx", "Sheet1", "A1:A10")]
        [InlineData("SH116-FmtNumId-1.xlsx", "Sheet1", "A1:A10")]
        [InlineData("SH117-FmtNumId-2.xlsx", "Sheet1", "A1:A10")]
        [InlineData("SH118-FmtNumId-3.xlsx", "Sheet1", "A1:A10")]
        [InlineData("SH119-FmtNumId-4.xlsx", "Sheet1", "A1:A10")]
        [InlineData("SH120-FmtNumId-9.xlsx", "Sheet1", "A1:A10")]
        [InlineData("SH121-FmtNumId-11.xlsx", "Sheet1", "A1:A10")]
        [InlineData("SH122-FmtNumId-12.xlsx", "Sheet1", "A1:A10")]
        [InlineData("SH123-FmtNumId-14.xlsx", "Sheet1", "A1:A10")]
        [InlineData("SH124-FmtNumId-15.xlsx", "Sheet1", "A1:A10")]
        [InlineData("SH125-FmtNumId-16.xlsx", "Sheet1", "A1:A10")]
        [InlineData("SH126-FmtNumId-17.xlsx", "Sheet1", "A1:A10")]
        [InlineData("SH127-FmtNumId-18.xlsx", "Sheet1", "A1:A10")]
        [InlineData("SH128-FmtNumId-19.xlsx", "Sheet1", "A1:A10")]
        [InlineData("SH129-FmtNumId-20.xlsx", "Sheet1", "A1:A10")]
        [InlineData("SH130-FmtNumId-21.xlsx", "Sheet1", "A1:A10")]
        [InlineData("SH131-FmtNumId-22.xlsx", "Sheet1", "A1:A10")]
        
        public void SH004_ConvertRange(string name, string sheetName, string range)
        {
            FileInfo sourceXlsx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            var sourceCopiedToDestXlsx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceXlsx.Name.Replace(".xlsx", "-1-Source.xlsx")));
            if (!sourceCopiedToDestXlsx.Exists)
                File.Copy(sourceXlsx.FullName, sourceCopiedToDestXlsx.FullName);

            var dataTemplateFileNameSuffix = string.Format("-2-Generated-XmlData-{0}.xml", range.Replace(":", ""));
            var dataXmlFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceXlsx.Name.Replace(".xlsx", dataTemplateFileNameSuffix)));
            using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(sourceXlsx.FullName, true))
            {
                var settings = new SmlToHtmlConverterSettings();
                var rangeXml = SmlDataRetriever.RetrieveRange(sDoc, sheetName, range);
                rangeXml.Save(dataXmlFi.FullName);
            }
        }
        

        [Theory]
        [InlineData("SH001-Table.xlsx", "MyTable")]
        [InlineData("SH003-TableWithDateInFirstColumn.xlsx", "MyTable")]
        [InlineData("SH004-TableAtOffsetLocation.xlsx", "MyTable")]
        [InlineData("SH005-Table-With-SharedStrings.xlsx", "Table1")]
        [InlineData("SH006-Table-No-SharedStrings.xlsx", "Table1")]
        [InlineData("SH107-9-x-9-Formatted-Table.xlsx", "Table1")]
        [InlineData("SH007-One-Cell-Table.xlsx", "Table1")]
        [InlineData("SH008-Table-With-Tall-Row.xlsx", "Table1")]
        [InlineData("SH009-Table-With-Wide-Column.xlsx", "Table1")]
        
        public void SH003_ConvertTable(string name, string tableName)
        {
            FileInfo sourceXlsx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            var sourceCopiedToDestXlsx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceXlsx.Name.Replace(".xlsx", "-1-Source.xlsx")));
            if (!sourceCopiedToDestXlsx.Exists)
                File.Copy(sourceXlsx.FullName, sourceCopiedToDestXlsx.FullName);

            var dataXmlFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceXlsx.Name.Replace(".xlsx", "-2-Generated-XmlData.xml")));
            using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(sourceXlsx.FullName, true))
            {
                var settings = new SmlToHtmlConverterSettings();
                var rangeXml = SmlDataRetriever.RetrieveTable(sDoc, tableName);
                rangeXml.Save(dataXmlFi.FullName);
            }
        }

        [Theory]
        [InlineData("Spreadsheet.xlsx", 2)]
        public void SH002_SheetNames(string name, int numberOfSheets)
        {
            FileInfo sourceXlsx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));
            using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(sourceXlsx.FullName, true))
            {
                var sheetNames = SmlDataRetriever.SheetNames(sDoc);
                Assert.Equal(numberOfSheets, sheetNames.Length);
            }
        }

        [Theory]
        [InlineData("SH001-Table.xlsx", 1)]
        [InlineData("SH002-TwoTablesTwoSheets.xlsx", 2)]
        public void SH001_TableNames(string name, int numberOfTables)
        {
            FileInfo sourceXlsx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));
            using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(sourceXlsx.FullName, true))
            {
                var table = SmlDataRetriever.TableNames(sDoc);
                Assert.Equal(numberOfTables, table.Length);
            }
        }
    }
}

#endif
