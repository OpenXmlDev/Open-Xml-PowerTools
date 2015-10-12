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

using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    /// <summary>
    /// Manages SpreadsheetDocument content
    /// </summary>
    public class SpreadsheetDocumentManager
    {
        private static XNamespace ns;
        private static XNamespace relationshipsns;
        private static int headerRow = 1;

        static SpreadsheetDocumentManager()
        {
            ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            relationshipsns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        }

        /// <summary>
        /// Creates a spreadsheet document from a value table
        /// </summary>
        /// <param name="filePath">Path to store the document</param>
        /// <param name="headerList">Contents of first row (header)</param>
        /// <param name="valueTable">Contents of data</param>
        /// <param name="initialRow">Row to start copying data from</param>
        /// <returns></returns>
        public static void Create(SpreadsheetDocument document, List<string> headerList, string[][] valueTable, int initialRow)
        {
            headerRow = initialRow;

            //Creates a worksheet with given data
            WorksheetPart worksheet = WorksheetAccessor.Create(document, headerList, valueTable, headerRow);
        }

        /// <summary>
        /// Creates a spreadsheet document with a chart from a value table
        /// </summary>
        /// <param name="filePath">Path to store the document</param>
        /// <param name="headerList">Contents of first row (header)</param>
        /// <param name="valueTable">Contents of data</param>
        /// <param name="chartType">Chart type</param>
        /// <param name="categoryColumn">Column to use as category for charting</param>
        /// <param name="columnsToChart">Columns to use as data series</param>
        /// <param name="initialRow">Row index to start copying data</param>
        /// <returns>SpreadsheetDocument</returns>
        //public static void Create(SpreadsheetDocument document, List<string> headerList, string[][] valueTable, ChartType chartType, string categoryColumn, List<string> columnsToChart, int initialRow)
        //{
        //    headerRow = initialRow;

        //    //Creates worksheet with data
        //    WorksheetPart worksheet = WorksheetAccessor.Create(document, headerList, valueTable, headerRow);
        //    //Creates chartsheet with given series and category
        //    string sheetName = GetSheetName(worksheet, document);
        //    ChartsheetPart chartsheet =
        //        ChartsheetAccessor.Create(document,
        //            chartType,
        //            GetValueReferences(sheetName, categoryColumn, headerList, columnsToChart, valueTable),
        //            GetHeaderReferences(sheetName, categoryColumn, headerList, columnsToChart, valueTable),
        //            GetCategoryReference(sheetName, categoryColumn, headerList, valueTable)
        //        );
        //}

        /// <summary>
        /// Gets the internal name of a worksheet from a document
        /// </summary>
        private static string GetSheetName(WorksheetPart worksheet, SpreadsheetDocument document)
        {
            //Gets the id of worksheet part
            string partId = document.WorkbookPart.GetIdOfPart(worksheet);
            XDocument workbookDocument = document.WorkbookPart.GetXDocument();
            //Gets the name from sheet tag related to worksheet
            string sheetName =
                workbookDocument.Root
                .Element(ns + "sheets")
                .Elements(ns + "sheet")
                .Where(
                    t =>
                        t.Attribute(relationshipsns + "id").Value == partId
                ).First()
                .Attribute("name").Value;
            return sheetName;
        }
        /// <summary>
        /// Gets the range reference for category
        /// </summary>
        /// <param name="sheetName">worksheet to take data from</param>
        /// <param name="headerColumn">name of column used as category</param>
        /// <param name="headerList">column names from data</param>
        /// <param name="valueTable">Data values</param>
        /// <returns></returns>
        private static string GetCategoryReference(string sheetName, string headerColumn, List<string> headerList, string[][] valueTable)
        {
            int categoryColumn = headerList.IndexOf(headerColumn.ToUpper()) + 1;
            int numRows = valueTable.GetLength(0);

            return GetRangeReference(
                sheetName,
                categoryColumn,
                headerRow + 1,
                categoryColumn,
                numRows + headerRow
            );
        }

        /// <summary>
        /// Gets a list of range references for each of the series headers
        /// </summary>
        /// <param name="sheetName">worksheet to take data from</param>
        /// <param name="headerColumn">name of column used as category</param>
        /// <param name="headerList">column names from data</param>
        /// <param name="valueTable">Data values</param>
        /// <param name="colsToChart">Columns used as data series</param>
        /// <returns></returns>
        private static List<string> GetHeaderReferences(string sheetName, string headerColumn, List<string> headerList, List<string> colsToChart, string[][] valueTable)
        {
            List<string> valueReferenceList = new List<string>();

            foreach (string column in colsToChart)
            {
                valueReferenceList.Add(
                    GetRangeReference(
                        sheetName,
                        headerList.IndexOf(column.ToUpper()) + 1,
                        headerRow
                    )
                );
            }
            return valueReferenceList;
        }

        /// <summary>
        /// Gets a list of range references for each of the series values
        /// </summary>
        /// <param name="sheetName">worksheet to take data from</param>
        /// <param name="headerColumn">name of column used as category</param>
        /// <param name="headerList">column names from data</param>
        /// <param name="valueTable">Data values</param>
        /// <param name="colsToChart">Columns used as data series</param>
        /// <returns></returns>
        private static List<string> GetValueReferences(string sheetName, string headerColumn, List<string> headerList, List<string> colsToChart, string[][] valueTable)
        {
            List<string> valueReferenceList = new List<string>();
            int numRows = valueTable.GetLength(0);

            foreach (string column in colsToChart)
            {
                int dataColumn = headerList.IndexOf(column.ToUpper()) + 1;
                valueReferenceList.Add(
                    GetRangeReference(
                        sheetName,
                        dataColumn,
                        headerRow + 1,
                        dataColumn,
                        numRows + headerRow
                    )
                );
            }
            return valueReferenceList;
        }

        /// <summary>
        /// Gets a formatted representation of a cell range from a worksheet
        /// </summary>
        private static string GetRangeReference(string worksheet, int column, int row)
        {
            return string.Format("{0}!{1}{2}", worksheet, WorksheetAccessor.GetColumnId(column), row);
        }

        /// <summary>
        /// Gets a formatted representation of a cell range from a worksheet
        /// </summary>
        private static string GetRangeReference(string worksheet, int startColumn, int startRow, int endColumn, int endRow)
        {
            return string.Format("{0}!{1}{2}:{3}{4}",
                worksheet,
                WorksheetAccessor.GetColumnId(startColumn),
                startRow,
                WorksheetAccessor.GetColumnId(endColumn),
                endRow
            );
        }

        /// <summary>
        /// Creates an empty (base) workbook document
        /// </summary>
        /// <returns></returns>
        private static XDocument CreateEmptyWorkbook()
        {
            XDocument document =
                new XDocument(
                    new XElement(ns + "workbook",
                        new XAttribute("xmlns", ns),
                        new XAttribute(XNamespace.Xmlns + "r", relationshipsns),
                        new XElement(ns + "sheets")
                    )
                );

            return document;
        }
    }
}