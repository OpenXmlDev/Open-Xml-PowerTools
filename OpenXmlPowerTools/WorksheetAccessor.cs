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
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using System.Xml;
using ExcelFormula;

namespace OpenXmlPowerTools
{
    // Classes for "bulk load" of a spreadsheet
    public class MemorySpreadsheet
    {
        private SortedList<int, MemoryRow> rowList;

        public MemorySpreadsheet()
        {
            rowList = new SortedList<int, MemoryRow>();
        }

        public void SetCellValue(int row, int column, object value)
        {
            if (!rowList.ContainsKey(row))
                rowList.Add(row, new MemoryRow(row));
            MemoryRow mr = rowList[row];
            mr.SetCell(new MemoryCell(column, value));
        }

        public void SetCellValue(int row, int column, object value, int styleIndex)
        {
            if (!rowList.ContainsKey(row))
                rowList.Add(row, new MemoryRow(row));
            MemoryRow mr = rowList[row];
            mr.SetCell(new MemoryCell(column, value, styleIndex));
        }

        public object GetCellValue(int row, int column)
        {
            if (!rowList.ContainsKey(row))
                return null;
            MemoryCell cell = rowList[row].GetCell(column);
            if (cell == null)
                return null;
            return cell.GetValue();
        }

        public XElement GetElements()
        {
            XElement root = new XElement(S.sheetData);
            foreach (KeyValuePair<int, MemoryRow> item in rowList)
                root.Add(item.Value.GetElements());
            return root;
        }
    }

    public class MemoryRow
    {
        private int row;
        private SortedList<int, MemoryCell> cellList;

        public MemoryRow(int Row)
        {
            row = Row;
            cellList = new SortedList<int, MemoryCell>();
        }

        public MemoryCell GetCell(int column)
        {
            if (!cellList.ContainsKey(column))
                return null;
            return cellList[column];
        }

        public void SetCell(MemoryCell cell)
        {
            if (cellList.ContainsKey(cell.GetColumn()))
                cellList.Remove(cell.GetColumn());
            cellList.Add(cell.GetColumn(), cell);
        }

        public XElement GetElements()
        {
            XElement root = new XElement(S.row, new XAttribute(NoNamespace.r, row));
            foreach (KeyValuePair<int, MemoryCell> item in cellList)
                root.Add(item.Value.GetElements(row));
            return root;
        }
    }

    public class MemoryCell
    {
        private int column;
        private object cellValue;
        private int styleIndex;

        public MemoryCell(int col, object value)
        {
            column = col;
            cellValue = value;
        }

        public MemoryCell(int col, object value, int style)
        {
            column = col;
            cellValue = value;
            styleIndex = style;
        }

        public int GetColumn()
        {
            return column;
        }
        public object GetValue()
        {
            return cellValue;
        }
        public int GetStyleIndex()
        {
            return styleIndex;
        }

        public XElement GetElements(int row)
        {
            string cellReference = WorksheetAccessor.GetColumnId(column) + row.ToString();
            XElement newCell = null;

            if (cellValue is int || cellValue is double)
                newCell = new XElement(S.c, new XAttribute(NoNamespace.r, cellReference), new XElement(S.v, cellValue.ToString()));
            else if (cellValue is bool)
                newCell = new XElement(S.c, new XAttribute(NoNamespace.r, cellReference), new XAttribute(NoNamespace.t, "b"), new XElement(S.v, (bool)cellValue ? "1" : "0"));
            else if (cellValue is string)
            {
                newCell = new XElement(S.c, new XAttribute(NoNamespace.r, cellReference), new XAttribute(NoNamespace.t, "inlineStr"),
                    new XElement(S._is, new XElement(S.t, cellValue.ToString())));
            }
            if (newCell == null)
                throw new ArgumentException("Invalid cell type.");
            if (styleIndex != 0)
                newCell.Add(new XAttribute(NoNamespace.s, styleIndex));

            return newCell;
        }
    }

    // Static methods to modify worksheets in SpreadsheetML
    public class WorksheetAccessor
    {
        private static XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static XNamespace relationshipsns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        // Finds the WorksheetPart by sheet name
        public static WorksheetPart GetWorksheet(SpreadsheetDocument document, string worksheetName)
        {
            XDocument workbook = document.WorkbookPart.GetXDocument();
            return (WorksheetPart)document.WorkbookPart.GetPartById(
                workbook.Root.Element(S.sheets).Elements(S.sheet).Where(
                    s => s.Attribute(NoNamespace.name).Value.ToLower().Equals(worksheetName.ToLower()))
                .FirstOrDefault().Attribute(R.id).Value);
        }

        // Creates a new worksheet with the specified name
        public static WorksheetPart AddWorksheet(SpreadsheetDocument document, string worksheetName)
        {
            // Create the empty sheet
            WorksheetPart worksheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.PutXDocument(new XDocument(
                new XElement(S.worksheet, new XAttribute("xmlns", S.s), new XAttribute(XNamespace.Xmlns + "r", R.r),
                    new XElement(S.sheetData))));
            XDocument wb = document.WorkbookPart.GetXDocument();

            // Generate a unique sheet ID number
            int sheetId = 1;
            if (wb.Root.Element(S.sheets).Elements(S.sheet).Count() != 0)
                sheetId = wb.Root.Element(S.sheets).Elements(S.sheet).Max(n => Convert.ToInt32(n.Attribute(NoNamespace.sheetId).Value)) + 1;

            // If name is null, generate a name based on the sheet ID
            if (worksheetName == null)
                worksheetName = "Sheet" + sheetId.ToString();

            // Create the new sheet element in the workbook
            wb.Root.Element(S.sheets).Add(new XElement(S.sheet,
                new XAttribute(NoNamespace.name, worksheetName),
                new XAttribute(NoNamespace.sheetId, sheetId),
                new XAttribute(R.id, document.WorkbookPart.GetIdOfPart(worksheetPart))));
            document.WorkbookPart.PutXDocument();
            return worksheetPart;
        }

        // Creates a new worksheet with the specified name and contents from a memory spreadsheet
        public static void SetSheetContents(SpreadsheetDocument document, WorksheetPart worksheet, MemorySpreadsheet contents)
        {
            XDocument worksheetXDocument = worksheet.GetXDocument();
            worksheetXDocument.Root.Element(S.sheetData).ReplaceWith(contents.GetElements());
            worksheet.PutXDocument();
        }

        // Translates the column number to the column reference string (e.g. 1 -> A, 2-> B)
        public static string GetColumnId(int columnNumber)
        {
            string result = "";
            do
            {
                result = ((char)((columnNumber - 1) % 26 + (int)'A')).ToString() + result;
                columnNumber = (columnNumber - 1) / 26;
            } while (columnNumber != 0);
            return result;
        }

        // Gets the value of the specified cell
        // Returned object can be double/Double, int/Int32, bool/Boolean or string/String types
        public static object GetCellValue(SpreadsheetDocument document, WorksheetPart worksheet, int column, int row)
        {
            XDocument worksheetXDocument = worksheet.GetXDocument();
            XElement cellValue = GetCell(worksheetXDocument, column, row);

            if (cellValue != null)
            {
                if (cellValue.Attribute(NoNamespace.t) == null)
                {
                    string value = cellValue.Element(S.v).Value;
                    if (value.Contains("."))
                        return Convert.ToDouble(value);
                    return Convert.ToInt32(value);
                }
                switch (cellValue.Attribute(NoNamespace.t).Value)
                {
                    case "b":
                        return (cellValue.Element(S.v).Value == "1");
                    case "s":
                        return GetSharedString(document, System.Convert.ToInt32(cellValue.Element(S.v).Value));
                    case "inlineStr":
                        return cellValue.Element(S._is).Element(S.t).Value;
                }
            }
            return null;
        }

        // Finds the shared string using its index
        private static string GetSharedString(SpreadsheetDocument document, int index)
        {
            XDocument sharedStringsXDocument = document.WorkbookPart.SharedStringTablePart.GetXDocument();
            return sharedStringsXDocument.Root.Elements().ElementAt<XElement>(index).Value;
        }

        // Gets the cell element (c) for the specified cell
        private static XElement GetCell(XDocument worksheet, int column, int row)
        {
            string cellReference = GetColumnId(column) + row.ToString();
            XElement rowElement = worksheet.Root
                   .Element(S.sheetData)
                   .Elements(S.row)
                   .Where(r => r.Attribute(NoNamespace.r).Value.Equals(row.ToString())).FirstOrDefault<XElement>();
            if (rowElement == null)
                return null;
            return rowElement.Elements(S.c).Where(c => c.Attribute(NoNamespace.r).Value.Equals(cellReference)).FirstOrDefault<XElement>();
        }

        // Sets the value for the specified cell
        // The "value" must be double/Double, int/Int32, bool/Boolean or string/String type
        public static void SetCellValue(SpreadsheetDocument document, WorksheetPart worksheet, int row, int column, object value)
        {
            XDocument worksheetXDocument = worksheet.GetXDocument();
            string cellReference = GetColumnId(column) + row.ToString();
            XElement newCell = null;

            if (value is int || value is double)
                newCell = new XElement(S.c, new XAttribute(NoNamespace.r, cellReference), new XElement(S.v, value.ToString()));
            else if (value is bool)
                newCell = new XElement(S.c, new XAttribute(NoNamespace.r, cellReference), new XAttribute(NoNamespace.t, "b"), new XElement(S.v, (bool)value ? "1" : "0"));
            else if (value is string)
            {
                newCell = new XElement(S.c, new XAttribute(NoNamespace.r, cellReference), new XAttribute(NoNamespace.t, "inlineStr"),
                    new XElement(S._is, new XElement(S.t, value.ToString())));
            }
            if (newCell == null)
                throw new ArgumentException("Invalid cell type.");

            SetCell(worksheetXDocument, newCell);
        }

        // Sets the specified cell
        private static void SetCell(XDocument worksheetXDocument, XElement newCell)
        {
            int row;
            int column;
            string cellReference = newCell.Attribute(NoNamespace.r).Value;
            GetRowColumn(cellReference, out row, out column);

            // Find the row containing the cell to add the value to
            XElement rowElement = worksheetXDocument.Root
                .Element(S.sheetData)
                .Elements(S.row)
                .Where(t => t.Attribute(NoNamespace.r).Value == row.ToString())
                .FirstOrDefault();

            if (rowElement == null)
            {
                //row element does not exist
                //create a new one
                rowElement = CreateEmptyRow(row);

                //row elements must appear in order inside sheetData element
                if (worksheetXDocument.Root.Element(S.sheetData).HasElements)
                {   //if there are more rows already defined at sheetData element
                    //find the row with the inmediate higher index for the row containing the cell to set the value to
                    XElement rowAfterElement = FindRowAfter(worksheetXDocument, row);
                    //if there is a row with an inmediate higher index already defined at sheetData
                    if (rowAfterElement != null)
                    {
                        //add the new row before the row with an inmediate higher index
                        rowAfterElement.AddBeforeSelf(rowElement);
                    }
                    else
                    {   //this row is going to be the one with the highest index (add it as the last element for sheetData)
                        worksheetXDocument.Root.Element(S.sheetData).Elements(S.row).Last().AddAfterSelf(rowElement);
                    }
                }
                else
                {   //there are no other rows already defined at sheetData
                    //Add a new row elemento to sheetData
                    worksheetXDocument.Root.Element(S.sheetData).Add(rowElement);
                }

                //Add the new cell to the row Element
                rowElement.Add(newCell);
            }
            else
            {
                //row containing the cell to set the value to is already defined at sheetData
                //look if cell already exist at that row
                XElement currentCell = rowElement
                    .Elements(S.c)
                    .Where(t => t.Attribute(NoNamespace.r).Value == cellReference)
                    .FirstOrDefault();

                if (currentCell == null)
                {   //cell element does not exist at row indicated as parameter
                    //find the inmediate right column for the cell to set the value to
                    XElement columnAfterXElement = FindColumAfter(worksheetXDocument, row, column);
                    if (columnAfterXElement != null)
                    {
                        //Insert the new cell before the inmediate right column
                        columnAfterXElement.AddBeforeSelf(newCell);
                    }
                    else
                    {   //There is no inmediate right cell 
                        //Add the new cell as the last element for the row
                        rowElement.Add(newCell);
                    }
                }
                else
                {
                    //cell alreay exist
                    //replace the current cell with that with the new value
                    currentCell.ReplaceWith(newCell);
                }
            }
        }

        // Finds the row element (r) with a higher number than the specified "row" number
        private static XElement FindRowAfter(XDocument worksheet, int row)
        {
            return worksheet.Root
                .Element(S.sheetData)
                .Elements(S.row)
                .FirstOrDefault(r => System.Convert.ToInt32(r.Attribute(NoNamespace.r).Value) > row);
        }

        // Finds the cell element (c) in the specified row that is after the specified "column" number
        private static XElement FindColumAfter(XDocument worksheet, int row, int column)
        {
            return worksheet.Root
                .Element(S.sheetData)
                .Elements(S.row)
                .FirstOrDefault(r => System.Convert.ToInt32(r.Attribute(NoNamespace.r).Value) == row)
                .Elements(S.c)
                .FirstOrDefault(c => GetColumnNumber(c.Attribute(NoNamespace.r).Value) > GetColumnNumber(GetColumnId(column) + row));
        }

        // Converts the column reference string to a column number (e.g. A -> 1, B -> 2)
        public static int GetColumnNumber(string cellReference)
        {
            int columnNumber = 0;
            foreach (char c in cellReference)
            {
                if (Char.IsLetter(c))
                    columnNumber = columnNumber * 26 + System.Convert.ToInt32(c) - System.Convert.ToInt32('A') + 1;
            }
            return columnNumber;
        }

        // Converts a cell reference string into the row and column numbers for that cell
        // e.g. G5 -> [row = 5, column = 7]
        private static void GetRowColumn(string cellReference, out int row, out int column)
        {
            row = 0;
            column = 0;
            foreach (char c in cellReference)
            {
                if (Char.IsLetter(c))
                    column = column * 26 + System.Convert.ToInt32(c) - System.Convert.ToInt32('A') + 1;
                else
                    row = row * 10 + System.Convert.ToInt32(c) - System.Convert.ToInt32('0');
            }
        }

        // Returns the row and column numbers and worksheet part for the named range
        public static WorksheetPart GetRange(SpreadsheetDocument doc, string rangeName, out int startRow, out int startColumn, out int endRow, out int endColumn)
        {
            XDocument book = doc.WorkbookPart.GetXDocument();
            if (book.Root.Element(S.definedNames) == null)
                throw new ArgumentException("Range name not found: " + rangeName);
            XElement element = book.Root.Element(S.definedNames).Elements(S.definedName)
                .Where(t => t.Attribute(NoNamespace.name).Value == rangeName).FirstOrDefault();
            if (element == null)
                throw new ArgumentException("Range name not found: " + rangeName);
            string sheetName = element.Value.Substring(0, element.Value.IndexOf('!'));
            string range = element.Value.Substring(element.Value.IndexOf('!') + 1).Replace("$","");
            int colonIndex = range.IndexOf(':');
            GetRowColumn(range.Substring(0, colonIndex), out startRow, out startColumn);
            GetRowColumn(range.Substring(colonIndex + 1), out endRow, out endColumn);
            return GetWorksheet(doc, sheetName);
        }

        // Sets the named range with the specified range of row and column numbers
        public static void SetRange(SpreadsheetDocument doc, string rangeName, string sheetName, int startRow, int startColumn, int endRow, int endColumn)
        {
            XDocument book = doc.WorkbookPart.GetXDocument();
            if (book.Root.Element(S.definedNames) == null)
                book.Root.Add(new XElement(S.definedNames));
            XElement element = book.Root.Element(S.definedNames).Elements(S.definedName)
                .Where(t => t.Attribute(NoNamespace.name).Value == rangeName).FirstOrDefault();
            if (element == null)
            {
                element = new XElement(S.definedName, new XAttribute(NoNamespace.name, rangeName));
                book.Root.Element(S.definedNames).Add(element);
            }
            element.SetValue(String.Format("{0}!${1}${2}:${3}${4}", sheetName, GetColumnId(startColumn), startRow, GetColumnId(endColumn), endRow));
            doc.WorkbookPart.PutXDocument();
        }

        // Sets the end row for the named range
        public static void UpdateRangeEndRow(SpreadsheetDocument doc, string rangeName, int lastRow)
        {
            // Update named range used by pivot table
            XDocument book = doc.WorkbookPart.GetXDocument();
            XElement element = book.Root.Element(S.definedNames).Elements(S.definedName)
                .Where(t => t.Attribute(NoNamespace.name).Value == rangeName).FirstOrDefault();
            if (element != null)
            {
                string original = element.Value;
                element.SetValue(original.Substring(0, original.Length - 1) + lastRow.ToString());
            }
            doc.WorkbookPart.PutXDocument();
        }

        // Creates an empty row element (r) with the specified row number
        private static XElement CreateEmptyRow(int row)
        {
            return new XElement(S.row, new XAttribute(NoNamespace.r, row.ToString()));
        }

        public static void ForceCalculateOnLoad(SpreadsheetDocument document)
        {
            XDocument book = document.WorkbookPart.GetXDocument();
            XElement element = book.Root.Element(S.calcPr);
            if (element == null)
            {
                book.Root.Add(new XElement(S.calcPr));
            }
            element.SetAttributeValue(NoNamespace.fullCalcOnLoad, "1");
            document.WorkbookPart.PutXDocument();
        }

        public static void FormulaReplaceSheetName(SpreadsheetDocument document, string oldName, string newName)
        {
            foreach (WorksheetPart sheetPart in document.WorkbookPart.WorksheetParts)
            {
                XDocument sheetDoc = sheetPart.GetXDocument();
                bool changed = false;
                foreach (XElement formula in sheetDoc.Descendants(S.f))
                {
                    ParseFormula parser = new ParseFormula(formula.Value);
                    string newFormula = parser.ReplaceSheetName(oldName, newName);
                    if (newFormula != formula.Value)
                    {
                        formula.SetValue(newFormula);
                        changed = true;
                    }
                }
                if (changed)
                {
                    sheetPart.PutXDocument();
                    ForceCalculateOnLoad(document);
                }
            }
        }

        // Copy all cells in the specified range to a new location
        public static void CopyCellRange(SpreadsheetDocument document, WorksheetPart worksheet, int startRow, int startColumn, int endRow, int endColumn,
            int toRow, int toColumn)
        {
            int rowOffset = toRow - startRow;
            int columnOffset = toColumn - startColumn;
            XDocument worksheetXDocument = worksheet.GetXDocument();
            for (int row = startRow; row <= endRow; row++)
                for (int column = startColumn; column <= endColumn; column++)
                {
                    XElement oldCell = GetCell(worksheetXDocument, column, row);
                    if (oldCell != null)
                    {
                        XElement newCell = new XElement(oldCell);
                        newCell.SetAttributeValue(NoNamespace.r, GetColumnId(column + columnOffset) + (row + rowOffset).ToString());
                        XElement formula = newCell.Element(S.f);
                        if (formula != null)
                        {
                            ParseFormula parser = new ParseFormula(formula.Value);
                            formula.SetValue(parser.ReplaceRelativeCell(rowOffset, columnOffset));
                        }
                        SetCell(worksheetXDocument, newCell);
                    }
                }
            worksheet.PutXDocument();
            ForceCalculateOnLoad(document);
        }

        // Creates a pivot table in the specified sheet using the specified range name
        // The new pivot table will not be configured with any fields in the rows, columns, filters or values
        public static PivotTablePart CreatePivotTable(SpreadsheetDocument document, string rangeName, WorksheetPart sheet)
        {
            int startRow, startColumn, endRow, endColumn;
            WorksheetPart sourceSheet = GetRange(document, rangeName, out startRow, out startColumn, out endRow, out endColumn);

            // Fill out pivotFields element (for PivotTablePart) and cacheFields element (for PivotTableCacheDefinitionPart)
            // with an element for each column in the source range
            XElement pivotFields = new XElement(S.pivotFields, new XAttribute(NoNamespace.count, (endColumn - startColumn + 1).ToString()));
            XElement cacheFields = new XElement(S.cacheFields, new XAttribute(NoNamespace.count, (endColumn - startColumn + 1).ToString()));
            for (int column = startColumn; column <= endColumn; column++)
            {
                pivotFields.Add(new XElement(S.pivotField, new XAttribute(NoNamespace.showAll, "0")));
                XElement sharedItems = new XElement(S.sharedItems);
                // Determine numeric sharedItems values, if any
                object value = GetCellValue(document, sourceSheet, column, startRow + 1);
                if (value is double || value is Int32)
                {
                    bool hasDouble = false;
                    double minValue = Convert.ToDouble(value);
                    double maxValue = Convert.ToDouble(value);
                    if (value is double)
                        hasDouble = true;
                    for (int row = startRow + 1; row <= endRow; row++)
                    {
                        value = GetCellValue(document, sourceSheet, column, row);
                        if (value is double)
                            hasDouble = true;
                        if (Convert.ToDouble(value) < minValue)
                            minValue = Convert.ToDouble(value);
                        if (Convert.ToDouble(value) > maxValue)
                            maxValue = Convert.ToDouble(value);
                    }
                    sharedItems.Add(new XAttribute(NoNamespace.containsSemiMixedTypes, "0"),
                        new XAttribute(NoNamespace.containsString, "0"), new XAttribute(NoNamespace.containsNumber, "1"),
                        new XAttribute(NoNamespace.minValue, minValue.ToString()), new XAttribute(NoNamespace.maxValue, maxValue.ToString()));
                    if (!hasDouble)
                        sharedItems.Add(new XAttribute(NoNamespace.containsInteger, "1"));
                }
                cacheFields.Add(new XElement(S.cacheField, new XAttribute(NoNamespace.name, GetCellValue(document, sourceSheet, column, startRow).ToString()),
                    new XAttribute(NoNamespace.numFmtId, "0"), sharedItems));
            }

            // Fill out pivotCacheRecords element (for PivotTableCacheRecordsPart) with an element
            // for each row in the source range
            XElement pivotCacheRecords = new XElement(S.pivotCacheRecords, new XAttribute("xmlns", S.s),
                new XAttribute(XNamespace.Xmlns + "r", R.r), new XAttribute(NoNamespace.count, (endRow - startRow).ToString()));
            for (int row = startRow + 1; row <= endRow; row++)
            {
                XElement r = new XElement(S.r);

                // Fill the record element with a value from each column in the source row
                for (int column = startColumn; column <= endColumn; column++)
                {
                    object value = GetCellValue(document, sourceSheet, column, row);
                    if (value is String)
                        r.Add(new XElement(S._s, new XAttribute(NoNamespace.v, value.ToString())));
                    else
                        r.Add(new XElement(S.n, new XAttribute(NoNamespace.v, value.ToString())));
                }
                pivotCacheRecords.Add(r);
            }

            // Create pivot table parts with proper links
            PivotTablePart pivotTable = sheet.AddNewPart<PivotTablePart>();
            PivotTableCacheDefinitionPart cacheDef = pivotTable.AddNewPart<PivotTableCacheDefinitionPart>();
            PivotTableCacheRecordsPart records = cacheDef.AddNewPart<PivotTableCacheRecordsPart>();
            document.WorkbookPart.AddPart<PivotTableCacheDefinitionPart>(cacheDef);

            // Set content for the PivotTableCacheRecordsPart and PivotTableCacheDefinitionPart
            records.PutXDocument(new XDocument(pivotCacheRecords));
            cacheDef.PutXDocument(new XDocument(new XElement(S.pivotCacheDefinition, new XAttribute("xmlns", S.s),
                new XAttribute(XNamespace.Xmlns + "r", R.r), new XAttribute(R.id, cacheDef.GetIdOfPart(records)),
                new XAttribute(NoNamespace.recordCount, (endRow - startRow).ToString()),
                    new XElement(S.cacheSource, new XAttribute(NoNamespace.type, "worksheet"),
                        new XElement(S.worksheetSource, new XAttribute(NoNamespace.name, rangeName))),
                    cacheFields)));

            // Create the pivotCache entry in the workbook part
            int cacheId = 1;
            XDocument wb = document.WorkbookPart.GetXDocument();
            if (wb.Root.Element(S.pivotCaches) == null)
                wb.Root.Add(new XElement(S.pivotCaches));
            else
            {
                if (wb.Root.Element(S.pivotCaches).Elements(S.pivotCache).Count() != 0)
                    cacheId = wb.Root.Element(S.pivotCaches).Elements(S.pivotCache).Max(n => Convert.ToInt32(n.Attribute(NoNamespace.cacheId).Value)) + 1;
            }
            wb.Root.Element(S.pivotCaches).Add(new XElement(S.pivotCache,
                new XAttribute(NoNamespace.cacheId, cacheId),
                new XAttribute(R.id, document.WorkbookPart.GetIdOfPart(cacheDef))));
            document.WorkbookPart.PutXDocument();

            // Set the content for the PivotTablePart
            pivotTable.PutXDocument(new XDocument(new XElement(S.pivotTableDefinition, new XAttribute("xmlns", S.s),
                new XAttribute(NoNamespace.name, "PivotTable1"), new XAttribute(NoNamespace.cacheId, cacheId.ToString()),
                new XAttribute(NoNamespace.dataCaption, "Values"),
                new XElement(S.location, new XAttribute(NoNamespace._ref, "A3:C20"),
                    new XAttribute(NoNamespace.firstHeaderRow, "1"), new XAttribute(NoNamespace.firstDataRow, "1"),
                    new XAttribute(NoNamespace.firstDataCol, "0")), pivotFields)));

            return pivotTable;
        }

        public enum PivotAxis { Row, Column, Page };
        public static void AddPivotAxis(SpreadsheetDocument document, WorksheetPart sheet, string fieldName, PivotAxis axis)
        {
            // Create indexed items in cache and definition
            PivotTablePart pivotTablePart = sheet.GetPartsOfType<PivotTablePart>().First();
            PivotTableCacheDefinitionPart cacheDefPart = pivotTablePart.GetPartsOfType<PivotTableCacheDefinitionPart>().First();
            PivotTableCacheRecordsPart recordsPart = cacheDefPart.GetPartsOfType<PivotTableCacheRecordsPart>().First();
            XDocument cacheDef = cacheDefPart.GetXDocument();
            int index = Array.FindIndex(cacheDef.Descendants(S.cacheField).ToArray(),
                z => z.Attribute(NoNamespace.name).Value == fieldName);
            XDocument records = recordsPart.GetXDocument();
            List<XElement> values = new List<XElement>();
            foreach (XElement rec in records.Descendants(S.r))
            {
                XElement val = rec.Elements().Skip(index).First();
                int x = Array.FindIndex(values.ToArray(), z => XElement.DeepEquals(z, val));
                if (x == -1)
                {
                    values.Add(val);
                    x = values.Count() - 1;
                }
                val.ReplaceWith(new XElement(S.x, new XAttribute(NoNamespace.v, x)));
            }
            XElement sharedItems = cacheDef.Descendants(S.cacheField).Skip(index).First().Element(S.sharedItems);
            sharedItems.Add(new XAttribute(NoNamespace.count, values.Count()), values);
            recordsPart.PutXDocument();
            cacheDefPart.PutXDocument();

            // Add axis definition to pivot table field
            XDocument pivotTable = pivotTablePart.GetXDocument();
            XElement pivotField = pivotTable.Descendants(S.pivotField).Skip(index).First();
            XElement items = new XElement(S.items, new XAttribute(NoNamespace.count, values.Count() + 1),
                values.OrderBy(z => z.Attribute(NoNamespace.v).Value).Select(z => new XElement(S.item,
                    new XAttribute(NoNamespace.x, Array.FindIndex(values.ToArray(),
                        a => a.Attribute(NoNamespace.v).Value == z.Attribute(NoNamespace.v).Value)))));
            items.Add(new XElement(S.item, new XAttribute(NoNamespace.t, "default")));
            switch (axis)
            {
                case PivotAxis.Column:
                    pivotField.Add(new XAttribute(NoNamespace.axis, "axisCol"), items);
                    // Add to colFields
                    {
                        XElement fields = pivotTable.Element(S.pivotTableDefinition).Element(S.colFields);
                        if (fields == null)
                        {
                            fields = new XElement(S.colFields, new XAttribute(NoNamespace.count, 0));
                            XElement rowFields = pivotTable.Element(S.pivotTableDefinition).Element(S.rowFields);
                            if (rowFields == null)
                                pivotTable.Element(S.pivotTableDefinition).Element(S.pivotFields).AddAfterSelf(fields);
                            else
                                rowFields.AddAfterSelf(fields);
                        }
                        fields.Add(new XElement(S.field, new XAttribute(NoNamespace.x, index)));
                        fields.Attribute(NoNamespace.count).Value = fields.Elements(S.field).Count().ToString();
                    }
                    break;
                case PivotAxis.Row:
                    pivotField.Add(new XAttribute(NoNamespace.axis, "axisRow"), items);
                    // Add to rowFields
                    {
                        XElement fields = pivotTable.Element(S.pivotTableDefinition).Element(S.rowFields);
                        if (fields == null)
                        {
                            fields = new XElement(S.rowFields, new XAttribute(NoNamespace.count, 0));
                            pivotTable.Element(S.pivotTableDefinition).Element(S.pivotFields).AddAfterSelf(fields);
                        }
                        fields.Add(new XElement(S.field, new XAttribute(NoNamespace.x, index)));
                        fields.Attribute(NoNamespace.count).Value = fields.Elements(S.field).Count().ToString();
                    }
                    break;
                case PivotAxis.Page:
                    pivotField.Add(new XAttribute(NoNamespace.axis, "axisPage"), items);
                    // Add to pageFields
                    {
                        XElement fields = pivotTable.Element(S.pivotTableDefinition).Element(S.pageFields);
                        if (fields == null)
                        {
                            fields = new XElement(S.pageFields, new XAttribute(NoNamespace.count, 0));
                            XElement prev = pivotTable.Element(S.pivotTableDefinition).Element(S.colFields);
                            if (prev == null)
                                prev = pivotTable.Element(S.pivotTableDefinition).Element(S.rowFields);
                            if (prev == null)
                                pivotTable.Element(S.pivotTableDefinition).Element(S.pivotFields).AddAfterSelf(fields);
                            else
                                prev.AddAfterSelf(fields);
                        }
                        fields.Add(new XElement(S.pageField, new XAttribute(NoNamespace.fld, index)));
                        fields.Attribute(NoNamespace.count).Value = fields.Elements(S.field).Count().ToString();
                    }
                    break;
            }
            pivotTablePart.PutXDocument();
            ForcePivotRefresh(cacheDefPart);
        }

        public static void AddDataValueLabel(SpreadsheetDocument document, WorksheetPart sheet, PivotAxis axis)
        {
            PivotTablePart pivotTablePart = sheet.GetPartsOfType<PivotTablePart>().First();
            XDocument pivotTable = pivotTablePart.GetXDocument();
            switch (axis)
            {
                case PivotAxis.Column:
                    // Add to colFields
                    {
                        XElement fields = pivotTable.Element(S.pivotTableDefinition).Element(S.colFields);
                        if (fields == null)
                        {
                            fields = new XElement(S.colFields, new XAttribute(NoNamespace.count, 0));
                            XElement rowFields = pivotTable.Element(S.pivotTableDefinition).Element(S.rowFields);
                            if (rowFields == null)
                                pivotTable.Element(S.pivotTableDefinition).Element(S.pivotFields).AddAfterSelf(fields);
                            else
                                rowFields.AddAfterSelf(fields);
                        }
                        fields.Add(new XElement(S.field, new XAttribute(NoNamespace.x, -2)));
                        fields.Attribute(NoNamespace.count).Value = fields.Elements(S.field).Count().ToString();
                    }
                    break;
                case PivotAxis.Row:
                    // Add to rowFields
                    {
                        XElement fields = pivotTable.Element(S.pivotTableDefinition).Element(S.rowFields);
                        if (fields == null)
                        {
                            fields = new XElement(S.rowFields, new XAttribute(NoNamespace.count, 0));
                            pivotTable.Element(S.pivotTableDefinition).Element(S.pivotFields).AddAfterSelf(fields);
                        }
                        fields.Add(new XElement(S.field, new XAttribute(NoNamespace.x, -2)));
                        fields.Attribute(NoNamespace.count).Value = fields.Elements(S.field).Count().ToString();
                    }
                    break;
                case PivotAxis.Page:
                    // Add to pageFields
                    {
                        XElement fields = pivotTable.Element(S.pivotTableDefinition).Element(S.pageFields);
                        if (fields == null)
                        {
                            fields = new XElement(S.pageFields, new XAttribute(NoNamespace.count, 0));
                            XElement prev = pivotTable.Element(S.pivotTableDefinition).Element(S.colFields);
                            if (prev == null)
                                prev = pivotTable.Element(S.pivotTableDefinition).Element(S.rowFields);
                            if (prev == null)
                                pivotTable.Element(S.pivotTableDefinition).Element(S.pivotFields).AddAfterSelf(fields);
                            else
                                prev.AddAfterSelf(fields);
                        }
                        fields.Add(new XElement(S.pageField, new XAttribute(NoNamespace.fld, -2)));
                        fields.Attribute(NoNamespace.count).Value = fields.Elements(S.field).Count().ToString();
                    }
                    break;
            }
            pivotTablePart.PutXDocument();
            PivotTableCacheDefinitionPart cacheDefPart = pivotTablePart.GetPartsOfType<PivotTableCacheDefinitionPart>().First();
            ForcePivotRefresh(cacheDefPart);
        }

        public static void AddDataValue(SpreadsheetDocument document, WorksheetPart sheet, string fieldName)
        {
            PivotTablePart pivotTablePart = sheet.GetPartsOfType<PivotTablePart>().First();
            PivotTableCacheDefinitionPart cacheDefPart = pivotTablePart.GetPartsOfType<PivotTableCacheDefinitionPart>().First();
            XDocument cacheDef = cacheDefPart.GetXDocument();
            int index = Array.FindIndex(cacheDef.Descendants(S.cacheField).ToArray(),
                z => z.Attribute(NoNamespace.name).Value == fieldName);
            XDocument pivotTable = pivotTablePart.GetXDocument();
            XElement pivotField = pivotTable.Descendants(S.pivotField).Skip(index).First();
            pivotField.Add(new XAttribute(NoNamespace.dataField, "1"));
            XElement fields = pivotTable.Element(S.pivotTableDefinition).Element(S.dataFields);
            if (fields == null)
            {
                fields = new XElement(S.dataFields, new XAttribute(NoNamespace.count, 0));
                XElement prev = pivotTable.Element(S.pivotTableDefinition).Element(S.pageFields);
                if (prev == null)
                    prev = pivotTable.Element(S.pivotTableDefinition).Element(S.colFields);
                if (prev == null)
                    prev = pivotTable.Element(S.pivotTableDefinition).Element(S.rowFields);
                if (prev == null)
                    prev = pivotTable.Element(S.pivotTableDefinition).Element(S.pivotFields);
                prev.AddAfterSelf(fields);
            }
            fields.Add(new XElement(S.dataField, new XAttribute(NoNamespace.name, "Sum of " + fieldName),
                        new XAttribute(NoNamespace.fld, index), new XAttribute(NoNamespace.baseField, 0),
                        new XAttribute(NoNamespace.baseItem, 0)));
            int count = fields.Elements(S.dataField).Count();
            fields.Attribute(NoNamespace.count).Value = count.ToString();
            if (count == 2)
            {   // Only when data field count goes from 1 to 2 do we add a special column to label the data fields
                AddDataValueLabel(document, sheet, PivotAxis.Column);
            }
            pivotTablePart.PutXDocument();
            ForcePivotRefresh(cacheDefPart);
        }

        private static void ForcePivotRefresh(PivotTableCacheDefinitionPart cacheDef)
        {
            XDocument doc = cacheDef.GetXDocument();
            XElement def = doc.Element(S.pivotCacheDefinition);
            if (def.Attribute(NoNamespace.refreshOnLoad) == null)
                def.Add(new XAttribute(NoNamespace.refreshOnLoad, 1));
            else
                def.Attribute(NoNamespace.refreshOnLoad).Value = "1";
            cacheDef.PutXDocument();
        }

        public static void CheckNumberFormat(SpreadsheetDocument document, int fmtID, string formatCode)
        {
            XElement numFmt = new XElement(S.numFmt, new XAttribute(NoNamespace.numFmtId, fmtID.ToString()),
                new XAttribute(NoNamespace.formatCode, formatCode));
            XDocument styles = document.WorkbookPart.WorkbookStylesPart.GetXDocument();
            XElement numFmts = styles.Root.Element(S.numFmts);
            if (numFmts == null)
            {
                styles.Root.Element(S.fonts).AddBeforeSelf(new XElement(S.numFmts, new XAttribute(NoNamespace.count, "0")));
                numFmts = styles.Root.Element(S.numFmts);
            }
            int index = Array.FindIndex(numFmts.Elements(S.numFmt).ToArray(),
                z => XElement.DeepEquals(z, numFmt));
            if (index == -1)
            {
                numFmts.Add(numFmt);
                numFmts.Attribute(NoNamespace.count).Value = numFmts.Elements(S.numFmt).Count().ToString();
                document.WorkbookPart.WorkbookStylesPart.PutXDocument();
            }
        }

        public class ColorInfo
        {
            public enum ColorType { Theme, Indexed };

            private bool Auto;
            private string RGB;
            private int Indexed;
            private int Theme;
            private double Tint;

            public ColorInfo()
            {
                Auto = true;
            }
            public ColorInfo(ColorType type, int value)
            {
                if (type == ColorType.Indexed)
                    Indexed = value;
                else if (type == ColorType.Theme)
                    Theme = value;
            }
            public ColorInfo(int theme, double tint)
            {
                Theme = theme;
                Tint = tint;
            }
            public ColorInfo(string rgb)
            {
                RGB = rgb;
            }

            public XElement GetXElement(XName colorName)
            {
                XElement color = new XElement(colorName);
                if (Auto)
                    color.Add(new XAttribute(NoNamespace.auto, "1"));
                else if (RGB != null)
                    color.Add(new XAttribute(NoNamespace.rgb, RGB));
                else if (Indexed != 0)
                    color.Add(new XAttribute(NoNamespace.indexed, Indexed));
                else
                    color.Add(new XAttribute(NoNamespace.theme, Theme));
                if (Tint != 0)
                    color.Add(new XAttribute(NoNamespace.tint, Tint));
                return color;
            }
        }

        public class Font
        {
            public enum SchemeType { None, Major, Minor };

            public bool Bold { get; set; }
            public ColorInfo Color { get; set; }
            public bool Condense { get; set; }
            public bool Extend { get; set; }
            public int Family { get; set; }
            public bool Italic { get; set; }
            public string Name { get; set; }
            public bool Outline { get; set; }
            public SchemeType Scheme { get; set; }
            public bool Shadow { get; set; }
            public bool StrikeThrough { get; set; }
            public int Size { get; set; }
            public bool Underline { get; set; }

            public XElement GetXElement()
            {
                XElement font = new XElement(S.font);
                if (Bold)
                    font.Add(new XElement(S.b));
                if (Italic)
                    font.Add(new XElement(S.i));
                if (Underline)
                    font.Add(new XElement(S.u));
                if (StrikeThrough)
                    font.Add(new XElement(S.strike));
                if (Condense)
                    font.Add(new XElement(S.condense));
                if (Extend)
                    font.Add(new XElement(S.extend));
                if (Outline)
                    font.Add(new XElement(S.outline));
                if (Shadow)
                    font.Add(new XElement(S.shadow));
                if (Size != 0)
                    font.Add(new XElement(S.sz, new XAttribute(NoNamespace.val, Size.ToString())));
                if (Color != null)
                    font.Add(Color.GetXElement(S.color));
                if (Name != null)
                    font.Add(new XElement(S.name, new XAttribute(NoNamespace.val, Name)));
                if (Family != 0)
                    font.Add(new XElement(S.family, new XAttribute(NoNamespace.val, Family.ToString())));
                switch (Scheme)
                {
                    case SchemeType.Major:
                        font.Add(new XElement(S.scheme, new XAttribute(NoNamespace.val, "major")));
                        break;
                    case SchemeType.Minor:
                        font.Add(new XElement(S.scheme, new XAttribute(NoNamespace.val, "minor")));
                        break;
                }
                return font;
            }
        }

        public static int GetFontIndex(SpreadsheetDocument document, Font f)
        {
            XElement font = f.GetXElement();
            XDocument styles = document.WorkbookPart.WorkbookStylesPart.GetXDocument();
            XElement fonts = styles.Root.Element(S.fonts);
            int index = Array.FindIndex(fonts.Elements(S.font).ToArray(),
                z => XElement.DeepEquals(z, font));
            if (index != -1)
                return index;
            fonts.Add(font);
            fonts.Attribute(NoNamespace.count).Value = fonts.Elements(S.font).Count().ToString();
            document.WorkbookPart.WorkbookStylesPart.PutXDocument();
            return fonts.Elements(S.font).Count() - 1;
        }

        public class PatternFill
        {
            public enum PatternType
            {
                None, Solid, DarkDown, DarkGray, DarkGrid, DarkHorizontal, DarkTrellis, DarkUp, DarkVertical,
                Gray0625, Gray125, LightDown, LightGray, LightGrid, LightHorizontal, LightTrellis, LightUp, LightVertical, MediumGray
            };

            private PatternType Pattern;
            private ColorInfo BgColor;
            private ColorInfo FgColor;

            public PatternFill(PatternType pattern, ColorInfo bgColor, ColorInfo fgColor)
            {
                Pattern = pattern;
                BgColor = bgColor;
                FgColor = fgColor;
            }

            public XElement GetXElement()
            {
                XElement pattern = new XElement(S.patternFill);
                switch (Pattern)
                {
                    case PatternType.DarkDown:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "darkDown"));
                        break;
                    case PatternType.DarkGray:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "darkGray"));
                        break;
                    case PatternType.DarkGrid:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "darkGrid"));
                        break;
                    case PatternType.DarkHorizontal:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "darkHorizontal"));
                        break;
                    case PatternType.DarkTrellis:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "darkTrellis"));
                        break;
                    case PatternType.DarkUp:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "darkUp"));
                        break;
                    case PatternType.DarkVertical:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "darkVertical"));
                        break;
                    case PatternType.Gray0625:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "gray0625"));
                        break;
                    case PatternType.Gray125:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "gray125"));
                        break;
                    case PatternType.LightDown:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "lightDown"));
                        break;
                    case PatternType.LightGray:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "lightGray"));
                        break;
                    case PatternType.LightGrid:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "lightGrid"));
                        break;
                    case PatternType.LightHorizontal:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "lightHorizontal"));
                        break;
                    case PatternType.LightTrellis:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "lightTrellis"));
                        break;
                    case PatternType.LightUp:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "lightUp"));
                        break;
                    case PatternType.LightVertical:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "lightVertical"));
                        break;
                    case PatternType.MediumGray:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "mediumGray"));
                        break;
                    case PatternType.None:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "none"));
                        break;
                    case PatternType.Solid:
                        pattern.Add(new XAttribute(NoNamespace.patternType, "solid"));
                        break;
                }
                if (FgColor != null)
                    pattern.Add(FgColor.GetXElement(S.fgColor));
                if (BgColor != null)
                    pattern.Add(BgColor.GetXElement(S.bgColor));
                return new XElement(S.fill, pattern);
            }
        }

        public class GradientStop
        {
            private double Position;
            private ColorInfo Color;

            public GradientStop(double position, ColorInfo color)
            {
                Position = position;
                Color = color;
            }

            public XElement GetXElement()
            {
                return new XElement(S.stop, new XAttribute(NoNamespace.position, Position), Color.GetXElement(S.color));
            }
        }

        public class GradientFill
        {
            private bool PathGradient;
            private int LinearDegree;
            private double PathTop;
            private double PathLeft;
            private double PathBottom;
            private double PathRight;
            private List<GradientStop> Stops;

            public GradientFill(int degree)
            {
                PathGradient = false;
                LinearDegree = degree;
                Stops = new List<GradientStop>();
            }

            public GradientFill(double top, double left, double bottom, double right)
            {
                PathGradient = true;
                PathTop = top;
                PathLeft = left;
                PathBottom = bottom;
                PathRight = right;
                Stops = new List<GradientStop>();
            }

            public void AddStop(GradientStop stop)
            {
                Stops.Add(stop);
            }

            public XElement GetXElement()
            {
                XElement gradient = new XElement(S.gradientFill);
                if (PathGradient)
                {
                    gradient.Add(new XAttribute(NoNamespace.type, "path"),
                        new XAttribute(NoNamespace.left, PathLeft.ToString()), new XAttribute(NoNamespace.right, PathRight.ToString()),
                        new XAttribute(NoNamespace.top, PathTop.ToString()), new XAttribute(NoNamespace.bottom, PathBottom.ToString()));
                }
                else
                {
                    gradient.Add(new XAttribute(NoNamespace.degree, LinearDegree.ToString()));
                }
                foreach (GradientStop stop in Stops)
                    gradient.Add(stop.GetXElement());
                return new XElement(S.fill, gradient);
            }
        }

        public static int GetFillIndex(SpreadsheetDocument document, PatternFill patternFill)
        {
            return GetFillIndex(document, patternFill.GetXElement());
        }

        public static int GetFillIndex(SpreadsheetDocument document, GradientFill gradientFill)
        {
            return GetFillIndex(document, gradientFill.GetXElement());
        }

        private static int GetFillIndex(SpreadsheetDocument document, XElement fill)
        {
            XDocument styles = document.WorkbookPart.WorkbookStylesPart.GetXDocument();
            XElement fills = styles.Root.Element(S.fills);
            int index = Array.FindIndex(fills.Elements(S.fill).ToArray(),
                z => XElement.DeepEquals(z, fill));
            if (index != -1)
                return index;
            fills.Add(fill);
            fills.Attribute(NoNamespace.count).Value = fills.Elements(S.fill).Count().ToString();
            document.WorkbookPart.WorkbookStylesPart.PutXDocument();
            return fills.Elements(S.fill).Count() - 1;
        }

        public class BorderLine
        {
            public enum LineStyle
            {
                None, DashDot, DashDotDot, Dashed, Dotted, Double, Hair,
                Medium, MediumDashDot, MediumDashDotDot, MediumDashed, SlantDashDot, Thick, Thin
            };
            private LineStyle Style;
            private ColorInfo Color;

            public BorderLine(LineStyle style, ColorInfo color)
            {
                Style = style;
                Color = color;
            }

            public XElement GetXElement(XName name)
            {
                XElement line = new XElement(name);
                switch (Style)
                {
                    case LineStyle.DashDot:
                        line.Add(new XAttribute(NoNamespace.style, "dashDot"));
                        break;
                    case LineStyle.DashDotDot:
                        line.Add(new XAttribute(NoNamespace.style, "dashDotDot"));
                        break;
                    case LineStyle.Dashed:
                        line.Add(new XAttribute(NoNamespace.style, "dashed"));
                        break;
                    case LineStyle.Dotted:
                        line.Add(new XAttribute(NoNamespace.style, "dotted"));
                        break;
                    case LineStyle.Double:
                        line.Add(new XAttribute(NoNamespace.style, "double"));
                        break;
                    case LineStyle.Hair:
                        line.Add(new XAttribute(NoNamespace.style, "hair"));
                        break;
                    case LineStyle.Medium:
                        line.Add(new XAttribute(NoNamespace.style, "medium"));
                        break;
                    case LineStyle.MediumDashDot:
                        line.Add(new XAttribute(NoNamespace.style, "mediumDashDot"));
                        break;
                    case LineStyle.MediumDashDotDot:
                        line.Add(new XAttribute(NoNamespace.style, "mediumDashDotDot"));
                        break;
                    case LineStyle.MediumDashed:
                        line.Add(new XAttribute(NoNamespace.style, "mediumDashed"));
                        break;
                    case LineStyle.SlantDashDot:
                        line.Add(new XAttribute(NoNamespace.style, "slantDashDot"));
                        break;
                    case LineStyle.Thick:
                        line.Add(new XAttribute(NoNamespace.style, "thick"));
                        break;
                    case LineStyle.Thin:
                        line.Add(new XAttribute(NoNamespace.style, "thin"));
                        break;
                }
                line.Add(Color.GetXElement(S.color));
                return line;
            }
        }

        public class Border
        {
            public BorderLine Top { get; set; }
            public BorderLine Bottom { get; set; }
            public BorderLine Left { get; set; }
            public BorderLine Right { get; set; }
            public BorderLine Horizontal { get; set; }
            public BorderLine Vertical { get; set; }
            public BorderLine Diagonal { get; set; }
            public bool DiagonalDown { get; set; }
            public bool DiagonalUp { get; set; }
            public bool Outline { get; set; }

            public XElement GetXElement()
            {
                XElement border = new XElement(S.border);
                if (DiagonalDown)
                    border.Add(new XAttribute(NoNamespace.diagonalDown, "1"));
                if (DiagonalUp)
                    border.Add(new XAttribute(NoNamespace.diagonalUp, "1"));
                if (Outline)
                    border.Add(new XAttribute(NoNamespace.outline, "1"));
                if (Left == null)
                    border.Add(new XElement(S.left));
                else
                    border.Add(Left.GetXElement(S.left));
                if (Right == null)
                    border.Add(new XElement(S.right));
                else
                    border.Add(Right.GetXElement(S.right));
                if (Top == null)
                    border.Add(new XElement(S.top));
                else
                    border.Add(Top.GetXElement(S.top));
                if (Bottom == null)
                    border.Add(new XElement(S.bottom));
                else
                    border.Add(Bottom.GetXElement(S.bottom));
                if (Diagonal == null)
                    border.Add(new XElement(S.diagonal));
                else
                    border.Add(Diagonal.GetXElement(S.diagonal));
                if (Horizontal != null)
                    border.Add(Horizontal.GetXElement(S.horizontal));
                if (Vertical != null)
                    border.Add(Vertical.GetXElement(S.vertical));
                return border;
            }
        }

        public static int GetBorderIndex(SpreadsheetDocument document, Border b)
        {
            XElement border = b.GetXElement();
            XDocument styles = document.WorkbookPart.WorkbookStylesPart.GetXDocument();
            XElement borders = styles.Root.Element(S.borders);
            int index = Array.FindIndex(borders.Elements(S.border).ToArray(),
                z => XElement.DeepEquals(z, border));
            if (index != -1)
                return index;
            borders.Add(border);
            borders.Attribute(NoNamespace.count).Value = borders.Elements(S.border).Count().ToString();
            document.WorkbookPart.WorkbookStylesPart.PutXDocument();
            return borders.Elements(S.border).Count() - 1;
        }

        public static int GetStyleIndex(SpreadsheetDocument document, string styleName)
        {
            XDocument styles = document.WorkbookPart.WorkbookStylesPart.GetXDocument();
            string xfId = styles.Root.Element(S.cellStyles).Elements(S.cellStyle)
                .Where(t => t.Attribute(NoNamespace.name).Value == styleName)
                .FirstOrDefault().Attribute(NoNamespace.xfId).Value;
            XElement cellXfs = styles.Root.Element(S.cellXfs);
            int index = Array.FindIndex(cellXfs.Elements(S.xf).ToArray(),
                z => z.Attribute(NoNamespace.xfId).Value == xfId);
            if (index != -1)
                return index;
            XElement cellStyleXf = styles.Root.Element(S.cellStyleXfs).Elements(S.xf).ToArray()[Convert.ToInt32(xfId)];
            if (cellStyleXf != null)
            {   // Create new xf element under cellXfs
                cellXfs.Add(new XElement(S.xf, new XAttribute(NoNamespace.numFmtId, cellStyleXf.Attribute(NoNamespace.numFmtId).Value),
                    new XAttribute(NoNamespace.fontId, cellStyleXf.Attribute(NoNamespace.fontId).Value),
                    new XAttribute(NoNamespace.fillId, cellStyleXf.Attribute(NoNamespace.fillId).Value),
                    new XAttribute(NoNamespace.borderId, cellStyleXf.Attribute(NoNamespace.borderId).Value),
                    new XAttribute(NoNamespace.xfId, xfId)));
                cellXfs.Attribute(NoNamespace.count).Value = cellXfs.Elements(S.xf).Count().ToString();
                document.WorkbookPart.WorkbookStylesPart.PutXDocument();
                return cellXfs.Elements(S.xf).Count() - 1;
            }

            return 0;
        }

        public class CellAlignment
        {
            public enum Horizontal { General, Center, CenterContinuous, Distributed, Fill, Justify, Left, Right };
            public enum Vertical { Bottom, Center, Distributed, Justify, Top };

            public Horizontal HorizontalAlignment { get; set; }
            public int Indent { get; set; }
            public bool JustifyLastLine { get; set; }
            public int ReadingOrder { get; set; }
            public bool ShrinkToFit { get; set; }
            public int TextRotation { get; set; }
            public Vertical VerticalAlignment { get; set; }
            public bool WrapText { get; set; }

            public CellAlignment()
            {
                HorizontalAlignment = Horizontal.General;
                Indent = 0;
                JustifyLastLine = false;
                ReadingOrder = 0;
                ShrinkToFit = false;
                TextRotation = 0;
                VerticalAlignment = Vertical.Bottom;
                WrapText = false;
            }

            public XElement GetXElement()
            {
                XElement align = new XElement(S.alignment);
                switch (HorizontalAlignment)
                {
                    case Horizontal.Center:
                        align.Add(new XAttribute(NoNamespace.horizontal, "center"));
                        break;
                    case Horizontal.CenterContinuous:
                        align.Add(new XAttribute(NoNamespace.horizontal, "centerContinuous"));
                        break;
                    case Horizontal.Distributed:
                        align.Add(new XAttribute(NoNamespace.horizontal, "distributed"));
                        break;
                    case Horizontal.Fill:
                        align.Add(new XAttribute(NoNamespace.horizontal, "fill"));
                        break;
                    case Horizontal.Justify:
                        align.Add(new XAttribute(NoNamespace.horizontal, "justify"));
                        break;
                    case Horizontal.Left:
                        align.Add(new XAttribute(NoNamespace.horizontal, "left"));
                        break;
                    case Horizontal.Right:
                        align.Add(new XAttribute(NoNamespace.horizontal, "right"));
                        break;
                }
                if (Indent != 0)
                    align.Add(new XAttribute(NoNamespace.indent, Indent));
                if (JustifyLastLine)
                    align.Add(new XAttribute(NoNamespace.justifyLastLine, true));
                if (ReadingOrder != 0)
                    align.Add(new XAttribute(NoNamespace.readingOrder, ReadingOrder));
                if (ShrinkToFit)
                    align.Add(new XAttribute(NoNamespace.shrinkToFit, true));
                if (TextRotation != 0)
                    align.Add(new XAttribute(NoNamespace.textRotation, TextRotation));
                switch (VerticalAlignment)
                {
                    case Vertical.Center:
                        align.Add(new XAttribute(NoNamespace.vertical, "center"));
                        break;
                    case Vertical.Distributed:
                        align.Add(new XAttribute(NoNamespace.vertical, "distributed"));
                        break;
                    case Vertical.Justify:
                        align.Add(new XAttribute(NoNamespace.vertical, "justify"));
                        break;
                    case Vertical.Top:
                        align.Add(new XAttribute(NoNamespace.vertical, "top"));
                        break;
                }
                if (WrapText)
                    align.Add(new XAttribute(NoNamespace.wrapText, true));
                return align;
            }
        }

        public static int GetStyleIndex(SpreadsheetDocument document, int numFmt, int font, int fill, int border, CellAlignment alignment, bool hidden, bool locked)
        {
            XElement xf = new XElement(S.xf, new XAttribute(NoNamespace.numFmtId, numFmt),
                new XAttribute(NoNamespace.fontId, font), new XAttribute(NoNamespace.fillId, fill),
                new XAttribute(NoNamespace.borderId, border), new XAttribute(NoNamespace.xfId, 0),
                new XAttribute(NoNamespace.applyNumberFormat, (numFmt == 0) ? 0 : 1),
                new XAttribute(NoNamespace.applyFont, (font == 0) ? 0 : 1),
                new XAttribute(NoNamespace.applyFill, (fill == 0) ? 0 : 1),
                new XAttribute(NoNamespace.applyBorder, (border == 0) ? 0 : 1));
            if (alignment != null)
            {
                xf.Add(new XAttribute(NoNamespace.applyAlignment, "1"));
                xf.Add(alignment.GetXElement());
            }
            else
                xf.Add(new XAttribute(NoNamespace.applyAlignment, "0"));
            if (hidden || locked)
            {
                XElement prot = new XElement(S.protection);
                if (hidden)
                    prot.Add(new XAttribute(NoNamespace.hidden, true));
                if (locked)
                    prot.Add(new XAttribute(NoNamespace.locked, true));
                xf.Add(prot);
                xf.Add(new XAttribute(NoNamespace.applyProtection, "1"));
            }
            else
                xf.Add(new XAttribute(NoNamespace.applyProtection, "0"));

            XDocument styles = document.WorkbookPart.WorkbookStylesPart.GetXDocument();
            XElement cellXfs = styles.Root.Element(S.cellXfs);
            int index = Array.FindIndex(cellXfs.Elements(S.xf).ToArray(),
                z => XElement.DeepEquals(z, xf));
            if (index != -1)
                return index;
            cellXfs.Add(xf);
            cellXfs.Attribute(NoNamespace.count).Value = cellXfs.Elements(S.xf).Count().ToString();
            document.WorkbookPart.WorkbookStylesPart.PutXDocument();
            return cellXfs.Elements(S.xf).Count() - 1;
        }

        public static void CreateDefaultStyles(SpreadsheetDocument document)
        {
            // Create the style part
            WorkbookStylesPart stylesPart = document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.PutXDocument(new XDocument(XElement.Parse(
@"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<styleSheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>
  <fonts count='18'>
    <font>
      <sz val='11'/>
      <color theme='1'/>
      <name val='Calibri'/>
      <family val='2'/>
      <scheme val='minor'/>
    </font>
    <font>
      <sz val='11'/>
      <color theme='1'/>
      <name val='Calibri'/>
      <family val='2'/>
      <scheme val='minor'/>
    </font>
    <font>
      <b/>
      <sz val='18'/>
      <color theme='3'/>
      <name val='Cambria'/>
      <family val='2'/>
      <scheme val='major'/>
    </font>
    <font>
      <b/>
      <sz val='15'/>
      <color theme='3'/>
      <name val='Calibri'/>
      <family val='2'/>
      <scheme val='minor'/>
    </font>
    <font>
      <b/>
      <sz val='13'/>
      <color theme='3'/>
      <name val='Calibri'/>
      <family val='2'/>
      <scheme val='minor'/>
    </font>
    <font>
      <b/>
      <sz val='11'/>
      <color theme='3'/>
      <name val='Calibri'/>
      <family val='2'/>
      <scheme val='minor'/>
    </font>
    <font>
      <sz val='11'/>
      <color rgb='FF006100'/>
      <name val='Calibri'/>
      <family val='2'/>
      <scheme val='minor'/>
    </font>
    <font>
      <sz val='11'/>
      <color rgb='FF9C0006'/>
      <name val='Calibri'/>
      <family val='2'/>
      <scheme val='minor'/>
    </font>
    <font>
      <sz val='11'/>
      <color rgb='FF9C6500'/>
      <name val='Calibri'/>
      <family val='2'/>
      <scheme val='minor'/>
    </font>
    <font>
      <sz val='11'/>
      <color rgb='FF3F3F76'/>
      <name val='Calibri'/>
      <family val='2'/>
      <scheme val='minor'/>
    </font>
    <font>
      <b/>
      <sz val='11'/>
      <color rgb='FF3F3F3F'/>
      <name val='Calibri'/>
      <family val='2'/>
      <scheme val='minor'/>
    </font>
    <font>
      <b/>
      <sz val='11'/>
      <color rgb='FFFA7D00'/>
      <name val='Calibri'/>
      <family val='2'/>
      <scheme val='minor'/>
    </font>
    <font>
      <sz val='11'/>
      <color rgb='FFFA7D00'/>
      <name val='Calibri'/>
      <family val='2'/>
      <scheme val='minor'/>
    </font>
    <font>
      <b/>
      <sz val='11'/>
      <color theme='0'/>
      <name val='Calibri'/>
      <family val='2'/>
      <scheme val='minor'/>
    </font>
    <font>
      <sz val='11'/>
      <color rgb='FFFF0000'/>
      <name val='Calibri'/>
      <family val='2'/>
      <scheme val='minor'/>
    </font>
    <font>
      <i/>
      <sz val='11'/>
      <color rgb='FF7F7F7F'/>
      <name val='Calibri'/>
      <family val='2'/>
      <scheme val='minor'/>
    </font>
    <font>
      <b/>
      <sz val='11'/>
      <color theme='1'/>
      <name val='Calibri'/>
      <family val='2'/>
      <scheme val='minor'/>
    </font>
    <font>
      <sz val='11'/>
      <color theme='0'/>
      <name val='Calibri'/>
      <family val='2'/>
      <scheme val='minor'/>
    </font>
  </fonts>
  <fills count='33'>
    <fill>
      <patternFill patternType='none'/>
    </fill>
    <fill>
      <patternFill patternType='gray125'/>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor rgb='FFC6EFCE'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor rgb='FFFFC7CE'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor rgb='FFFFEB9C'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor rgb='FFFFCC99'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor rgb='FFF2F2F2'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor rgb='FFA5A5A5'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor rgb='FFFFFFCC'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='4'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='4' tint='0.79998168889431442'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='4' tint='0.59999389629810485'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='4' tint='0.39997558519241921'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='5'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='5' tint='0.79998168889431442'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='5' tint='0.59999389629810485'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='5' tint='0.39997558519241921'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='6'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='6' tint='0.79998168889431442'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='6' tint='0.59999389629810485'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='6' tint='0.39997558519241921'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='7'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='7' tint='0.79998168889431442'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='7' tint='0.59999389629810485'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='7' tint='0.39997558519241921'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='8'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='8' tint='0.79998168889431442'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='8' tint='0.59999389629810485'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='8' tint='0.39997558519241921'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='9'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='9' tint='0.79998168889431442'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='9' tint='0.59999389629810485'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType='solid'>
        <fgColor theme='9' tint='0.39997558519241921'/>
        <bgColor indexed='65'/>
      </patternFill>
    </fill>
  </fills>
  <borders count='10'>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom style='thick'>
        <color theme='4'/>
      </bottom>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom style='thick'>
        <color theme='4' tint='0.499984740745262'/>
      </bottom>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom style='medium'>
        <color theme='4' tint='0.39997558519241921'/>
      </bottom>
      <diagonal/>
    </border>
    <border>
      <left style='thin'>
        <color rgb='FF7F7F7F'/>
      </left>
      <right style='thin'>
        <color rgb='FF7F7F7F'/>
      </right>
      <top style='thin'>
        <color rgb='FF7F7F7F'/>
      </top>
      <bottom style='thin'>
        <color rgb='FF7F7F7F'/>
      </bottom>
      <diagonal/>
    </border>
    <border>
      <left style='thin'>
        <color rgb='FF3F3F3F'/>
      </left>
      <right style='thin'>
        <color rgb='FF3F3F3F'/>
      </right>
      <top style='thin'>
        <color rgb='FF3F3F3F'/>
      </top>
      <bottom style='thin'>
        <color rgb='FF3F3F3F'/>
      </bottom>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom style='double'>
        <color rgb='FFFF8001'/>
      </bottom>
      <diagonal/>
    </border>
    <border>
      <left style='double'>
        <color rgb='FF3F3F3F'/>
      </left>
      <right style='double'>
        <color rgb='FF3F3F3F'/>
      </right>
      <top style='double'>
        <color rgb='FF3F3F3F'/>
      </top>
      <bottom style='double'>
        <color rgb='FF3F3F3F'/>
      </bottom>
      <diagonal/>
    </border>
    <border>
      <left style='thin'>
        <color rgb='FFB2B2B2'/>
      </left>
      <right style='thin'>
        <color rgb='FFB2B2B2'/>
      </right>
      <top style='thin'>
        <color rgb='FFB2B2B2'/>
      </top>
      <bottom style='thin'>
        <color rgb='FFB2B2B2'/>
      </bottom>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style='thin'>
        <color theme='4'/>
      </top>
      <bottom style='double'>
        <color theme='4'/>
      </bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count='42'>
    <xf numFmtId='0' fontId='0' fillId='0' borderId='0'/>
    <xf numFmtId='0' fontId='2' fillId='0' borderId='0' applyNumberFormat='0' applyFill='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='3' fillId='0' borderId='1' applyNumberFormat='0' applyFill='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='4' fillId='0' borderId='2' applyNumberFormat='0' applyFill='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='5' fillId='0' borderId='3' applyNumberFormat='0' applyFill='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='5' fillId='0' borderId='0' applyNumberFormat='0' applyFill='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='6' fillId='2' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='7' fillId='3' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='8' fillId='4' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='9' fillId='5' borderId='4' applyNumberFormat='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='10' fillId='6' borderId='5' applyNumberFormat='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='11' fillId='6' borderId='4' applyNumberFormat='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='12' fillId='0' borderId='6' applyNumberFormat='0' applyFill='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='13' fillId='7' borderId='7' applyNumberFormat='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='14' fillId='0' borderId='0' applyNumberFormat='0' applyFill='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='1' fillId='8' borderId='8' applyNumberFormat='0' applyFont='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='15' fillId='0' borderId='0' applyNumberFormat='0' applyFill='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='16' fillId='0' borderId='9' applyNumberFormat='0' applyFill='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='17' fillId='9' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='1' fillId='10' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='1' fillId='11' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='17' fillId='12' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='17' fillId='13' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='1' fillId='14' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='1' fillId='15' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='17' fillId='16' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='17' fillId='17' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='1' fillId='18' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='1' fillId='19' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='17' fillId='20' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='17' fillId='21' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='1' fillId='22' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='1' fillId='23' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='17' fillId='24' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='17' fillId='25' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='1' fillId='26' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='1' fillId='27' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='17' fillId='28' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='17' fillId='29' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='1' fillId='30' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='1' fillId='31' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
    <xf numFmtId='0' fontId='17' fillId='32' borderId='0' applyNumberFormat='0' applyBorder='0' applyAlignment='0' applyProtection='0'/>
  </cellStyleXfs>
  <cellXfs count='1'>
    <xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0'/>
  </cellXfs>
  <cellStyles count='42'>
    <cellStyle name='20% - Accent1' xfId='19' builtinId='30' customBuiltin='1'/>
    <cellStyle name='20% - Accent2' xfId='23' builtinId='34' customBuiltin='1'/>
    <cellStyle name='20% - Accent3' xfId='27' builtinId='38' customBuiltin='1'/>
    <cellStyle name='20% - Accent4' xfId='31' builtinId='42' customBuiltin='1'/>
    <cellStyle name='20% - Accent5' xfId='35' builtinId='46' customBuiltin='1'/>
    <cellStyle name='20% - Accent6' xfId='39' builtinId='50' customBuiltin='1'/>
    <cellStyle name='40% - Accent1' xfId='20' builtinId='31' customBuiltin='1'/>
    <cellStyle name='40% - Accent2' xfId='24' builtinId='35' customBuiltin='1'/>
    <cellStyle name='40% - Accent3' xfId='28' builtinId='39' customBuiltin='1'/>
    <cellStyle name='40% - Accent4' xfId='32' builtinId='43' customBuiltin='1'/>
    <cellStyle name='40% - Accent5' xfId='36' builtinId='47' customBuiltin='1'/>
    <cellStyle name='40% - Accent6' xfId='40' builtinId='51' customBuiltin='1'/>
    <cellStyle name='60% - Accent1' xfId='21' builtinId='32' customBuiltin='1'/>
    <cellStyle name='60% - Accent2' xfId='25' builtinId='36' customBuiltin='1'/>
    <cellStyle name='60% - Accent3' xfId='29' builtinId='40' customBuiltin='1'/>
    <cellStyle name='60% - Accent4' xfId='33' builtinId='44' customBuiltin='1'/>
    <cellStyle name='60% - Accent5' xfId='37' builtinId='48' customBuiltin='1'/>
    <cellStyle name='60% - Accent6' xfId='41' builtinId='52' customBuiltin='1'/>
    <cellStyle name='Accent1' xfId='18' builtinId='29' customBuiltin='1'/>
    <cellStyle name='Accent2' xfId='22' builtinId='33' customBuiltin='1'/>
    <cellStyle name='Accent3' xfId='26' builtinId='37' customBuiltin='1'/>
    <cellStyle name='Accent4' xfId='30' builtinId='41' customBuiltin='1'/>
    <cellStyle name='Accent5' xfId='34' builtinId='45' customBuiltin='1'/>
    <cellStyle name='Accent6' xfId='38' builtinId='49' customBuiltin='1'/>
    <cellStyle name='Bad' xfId='7' builtinId='27' customBuiltin='1'/>
    <cellStyle name='Calculation' xfId='11' builtinId='22' customBuiltin='1'/>
    <cellStyle name='Check Cell' xfId='13' builtinId='23' customBuiltin='1'/>
    <cellStyle name='Explanatory Text' xfId='16' builtinId='53' customBuiltin='1'/>
    <cellStyle name='Good' xfId='6' builtinId='26' customBuiltin='1'/>
    <cellStyle name='Heading 1' xfId='2' builtinId='16' customBuiltin='1'/>
    <cellStyle name='Heading 2' xfId='3' builtinId='17' customBuiltin='1'/>
    <cellStyle name='Heading 3' xfId='4' builtinId='18' customBuiltin='1'/>
    <cellStyle name='Heading 4' xfId='5' builtinId='19' customBuiltin='1'/>
    <cellStyle name='Input' xfId='9' builtinId='20' customBuiltin='1'/>
    <cellStyle name='Linked Cell' xfId='12' builtinId='24' customBuiltin='1'/>
    <cellStyle name='Neutral' xfId='8' builtinId='28' customBuiltin='1'/>
    <cellStyle name='Normal' xfId='0' builtinId='0'/>
    <cellStyle name='Note' xfId='15' builtinId='10' customBuiltin='1'/>
    <cellStyle name='Output' xfId='10' builtinId='21' customBuiltin='1'/>
    <cellStyle name='Title' xfId='1' builtinId='15' customBuiltin='1'/>
    <cellStyle name='Total' xfId='17' builtinId='25' customBuiltin='1'/>
    <cellStyle name='Warning Text' xfId='14' builtinId='11' customBuiltin='1'/>
  </cellStyles>
  <dxfs count='0'/>
  <tableStyles count='0' defaultTableStyle='TableStyleMedium9' defaultPivotStyle='PivotStyleLight16'/>
</styleSheet>")));
        }

        /// <summary>
        /// Creates a worksheet document and inserts data into it
        /// </summary>
        /// <param name="headerList">List of values that will act as the header</param>
        /// <param name="valueTable">Values for worksheet content</param>
        /// <param name="headerRow">Header row</param>
        /// <returns></returns>
        internal static WorksheetPart Create(SpreadsheetDocument document, List<string> headerList, string[][] valueTable, int headerRow)
        {
            XDocument xDocument = CreateEmptyWorksheet();
            
            for (int i = 0; i < headerList.Count; i++)
            {
                AddValue(xDocument, headerRow, i + 1, headerList[i]);
            }
            int rows = valueTable.GetLength(0);
            int cols = valueTable[0].GetLength(0);

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    AddValue(xDocument, i + headerRow + 1, j + 1, valueTable[i][j]);
                }
            }
            WorksheetPart part = Add(document, xDocument);
            return part;
        }

        /// <summary>
        /// Creates element structure needed to describe an empty worksheet
        /// </summary>
        /// <returns>Document with contents for an empty worksheet</returns>
        private static XDocument CreateEmptyWorksheet()
        {
            XDocument document =
                new XDocument(
                    new XElement(ns + "worksheet",
                        new XAttribute("xmlns", ns),
                        new XAttribute(XNamespace.Xmlns + "r", relationshipsns),
                        new XElement(ns + "sheetData")
                    )
                );
            return document;
        }

        /// <summary>
        /// Adds a value to a cell inside a worksheet document
        /// </summary>
        /// <param name="worksheet">document to add values</param>
        /// <param name="row">Row</param>
        /// <param name="column">Column</param>
        /// <param name="value">Value to add</param>
        private static void AddValue(XDocument worksheet, int row, int column, string value)
        {
            //Set the cell reference
            string cellReference = GetColumnId(column) + row.ToString();
            double numericValue;
            //Determining if value for cell is text or numeric
            bool valueIsNumeric = double.TryParse(value, out numericValue);

            //Creating the new cell element (markup)
            XElement newCellXElement = valueIsNumeric ?
                    new XElement(ns + "c",
                        new XAttribute("r", cellReference),
                        new XElement(ns + "v", numericValue)
                    )
                :
                    new XElement(ns + "c",
                        new XAttribute("r", cellReference),
                        new XAttribute("t", "inlineStr"),
                        new XElement(ns + "is",
                            new XElement(ns + "t", value)
                        )
                    );

            // Find the row containing the cell to add the value to
            XName rowName = "r";
            XElement rowElement =
                worksheet.Root
                    .Element(ns + "sheetData")
                    .Elements(ns + "row")
                    .Where(
                        t => t.Attribute(rowName).Value == row.ToString()
                    )
                    .FirstOrDefault();

            if (rowElement == null)
            {
                //row element does not exist
                //create a new one
                rowElement = CreateEmptyRow(row);

                //row elements must appear in order inside sheetData element
                if (worksheet.Root
                 .Element(ns + "sheetData").HasElements)
                {   //if there are more rows already defined at sheetData element
                    //find the row with the inmediate higher index for the row containing the cell to set the value to
                    XElement rowAfterElement = FindRowAfter(worksheet, row);
                    //if there is a row with an inmediate higher index already defined at sheetData
                    if (rowAfterElement != null)
                    {
                        //add the new row before the row with an inmediate higher index
                        rowAfterElement.AddBeforeSelf(rowElement);
                    }
                    else
                    {   //this row is going to be the one with the highest index (add it as the last element for sheetData)
                        worksheet.Root.Element(ns + "sheetData").Elements(ns + "row").Last().AddAfterSelf(rowElement);
                    }
                }
                else
                {   //there are no other rows already defined at sheetData
                    //Add a new row elemento to sheetData
                    worksheet
                        .Root
                        .Element(ns + "sheetData")
                        .Add(
                            rowElement //= CreateEmptyRow(row)
                        );
                }

                //Add the new cell to the row Element
                rowElement.Add(newCellXElement);
            }
            else
            {
                //row containing the cell to set the value to is already defined at sheetData
                //look if cell already exist at that row
                XElement currentCellXElement = rowElement
                    .Elements(ns + "c")
                    .Where(
                        t => t.Attribute("r").Value == cellReference
                    ).FirstOrDefault();

                if (currentCellXElement == null)
                {   //cell element does not exist at row indicated as parameter
                    //find the inmediate right column for the cell to set the value to
                    XElement columnAfterXElement = FindColumAfter(worksheet, row, column);
                    if (columnAfterXElement != null)
                    {
                        //Insert the new cell before the inmediate right column
                        columnAfterXElement.AddBeforeSelf(newCellXElement);
                    }
                    else
                    {   //There is no inmediate right cell 
                        //Add the new cell as the last element for the row
                        rowElement.Add(newCellXElement);
                    }
                }
                else
                {
                    //cell alreay exist
                    //replace the current cell with that with the new value
                    currentCellXElement.ReplaceWith(newCellXElement);
                }
            }
        }

        /// <summary>
        /// Adds a given worksheet to the document
        /// </summary>
        /// <param name="worksheet">Worksheet document to add</param>
        /// <returns>Worksheet part just added</returns>
        public static WorksheetPart Add(SpreadsheetDocument doc, XDocument worksheet)
        {
            // Associates base content to a new worksheet part
            WorkbookPart workbook = doc.WorkbookPart;
            WorksheetPart worksheetPart = workbook.AddNewPart<WorksheetPart>();
            worksheetPart.PutXDocument(worksheet);

            // Associates the worksheet part to the workbook part
            XDocument document = doc.WorkbookPart.GetXDocument();
            int sheetId =
                document.Root
                .Element(ns + "sheets")
                .Elements(ns + "sheet")
                .Count() + 1;

            int worksheetCount =
                document.Root
                .Element(ns + "sheets")
                .Elements(ns + "sheet")
                .Where(
                    t =>
                        t.Attribute("name").Value.StartsWith("sheet", StringComparison.OrdinalIgnoreCase)
                )
                .Count() + 1;

            // Adds content to workbook document to reference worksheet document
            document.Root
                .Element(ns + "sheets")
                .Add(
                    new XElement(ns + "sheet",
                        new XAttribute("name", string.Format("sheet{0}", worksheetCount)),
                        new XAttribute("sheetId", sheetId),
                        new XAttribute(relationshipsns + "id", workbook.GetIdOfPart(worksheetPart))
                    )
                );
            doc.WorkbookPart.PutXDocument();
            return worksheetPart;
        }
    }
}
