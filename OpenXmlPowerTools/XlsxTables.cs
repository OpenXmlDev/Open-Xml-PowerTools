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
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace OpenXmlPowerTools
{
    public class Table
    {
        public int Id { get; set; }
        public string TableName { get; set; }
        public string DisplayName { get; set; }
        public XElement TableStyleInfo { get; set; }
        public string Ref { get; set; }
        public int LeftColumn { get; set; }
        public int RightColumn { get; set; }
        public int TopRow { get; set; }
        public int BottomRow { get; set; }
        public int? HeaderRowCount { get; set; }
        public int? TotalsRowCount { get; set; }
        public string TableType { get; set; }  // external data query, data in worksheet, or XML data
        public TableDefinitionPart TableDefinitionPart { get; set; }
        public WorksheetPart Parent { get; set; }
        public Table(WorksheetPart parent) { Parent = parent; }
        public IEnumerable<TableColumn> TableColumns()
        {
            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            return TableDefinitionPart
                .GetXDocument()
                .Root
                .Element(x + "tableColumns")
                .Elements(x + "tableColumn")
                .Select((c, i) =>
                    new TableColumn(this)
                    {
                        Id = (int)c.Attribute("id"),
                        ColumnNumber = this.LeftColumn + i,
                        Name = (string)c.Attribute("name"),
                        DataDxfId = (int?)c.Attribute("dataDxfId"),
                        QueryTableFieldId = (int?)c.Attribute("queryTableFieldId"),
                        UniqueName = (string)c.Attribute("uniqueName"),
                        ColumnIndex = i,
                    }
                );
        }
        public IEnumerable<TableRow> TableRows()
        {
            string refStart = Ref.Split(':').First();
            int rowStart = Int32.Parse(XlsxTables.SplitAddress(refStart)[1]);
            string refEnd = Ref.Split(':').ElementAt(1);
            int rowEnd = Int32.Parse(XlsxTables.SplitAddress(refEnd)[1]);
            int headerRowsCount = HeaderRowCount == null ? 0 : (int)HeaderRowCount;
            int totalRowsCount = TotalsRowCount == null ? 0 : (int)TotalsRowCount;
            return Parent
                .Rows()
                .Skip(headerRowsCount)
                .SkipLast(totalRowsCount)
                .Where(r =>
                {
                    int rowId = Int32.Parse(r.RowId);
                    return rowId >= rowStart && rowId <= rowEnd;
                }
                )
                .Select(r => new TableRow(this) { Row = r });
        }
    }

    public class TableColumn
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int? DataDxfId { get; set; }
        public int? QueryTableFieldId { get; set; }
        public string UniqueName { get; set; }
        public int ColumnNumber { get; set; }
        public int ColumnIndex { get; set; }
        public Table Parent { get; set; }
        public TableColumn(Table parent) { Parent = parent; }
    }

    public class TableRow
    {
        public Row Row { get; set; }
        public Table Parent { get; set; }
        public TableRow(Table parent) { Parent = parent; }
        public TableCell this[string columnName]
        {
            get
            {
                TableColumn tc = Parent
                    .TableColumns()
                    .Where(x => x.Name.ToLower() == columnName.ToLower())
                    .FirstOrDefault();
                if (tc == null)
                    throw new Exception("Invalid column name: " + columnName);
                string[] refs = Parent.Ref.Split(':');
                string[] startRefs = XlsxTables.SplitAddress(refs[0]);
                string columnAddress = XlsxTables.IndexToColumnAddress(XlsxTables.ColumnAddressToIndex(startRefs[0]) + tc.ColumnIndex);
                Cell cell = Row.Cells().Where(c => c.ColumnAddress == columnAddress).FirstOrDefault();
                if (cell != null)
                {
                    if (cell.Type == "s")
                        return new TableCell(cell.SharedString);
                    else
                        return new TableCell(cell.Value);
                }
                else
                    return new TableCell("");
            }
        }
    }

    public class TableCell : IEquatable<TableCell>
    {
        public string Value { get; set; }
        public TableCell(string v)
        {
            Value = v;
        }
        public override string ToString()
        {
            return Value;
        }
        public override bool Equals(object obj)
        {
            return this.Value == ((TableCell)obj).Value;
        }
        bool IEquatable<TableCell>.Equals(TableCell other)
        {
            return this.Value == other.Value;
        }
        public override int GetHashCode()
        {
            return this.Value.GetHashCode();
        }
        public static bool operator ==(TableCell left, TableCell right)
        {
            if ((object)left != (object)right) return false;
            return left.Value == right.Value;
        }
        public static bool operator !=(TableCell left, TableCell right)
        {
            if ((object)left != (object)right) return false;
            return left.Value != right.Value;
        }
        public static explicit operator string(TableCell cell)
        {
            if (cell == null) return null;
            return cell.Value;
        }
        public static explicit operator bool(TableCell cell)
        {
            if (cell == null) throw new ArgumentNullException("TableCell");
            return cell.Value == "1";
        }
        public static explicit operator bool?(TableCell cell)
        {
            if (cell == null) return null;
            return cell.Value == "1";
        }
        public static explicit operator int(TableCell cell)
        {
            if (cell == null) throw new ArgumentNullException("TableCell");
            return Int32.Parse(cell.Value);
        }
        public static explicit operator int?(TableCell cell)
        {
            if (cell == null) return null;
            return Int32.Parse(cell.Value);
        }
        public static explicit operator uint(TableCell cell)
        {
            if (cell == null) throw new ArgumentNullException("TableCell");
            return UInt32.Parse(cell.Value);
        }
        public static explicit operator uint?(TableCell cell)
        {
            if (cell == null) return null;
            return UInt32.Parse(cell.Value);
        }
        public static explicit operator long(TableCell cell)
        {
            if (cell == null) throw new ArgumentNullException("TableCell");
            return Int64.Parse(cell.Value);
        }
        public static explicit operator long?(TableCell cell)
        {
            if (cell == null) return null;
            return Int64.Parse(cell.Value);
        }
        public static explicit operator ulong(TableCell cell)
        {
            if (cell == null) throw new ArgumentNullException("TableCell");
            return UInt64.Parse(cell.Value);
        }
        public static explicit operator ulong?(TableCell cell)
        {
            if (cell == null) return null;
            return UInt64.Parse(cell.Value);
        }
        public static explicit operator float(TableCell cell)
        {
            if (cell == null) throw new ArgumentNullException("TableCell");
            return Single.Parse(cell.Value);
        }
        public static explicit operator float?(TableCell cell)
        {
            if (cell == null) return null;
            return Single.Parse(cell.Value);
        }
        public static explicit operator double(TableCell cell)
        {
            if (cell == null) throw new ArgumentNullException("TableCell");
            return Double.Parse(cell.Value);
        }
        public static explicit operator double?(TableCell cell)
        {
            if (cell == null) return null;
            return Double.Parse(cell.Value);
        }
        public static explicit operator decimal(TableCell cell)
        {
            if (cell == null) throw new ArgumentNullException("TableCell");
            return Decimal.Parse(cell.Value);
        }
        public static explicit operator decimal?(TableCell cell)
        {
            if (cell == null) return null;
            return Decimal.Parse(cell.Value);
        }
        public static implicit operator DateTime(TableCell cell)
        {
            if (cell == null) throw new ArgumentNullException("TableCell");
            return new DateTime(1900, 1, 1).AddDays(Int32.Parse(cell.Value) - 2);
        }
        public static implicit operator DateTime?(TableCell cell)
        {
            if (cell == null) return null;
            return new DateTime(1900, 1, 1).AddDays(Int32.Parse(cell.Value) - 2);
        }
    }

    public class Row
    {
        public XElement RowElement { get; set; }
        public string RowId { get; set; }
        public string Spans { get; set; }
        public List<Cell> Cells()
        {
            XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            SpreadsheetDocument doc = (SpreadsheetDocument)Parent.OpenXmlPackage;
            SharedStringTablePart sharedStringTable = doc.WorkbookPart.SharedStringTablePart;
            IEnumerable<XElement> cells = this.RowElement.Elements(S.c);
            var r = cells
                .Select(cell => {
                    var cellType = (string)cell.Attribute("t");
                    var sharedString = cellType == "s" ?
                        sharedStringTable
                        .GetXDocument()
                        .Root
                        .Elements(s + "si")
                        .Skip((int)cell.Element(s + "v"))
                        .First()
                        .Descendants(s + "t")
                        .StringConcatenate(e => (string)e)
                        : null;
                    var column = (string)cell.Attribute("r");
                    var columnAddress = column.Split('0', '1', '2', '3', '4', '5', '6', '7', '8', '9').First();
                    var columnIndex = XlsxTables.ColumnAddressToIndex(columnAddress);
                    var newCell = new Cell(this)
                    {
                        CellElement = cell,
                        Row = (string)RowElement.Attribute("r"),
                        Column = column,
                        ColumnAddress = columnAddress,
                        ColumnIndex = columnIndex,
                        Type = cellType,
                        Formula = (string)cell.Element(S.f),
                        Style = (int?)cell.Attribute("s"),
                        Value = (string)cell.Element(S.v),
                        SharedString = sharedString
                    };
                    return newCell;
                });
            var ra = r.ToList();
            return ra;
        }
        public WorksheetPart Parent { get; set; }
        public Row(WorksheetPart parent) { Parent = parent; }
    }

    public class Cell
    {
        public XElement CellElement { get; set; }
        public string Row { get; set; }
        public string Column { get; set; }
        public string ColumnAddress { get; set; }
        public int ColumnIndex { get; set; }
        public string Type { get; set; }
        public string Value { get; set; }
        public string Formula { get; set; }
        public int? Style { get; set; }
        public string SharedString { get; set; }
        public Row Parent { get; set; }
        public Cell(Row parent) { Parent = parent; }
    }

    public static class XlsxTables
    {
        public static IEnumerable<Table> Tables(this SpreadsheetDocument spreadsheet)
        {
            foreach (var worksheetPart in spreadsheet.WorkbookPart.WorksheetParts)
                foreach (var table in worksheetPart.TableDefinitionParts)
                {
                    XDocument tableDefDoc = table.GetXDocument();

                    Table t = new Table(worksheetPart)
                    {
                        Id = (int)tableDefDoc.Root.Attribute("id"),
                        TableName = (string)tableDefDoc.Root.Attribute("name"),
                        DisplayName = (string)tableDefDoc.Root.Attribute("displayName"),
                        TableStyleInfo = tableDefDoc.Root.Element(S.tableStyleInfo),
                        Ref = (string)tableDefDoc.Root.Attribute("ref"),
                        TotalsRowCount = (int?)tableDefDoc.Root.Attribute("totalsRowCount"),
                        //HeaderRowCount = (int?)tableDefDoc.Root.Attribute("headerRowCount"),
                        HeaderRowCount = 1,  // currently there always is a header row
                        TableType = (string)tableDefDoc.Root.Attribute("tableType"),
                        TableDefinitionPart = table
                    };
                    int leftColumn, topRow, rightColumn, bottomRow;
                    ParseRange(t.Ref, out leftColumn, out topRow, out rightColumn, out bottomRow);
                    t.LeftColumn = leftColumn;
                    t.TopRow = topRow;
                    t.RightColumn = rightColumn;
                    t.BottomRow = bottomRow;
                    yield return t;
                }
        }

        public static void ParseRange(string theRef, out int leftColumn, out int topRow, out int rightColumn, out int bottomRow)
        {
            // C5:E7
            var spl = theRef.Split(':');
            string refStart = spl.First();
            var refStartSplit = XlsxTables.SplitAddress(refStart);
            leftColumn = XlsxTables.ColumnAddressToIndex(refStartSplit[0]);
            topRow = Int32.Parse(refStartSplit[1]);

            string refEnd = spl.ElementAt(1);
            var refEndSplit = XlsxTables.SplitAddress(refEnd);
            rightColumn = XlsxTables.ColumnAddressToIndex(refEndSplit[0]);
            bottomRow = Int32.Parse(refEndSplit[1]);
        }

        public static Table Table(this SpreadsheetDocument spreadsheet,
            string tableName)
        {
            return spreadsheet.Tables().Where(t => t.TableName.ToLower() == tableName.ToLower()).FirstOrDefault();
        }

        public static IEnumerable<Row> Rows(this WorksheetPart worksheetPart)
        {
            XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var rows = worksheetPart
                .GetXDocument()
                .Root
                .Elements(S.sheetData)
                .Elements(S.row)
                .Select(r =>
                {
                    var row = new Row(worksheetPart)
                    {
                        RowElement = r,
                        RowId = (string)r.Attribute("r"),
                        Spans = (string)r.Attribute("spans")
                    };
                    return row;
                });
            return rows;
        }

        public static string[] SplitAddress(string address)
        {
            int i;
            for (i = 0; i < address.Length; i++)
                if (address[i] >= '0' && address[i] <= '9')
                    break;
            if (i == address.Length)
                throw new FileFormatException("Invalid spreadsheet.  Bad cell address.");
            return new[] {
                address.Substring(0, i),
                address.Substring(i)
            };
        }

        public static string IndexToColumnAddress(int index)
        {
            if (index < 26)
            {
                char c = (char)((int)'A' + index);
                string s = new string(c, 1);
                return s;
            }
            if (index < 702)
            {
                int i = index - 26;
                int i1 = (int)(i / 26);
                int i2 = i % 26;
                string s = new string((char)((int)'A' + i1), 1) +
                    new string((char)((int)'A' + i2), 1);
                return s;
            }
            if (index < 18278)
            {
                int i = index - 702;
                int i1 = (int)(i / 676);
                i = i - i1 * 676;
                int i2 = (int)(i / 26);
                int i3 = i % 26;
                string s = new string((char)((int)'A' + i1), 1) +
                    new string((char)((int)'A' + i2), 1) +
                    new string((char)((int)'A' + i3), 1);
                return s;
            }
            throw new Exception("Invalid column address");
        }

        public static int ColumnAddressToIndex(string columnAddress)
        {
            if (columnAddress.Length == 1)
            {
                char c = columnAddress[0];
                int i = c - 'A';
                return i;
            }
            if (columnAddress.Length == 2)
            {
                char c1 = columnAddress[0];
                char c2 = columnAddress[1];
                int i1 = c1 - 'A';
                int i2 = c2 - 'A';
                return (i1 + 1) * 26 + i2;
            }
            if (columnAddress.Length == 3)
            {
                char c1 = columnAddress[0];
                char c2 = columnAddress[1];
                char c3 = columnAddress[2];
                int i1 = c1 - 'A';
                int i2 = c2 - 'A';
                int i3 = c3 - 'A';
                return (i1 + 1) * 676 + (i2 + 1) * 26 + i3;
            }
            throw new FileFormatException("Invalid spreadsheet: Invalid column address.");
        }
    }
}
