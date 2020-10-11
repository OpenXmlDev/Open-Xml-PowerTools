using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

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

        public Table(WorksheetPart parent)
        {
            Parent = parent;
        }

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
                        ColumnNumber = LeftColumn + i,
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
            var refStart = Ref.Split(':').First();
            var rowStart = int.Parse(XlsxTables.SplitAddress(refStart)[1]);
            var refEnd = Ref.Split(':').ElementAt(1);
            var rowEnd = int.Parse(XlsxTables.SplitAddress(refEnd)[1]);
            var headerRowsCount = HeaderRowCount == null ? 0 : (int)HeaderRowCount;
            var totalRowsCount = TotalsRowCount == null ? 0 : (int)TotalsRowCount;
            return Parent
                .Rows()
                .Skip(headerRowsCount)
                .SkipLast(totalRowsCount)
                .Where(r =>
                {
                    var rowId = int.Parse(r.RowId);
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

        public TableColumn(Table parent)
        {
            Parent = parent;
        }
    }

    public class TableRow
    {
        public Row Row { get; set; }
        public Table Parent { get; set; }

        public TableRow(Table parent)
        {
            Parent = parent;
        }

        public TableCell this[string columnName]
        {
            get
            {
                var tc = Parent
                    .TableColumns()
                    .Where(x => x.Name.ToLower() == columnName.ToLower())
                    .FirstOrDefault();
                if (tc == null)
                {
                    throw new Exception("Invalid column name: " + columnName);
                }

                var refs = Parent.Ref.Split(':');
                var startRefs = XlsxTables.SplitAddress(refs[0]);
                var columnAddress = XlsxTables.IndexToColumnAddress(XlsxTables.ColumnAddressToIndex(startRefs[0]) + tc.ColumnIndex);
                var cell = Row.Cells().Where(c => c.ColumnAddress == columnAddress).FirstOrDefault();
                if (cell != null)
                {
                    if (cell.Type == "s")
                    {
                        return new TableCell(cell.SharedString);
                    }
                    else
                    {
                        return new TableCell(cell.Value);
                    }
                }
                else
                {
                    return new TableCell("");
                }
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
            return Value == ((TableCell)obj).Value;
        }

        bool IEquatable<TableCell>.Equals(TableCell other)
        {
            return Value == other.Value;
        }

        public override int GetHashCode()
        {
            return Value.GetHashCode();
        }

        public static bool operator ==(TableCell left, TableCell right)
        {
            if (left != (object)right)
            {
                return false;
            }

            return left.Value == right.Value;
        }

        public static bool operator !=(TableCell left, TableCell right)
        {
            if (left != (object)right)
            {
                return false;
            }

            return left.Value != right.Value;
        }

        public static explicit operator string(TableCell cell)
        {
            if (cell == null)
            {
                return null;
            }

            return cell.Value;
        }

        public static explicit operator bool(TableCell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException("TableCell");
            }

            return cell.Value == "1";
        }

        public static explicit operator bool?(TableCell cell)
        {
            if (cell == null)
            {
                return null;
            }

            return cell.Value == "1";
        }

        public static explicit operator int(TableCell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException("TableCell");
            }

            return int.Parse(cell.Value);
        }

        public static explicit operator int?(TableCell cell)
        {
            if (cell == null)
            {
                return null;
            }

            return int.Parse(cell.Value);
        }

        public static explicit operator uint(TableCell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException("TableCell");
            }

            return uint.Parse(cell.Value);
        }

        public static explicit operator uint?(TableCell cell)
        {
            if (cell == null)
            {
                return null;
            }

            return uint.Parse(cell.Value);
        }

        public static explicit operator long(TableCell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException("TableCell");
            }

            return long.Parse(cell.Value);
        }

        public static explicit operator long?(TableCell cell)
        {
            if (cell == null)
            {
                return null;
            }

            return long.Parse(cell.Value);
        }

        public static explicit operator ulong(TableCell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException("TableCell");
            }

            return ulong.Parse(cell.Value);
        }

        public static explicit operator ulong?(TableCell cell)
        {
            if (cell == null)
            {
                return null;
            }

            return ulong.Parse(cell.Value);
        }

        public static explicit operator float(TableCell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException("TableCell");
            }

            return float.Parse(cell.Value);
        }

        public static explicit operator float?(TableCell cell)
        {
            if (cell == null)
            {
                return null;
            }

            return float.Parse(cell.Value);
        }

        public static explicit operator double(TableCell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException("TableCell");
            }

            return double.Parse(cell.Value);
        }

        public static explicit operator double?(TableCell cell)
        {
            if (cell == null)
            {
                return null;
            }

            return double.Parse(cell.Value);
        }

        public static explicit operator decimal(TableCell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException("TableCell");
            }

            return decimal.Parse(cell.Value);
        }

        public static explicit operator decimal?(TableCell cell)
        {
            if (cell == null)
            {
                return null;
            }

            return decimal.Parse(cell.Value);
        }

        public static implicit operator DateTime(TableCell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException("TableCell");
            }

            return new DateTime(1900, 1, 1).AddDays(int.Parse(cell.Value) - 2);
        }

        public static implicit operator DateTime?(TableCell cell)
        {
            if (cell == null)
            {
                return null;
            }

            return new DateTime(1900, 1, 1).AddDays(int.Parse(cell.Value) - 2);
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
            var doc = (SpreadsheetDocument)Parent.OpenXmlPackage;
            var sharedStringTable = doc.WorkbookPart.SharedStringTablePart;
            var cells = RowElement.Elements(S.c);
            var r = cells
                .Select(cell =>
                {
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

        public Row(WorksheetPart parent)
        {
            Parent = parent;
        }
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

        public Cell(Row parent)
        {
            Parent = parent;
        }
    }

    public static class XlsxTables
    {
        public static IEnumerable<Table> Tables(this SpreadsheetDocument spreadsheet)
        {
            foreach (var worksheetPart in spreadsheet.WorkbookPart.WorksheetParts)
            {
                foreach (var table in worksheetPart.TableDefinitionParts)
                {
                    var tableDefDoc = table.GetXDocument();

                    var t = new Table(worksheetPart)
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
                    ParseRange(t.Ref, out var leftColumn, out var topRow, out var rightColumn, out var bottomRow);
                    t.LeftColumn = leftColumn;
                    t.TopRow = topRow;
                    t.RightColumn = rightColumn;
                    t.BottomRow = bottomRow;
                    yield return t;
                }
            }
        }

        public static void ParseRange(string theRef, out int leftColumn, out int topRow, out int rightColumn, out int bottomRow)
        {
            // C5:E7
            var spl = theRef.Split(':');
            var refStart = spl.First();
            var refStartSplit = SplitAddress(refStart);
            leftColumn = ColumnAddressToIndex(refStartSplit[0]);
            topRow = int.Parse(refStartSplit[1]);

            var refEnd = spl.ElementAt(1);
            var refEndSplit = SplitAddress(refEnd);
            rightColumn = ColumnAddressToIndex(refEndSplit[0]);
            bottomRow = int.Parse(refEndSplit[1]);
        }

        public static Table Table(this SpreadsheetDocument spreadsheet,
            string tableName)
        {
            return spreadsheet.Tables().Where(t => t.TableName.ToLower() == tableName.ToLower()).FirstOrDefault();
        }

        public static IEnumerable<Row> Rows(this WorksheetPart worksheetPart)
        {
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
            {
                if (address[i] >= '0' && address[i] <= '9')
                {
                    break;
                }
            }

            if (i == address.Length)
            {
                throw new FileFormatException("Invalid spreadsheet.  Bad cell address.");
            }

            return new[] {
                address.Substring(0, i),
                address.Substring(i)
            };
        }

        public static string IndexToColumnAddress(int index)
        {
            if (index < 26)
            {
                var c = (char)('A' + index);
                var s = new string(c, 1);
                return s;
            }
            if (index < 702)
            {
                var i = index - 26;
                var i1 = i / 26;
                var i2 = i % 26;
                var s = new string((char)('A' + i1), 1) +
                    new string((char)('A' + i2), 1);
                return s;
            }
            if (index < 18278)
            {
                var i = index - 702;
                var i1 = i / 676;
                i -= i1 * 676;
                var i2 = i / 26;
                var i3 = i % 26;
                var s = new string((char)('A' + i1), 1) +
                    new string((char)('A' + i2), 1) +
                    new string((char)('A' + i3), 1);
                return s;
            }
            throw new Exception("Invalid column address");
        }

        public static int ColumnAddressToIndex(string columnAddress)
        {
            if (columnAddress.Length == 1)
            {
                var c = columnAddress[0];
                var i = c - 'A';
                return i;
            }
            if (columnAddress.Length == 2)
            {
                var c1 = columnAddress[0];
                var c2 = columnAddress[1];
                var i1 = c1 - 'A';
                var i2 = c2 - 'A';
                return (i1 + 1) * 26 + i2;
            }
            if (columnAddress.Length == 3)
            {
                var c1 = columnAddress[0];
                var c2 = columnAddress[1];
                var c3 = columnAddress[2];
                var i1 = c1 - 'A';
                var i2 = c2 - 'A';
                var i3 = c3 - 'A';
                return (i1 + 1) * 676 + (i2 + 1) * 26 + i3;
            }
            throw new FileFormatException("Invalid spreadsheet: Invalid column address.");
        }
    }
}