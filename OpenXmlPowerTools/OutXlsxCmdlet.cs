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

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using SSW = OpenXmlPowerTools;


namespace OpenXmlPowerTools.Commands
{
    [Cmdlet(VerbsData.Out, "Xlsx", SupportsShouldProcess = true)]
    [OutputType("OpenXmlPowerToolsDocument")]
    public class OutXlsxCmdlet : PowerToolsCreateCmdlet
    {
        #region Parameters
        private const string DefaultSheetName = "Sheet1";
        private PSObject[] pipeObjects;
        private Collection<PSObject> processedObjects = new Collection<PSObject>();
        string fileName;
        bool openWithExcel;
        string sheetName;
        string tableName;

        /// <summary>
        /// FileName parameter
        /// </summary>
        [Parameter(
            Position = 0,
            Mandatory = true,
            ValueFromPipeline = false,
            HelpMessage = "Path of file in which to store results")
        ]
        public string FileName
        {
            get
            {
                return fileName;
            }
            set
            {
                fileName = Path.Combine(SessionState.Path.CurrentLocation.Path, value);
            }
        }

        /// <summary>
        /// InputObject parameter
        /// </summary>
        [Parameter(
            Mandatory = false,
            ValueFromPipeline = true,
            HelpMessage = "Objects passed by pipe to be included in spreadsheet")
        ]
        public PSObject[] InputObject
        {
            get
            {
                return pipeObjects;
            }
            set
            {
                pipeObjects = value;
            }
        }


        /// <summary>
        /// SheetName parameter
        /// </summary>
        [Parameter(
            Mandatory = false,
            ValueFromPipeline = false,
            HelpMessage = "Specify the sheet name")
        ]
        public string SheetName
        {
            get
            {
                return sheetName;
            }
            set
            {
                sheetName = value;
            }
        }


        /// <summary>
        /// TableName parameter
        /// </summary>
        [Parameter(
            Mandatory = false,
            ValueFromPipeline = false,
            HelpMessage = "Specify the table name")
        ]
        public string TableName
        {
            get
            {
                return tableName;
            }
            set
            {
                tableName = value;
            }
        }


        /// <summary>
        /// OpenWithExcel parameter
        /// </summary>
        [Parameter(
            Mandatory = false,
            ValueFromPipeline = false,
            HelpMessage = "Open with Excel")
        ]
        public SwitchParameter OpenWithExcel
        {
            get
            {
                return openWithExcel;
            }
            set
            {
                openWithExcel = value;
            }
        }

        #endregion

        # region Fields
        List<Type> typeCollection = new List<Type>();
        List<SSW.CellDfn> columnHeadings = new List<SSW.CellDfn>();
        Types type = Types.none;
        string[][] result;
        #endregion

        #region Cmdlet Overrides

        protected override void ProcessRecord()
        {
            if (pipeObjects != null)
            {
                foreach (PSObject pipeObject in pipeObjects)
                    processedObjects.Add(pipeObject);
            }
        }

        protected override void EndProcessing()
        {
            ValidateSheetName(fileName, sheetName);
            if (tableName != null)
            {
                ValidateTableName(tableName);
            }

            FileInfo fi = new FileInfo(fileName);
            if (fi.Extension.ToLower() != ".xlsx")
                fi = new FileInfo(Path.Combine(fi.DirectoryName, fi.Name.Substring(0, fi.Name.Length - fi.Extension.Length) + ".xlsx"));
            fileName = fi.FullName;

            if (!File.Exists(fileName) || ShouldProcess(fileName, "Out-Xlsx"))
            {
                if (!File.Exists(fileName))
                {
                    using (OpenXmlMemoryStreamDocument streamDoc = OpenXmlMemoryStreamDocument.CreateSpreadsheetDocument())
                    {
                        using (SpreadsheetDocument document = streamDoc.GetSpreadsheetDocument())
                        {
                            if (processedObjects.Count > 0)
                            {
                                GetPSObjectType();
                                CreateColumns();
                                FillValues();
                            }
                            GenerateNewSpreadSheetDocument(document, sheetName, tableName);
                        }
                    }
                }
                else
                    OverWriteSpreadSheetDocument(fileName);

                if (openWithExcel)
                {
                    FileInfo file = new FileInfo(fileName);
                    if (file.Exists)
                        System.Diagnostics.Process.Start(fileName);
                }
            }

        }

        private void GenerateNewSpreadSheetDocument(SpreadsheetDocument document, string sheetName, string tableName)
        {
            List<SSW.RowDfn> rowCollection = GetRowCollection(result);

            // Create new work sheet to te document
            List<SSW.WorksheetDfn> workSheetCollection = new List<SSW.WorksheetDfn>();
            SSW.WorksheetDfn workSheet = new SSW.WorksheetDfn();

            if (string.IsNullOrEmpty(sheetName))
                workSheet.Name = DefaultSheetName;
            else
                workSheet.Name = this.sheetName;

            workSheet.TableName = tableName;
            workSheet.ColumnHeadings = columnHeadings;
            workSheet.Rows = rowCollection;

            workSheetCollection.Add(workSheet);

            // Create work book
            SSW.WorkbookDfn workBook = new SSW.WorkbookDfn();
            workBook.Worksheets = workSheetCollection;

            // Create Excel File
            SSW.SpreadsheetWriter.Write(fileName, workBook);

        }

        private void OverWriteSpreadSheetDocument(string fileName)
        {
            using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(fileName, true))
            {
                if (processedObjects.Count > 0)
                {
                    GetPSObjectType();
                    CreateColumns();
                    FillValues();

                    List<SSW.RowDfn> rowCollection = GetRowCollection(result);

                    //Add work sheet to existing document
                    SSW.WorksheetDfn workSheet = new SSW.WorksheetDfn();
                    workSheet.Name = this.sheetName;
                    workSheet.TableName = tableName;
                    workSheet.ColumnHeadings = columnHeadings;
                    workSheet.Rows = rowCollection;

                    SSW.SpreadsheetWriter.AddWorksheet(sDoc, workSheet);
                }
            }
        }

        private List<SSW.RowDfn> GetRowCollection(string[][] result)
        {
            List<SSW.RowDfn> rowCollection = new List<SSW.RowDfn>();
            for (int i = 0; i < result.Count(); i++)
            {
                string[] dataRow = result[i];
                SSW.RowDfn row = new SSW.RowDfn();
                List<SSW.CellDfn> cellCollection = new List<SSW.CellDfn>();

                for (int j = 0; j < dataRow.Count(); j++)
                {
                    SSW.CellDfn dataCell = new SSW.CellDfn();
                    dataCell.Value = dataRow[j];
                    dataCell.HorizontalCellAlignment = SSW.HorizontalCellAlignment.Left;
                    dataCell.CellDataType = CellDataType.String;
                    cellCollection.Add(dataCell);
                }

                row.Cells = cellCollection;
                rowCollection.Add(row);
            }

            return rowCollection;
        }

        private void CreateColumns()
        {
            switch (type)
            {
                case Types.ReferenceType:
                    {
                        foreach (PSObject obj in processedObjects)
                        {
                            if (!typeCollection.Contains(obj.BaseObject.GetType()))
                            {
                                typeCollection.Add(obj.BaseObject.GetType());
                                CreateColumnHeadings(obj);
                            }
                        }
                        break;
                    }

                case Types.ScalarType:
                    {
                        SSW.CellDfn columnHeading = new SSW.CellDfn();
                        columnHeading.Value = processedObjects.First().BaseObject.GetType().FullName;
                        columnHeadings.Add(columnHeading);
                        break;
                    }
                case Types.ScalarTypes:
                    {
                        SSW.CellDfn indexColumnHeading = new SSW.CellDfn();
                        indexColumnHeading.Value = "Index";
                        columnHeadings.Add(indexColumnHeading);

                        SSW.CellDfn valueColumnHeading = new SSW.CellDfn();
                        valueColumnHeading.Value = "Value";
                        columnHeadings.Add(valueColumnHeading);

                        SSW.CellDfn typeColumnHeading = new SSW.CellDfn();
                        typeColumnHeading.Value = "Type";
                        columnHeadings.Add(typeColumnHeading);

                        break;
                    }
                default:
                    {
                        break;
                    }
            }
        }

        private void CreateColumnHeadings(PSObject obj)
        {
            foreach (PSPropertyInfo property in obj.Properties)
            {
                if (columnHeadings.Count == 0)
                {
                    SSW.CellDfn columnHeading = new SSW.CellDfn();
                    columnHeading.Value = property.Name;
                    columnHeadings.Add(columnHeading);
                }
                else if (!columnHeadings.Exists(e => e.Value.Equals(property.Name)))
                {
                    SSW.CellDfn columnHeading = new SSW.CellDfn();
                    columnHeading.Value = property.Name;
                    columnHeadings.Add(columnHeading);
                }
            }
        }

        private void FillValues()
        {
            result = new string[processedObjects.Count][];

            switch (type)
            {
                case Types.ReferenceType:
                    {
                        for (int i = 0; i < processedObjects.Count; i++)
                        {
                            PSObject obj = processedObjects[i];
                            result[i] = new string[columnHeadings.Count];
                            for (int j = 0; j < columnHeadings.Count; j++)
                            {
                                string propertyName = Convert.ToString(columnHeadings[j].Value);
                                try
                                {
                                    if (obj.Properties[propertyName] != null)
                                    {
                                        string value = Convert.ToString(obj.Properties[propertyName].Value);
                                        if (!string.IsNullOrEmpty(value))
                                        {
                                            result[i][j] = value;
                                        }
                                    }
                                }
                                catch (GetValueInvocationException e)
                                {
                                    WriteDebug(string.Format(CultureInfo.InvariantCulture, "Exception ({0}) ", e.Message));
                                }
                            }
                        }
                        break;
                    }

                case Types.ScalarType:
                    {
                        for (int i = 0; i < processedObjects.Count; i++)
                        {
                            result[i] = new string[1];
                            result[i][0] = Convert.ToString(processedObjects[i]);
                        }
                        break;
                    }

                case Types.ScalarTypes:
                    {

                        for (int i = 0; i < processedObjects.Count; i++)
                        {
                            result[i] = new string[3];
                            result[i][0] = Convert.ToString(i);
                            result[i][1] = Convert.ToString(processedObjects[i]);
                            result[i][2] = processedObjects[i].BaseObject.GetType().FullName;
                        }

                        break;
                    }
                default:
                    {
                        break;
                    }
            }
        }

        private void GetPSObjectType()
        {
            if (!DefaultScalarTypes.IsTypeInList(processedObjects[0].TypeNames))
            {
                type = Types.ReferenceType;
                return;
            }

            string firstType = processedObjects[0].BaseObject.GetType().FullName;

            for (int i = 1; i < processedObjects.Count && processedObjects.Count > 1; i++)
            {
                if (firstType != processedObjects[i].BaseObject.GetType().FullName)
                {
                    type = Types.ScalarTypes;
                    return;
                }
            }

            type = Types.ScalarType;
        }       

        private void ValidateSheetName(string fileName, string sheetName)
        {
            if (File.Exists(fileName))
            {
                if (string.IsNullOrEmpty(sheetName))
                    throw new Exception("Sheet name is missing. Specify a sheet name.");
            }
        }

        public static void ValidateTableName(string tableName)
        {
            if (tableName.Contains(' '))
                throw new Exception("Table name contains a space.");
            if (tableName.Length > 255)
                throw new Exception("Table name length exceeds 255.");
            char first = tableName[0];
            char last = tableName[tableName.Length - 1];
            if (!(char.IsLetter(first) || first == '_'))
                throw new Exception("Table name does not start with a letter or underscore.");
            if (char.IsLetter(last) || last == '_')
                return;
            if (char.IsDigit(last))
            {
                char[] anyOf = new[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', };
                int firstDigit = tableName.IndexOfAny(anyOf);
                if (firstDigit <= 3)
                    throw new Exception("Invalid table name.");
            }
        }

        #endregion
    }

    public enum Types
    {
        ReferenceType,
        ScalarType,
        ScalarTypes,
        none
    }
}