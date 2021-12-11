﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace ExampleFormulas
{
    internal class ExampleFormulas
    {
        private static void Main(string[] args)
        {
            DateTime n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000,
                n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            // Change sheet name in formulas
            using (var streamDoc = new OpenXmlMemoryStreamDocument(
                OpenXmlPowerToolsDocument.FromFileName("../../../Formulas.xlsx")))
            {
                using (SpreadsheetDocument doc = streamDoc.GetSpreadsheetDocument())
                {
                    WorksheetAccessor.FormulaReplaceSheetName(doc, "Source", "'Source 2'");
                }

                streamDoc.GetModifiedSmlDocument().SaveAs(Path.Combine(tempDi.FullName, "FormulasUpdated.xlsx"));
            }

            // Change sheet name in formulas
            using (var streamDoc = new OpenXmlMemoryStreamDocument(
                OpenXmlPowerToolsDocument.FromFileName("../../../Formulas.xlsx")))
            {
                using (SpreadsheetDocument doc = streamDoc.GetSpreadsheetDocument())
                {
                    WorksheetPart sheet = WorksheetAccessor.GetWorksheet(doc, "References");
                    WorksheetAccessor.CopyCellRange(doc, sheet, 1, 1, 7, 5, 4, 8);
                }

                streamDoc.GetModifiedSmlDocument().SaveAs(Path.Combine(tempDi.FullName, "FormulasCopied.xlsx"));
            }
        }
    }
}
