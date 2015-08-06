using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using ExcelFormula;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace ExampleFormulas
{
    class ExampleFormulas
    {
        static void Main(string[] args)
        {
            // Change sheet name in formulas
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(
                SmlDocument.FromFileName("../../Formulas.xlsx")))
            {
                using (SpreadsheetDocument doc = streamDoc.GetSpreadsheetDocument())
                {
                    WorksheetAccessor.FormulaReplaceSheetName(doc, "Source", "'Source 2'");
                }
                streamDoc.GetModifiedSmlDocument().SaveAs("../../FormulasUpdated.xlsx");
            }

            // Change sheet name in formulas
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(
                SmlDocument.FromFileName("../../Formulas.xlsx")))
            {
                using (SpreadsheetDocument doc = streamDoc.GetSpreadsheetDocument())
                {
                    WorksheetPart sheet = WorksheetAccessor.GetWorksheet(doc, "References");
                    WorksheetAccessor.CopyCellRange(doc, sheet, 1, 1, 7, 5, 4, 8);
                }
                streamDoc.GetModifiedSmlDocument().SaveAs("../../FormulasCopied.xlsx");
            }
        }
    }
}
