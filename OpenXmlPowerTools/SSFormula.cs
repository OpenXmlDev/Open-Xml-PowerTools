// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text;

namespace ExcelFormula
{
    public class ParseFormula
    {
        private readonly ExcelFormula parser;

        public ParseFormula(string formula)
        {
            parser = new ExcelFormula(formula, Console.Out);
            var parserResult = false;
            try
            {
                parserResult = parser.Formula();
            }
            catch (Peg.Base.PegException)
            {
            }
            if (!parserResult)
            {
                parser.Warning("Error processing " + formula);
            }
        }

        public string ReplaceSheetName(string oldName, string newName)
        {
            var text = new StringBuilder(parser.GetSource());
            ReplaceNode(parser.GetRoot(), (int)EExcelFormula.SheetName, oldName, newName, text);
            return text.ToString();
        }

        public string ReplaceRelativeCell(int rowOffset, int colOffset)
        {
            var text = new StringBuilder(parser.GetSource());
            ReplaceRelativeCell(parser.GetRoot(), rowOffset, colOffset, text);
            return text.ToString();
        }

        // Recursive function that will replace values from last to first
        private void ReplaceNode(Peg.Base.PegNode node, int id, string oldName, string newName, StringBuilder text)
        {
            if (node.next_ != null)
            {
                ReplaceNode(node.next_, id, oldName, newName, text);
            }

            if (node.id_ == id && parser.GetSource().Substring(node.match_.posBeg_, node.match_.Length) == oldName)
            {
                text.Remove(node.match_.posBeg_, node.match_.Length);
                text.Insert(node.match_.posBeg_, newName);
            }
            else if (node.child_ != null)
            {
                ReplaceNode(node.child_, id, oldName, newName, text);
            }
        }

        // Recursive function that will adjust relative cells from last to first
        private void ReplaceRelativeCell(Peg.Base.PegNode node, int rowOffset, int colOffset, StringBuilder text)
        {
            if (node.next_ != null)
            {
                ReplaceRelativeCell(node.next_, rowOffset, colOffset, text);
            }

            if (node.id_ == (int)EExcelFormula.A1Row && parser.GetSource().Substring(node.match_.posBeg_, 1) != "$")
            {
                var rowNumber = Convert.ToInt32(parser.GetSource().Substring(node.match_.posBeg_, node.match_.Length));
                text.Remove(node.match_.posBeg_, node.match_.Length);
                text.Insert(node.match_.posBeg_, Convert.ToString(rowNumber + rowOffset));
            }
            else if (node.id_ == (int)EExcelFormula.A1Column && parser.GetSource().Substring(node.match_.posBeg_, 1) != "$")
            {
                var colNumber = GetColumnNumber(parser.GetSource().Substring(node.match_.posBeg_, node.match_.Length));
                text.Remove(node.match_.posBeg_, node.match_.Length);
                text.Insert(node.match_.posBeg_, GetColumnId(colNumber + colOffset));
            }
            else if (node.child_ != null)
            {
                ReplaceRelativeCell(node.child_, rowOffset, colOffset, text);
            }
        }

        // Converts the column reference string to a column number (e.g. A -> 1, B -> 2)
        private static int GetColumnNumber(string cellReference)
        {
            var columnNumber = 0;
            foreach (var c in cellReference)
            {
                if (char.IsLetter(c))
                {
                    columnNumber = columnNumber * 26 + System.Convert.ToInt32(c) - System.Convert.ToInt32('A') + 1;
                }
            }
            return columnNumber;
        }

        // Translates the column number to the column reference string (e.g. 1 -> A, 2-> B)
        private static string GetColumnId(int columnNumber)
        {
            var result = "";
            do
            {
                result = ((char)((columnNumber - 1) % 26 + 'A')).ToString() + result;
                columnNumber = (columnNumber - 1) / 26;
            } while (columnNumber != 0);
            return result;
        }
    }
}
