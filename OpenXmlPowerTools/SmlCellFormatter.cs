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
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public class SmlCellFormatter
    {
        private enum CellType
        {
            General,
            Number,
            Date,
        };

        private class FormatConfig
        {
            public CellType CellType;
            public string FormatCode;
        }

        private static Dictionary<string, FormatConfig> ExcelFormatCodeToNetFormatCodeExceptionMap = new Dictionary<string, FormatConfig>()
        {
            {
                "# ?/?",
                new FormatConfig
                {
                    CellType = SmlCellFormatter.CellType.Number,
                    FormatCode = "0.00",
                }
            },
            {
                "# ??/??",
                new FormatConfig
                {
                    CellType = SmlCellFormatter.CellType.Number,
                    FormatCode = "0.00",
                }
            },
        };

        // Up to four sections of format codes can be specified. The format codes, separated by semicolons, define the
        // formats for positive numbers, negative numbers, zero values, and text, in that order. If only two sections are
        // specified, the first is used for positive numbers and zeros, and the second is used for negative numbers. If only
        // one section is specified, it is used for all numbers. To skip a section, the ending semicolon for that section shall
        // be written.
        public static string FormatCell(string formatCode, string value, out string color)
        {
            color = null;

            if (formatCode == null)
                formatCode = "General";

            var splitFormatCode = formatCode.Split(';');
            if (splitFormatCode.Length == 1)
            {
                double dv;
                if (double.TryParse(value, out dv))
                {
                    return FormatDouble(formatCode, dv, out color);
                }
                return value;
            }
            if (splitFormatCode.Length == 2)
            {
                double dv;
                if (double.TryParse(value, out dv))
                {
                    if (dv > 0)
                    {
                        return FormatDouble(splitFormatCode[0], dv, out color);
                    }
                    else
                    {
                        return FormatDouble(splitFormatCode[1], dv, out color);
                    }
                }
                return value;

            }
            // positive, negative, zero, text
            // _("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)
            if (splitFormatCode.Length == 4)
            {
                double dv;
                if (double.TryParse(value, out dv))
                {
                    if (dv > 0)
                    {
                        var z1 = FormatDouble(splitFormatCode[0], dv, out color);
                        return z1;
                    }
                    else if (dv < 0)
                    {
                        var z2 = FormatDouble(splitFormatCode[1], dv, out color);
                        return z2;
                    }
                    else // == 0
                    {
                        var z3 = FormatDouble(splitFormatCode[2], dv, out color);
                        return z3;
                    }
                }
                string fmt = splitFormatCode[3].Replace("@", "{0}").Replace("\"", "");
                try
                {
                    var s = string.Format(fmt, value);
                    return s;
                }
                catch (Exception)
                {
                    return value;
                }
            }
            return value;
        }

        static Regex UnderRegex = new Regex("_.");

        // The following Regex transforms currency specifies into a character / string
        // that string.Format can use to properly produce the correct text.
        // "[$£-809]"    => "£"
        // "[$€-2]"      => "€"
        // "[$¥-804]"    => "¥
        // "[$CHF-100C]" => "CHF"
        static string s_CurrRegex = @"\[\$(?<curr>.*-).*\]";

        private static string ConvertFormatCode(string formatCode)
        {
            var newFormatCode = formatCode
                .Replace("mmm-", "MMM-")
                .Replace("-mmm", "-MMM")
                .Replace("mm-", "MM-")
                .Replace("mmmm", "MMMM")
                .Replace("AM/PM", "tt")
                .Replace("m/", "M/")
                .Replace("*", "")
                .Replace("?", "#")
                ;
            var withTrimmedUnderscores = UnderRegex.Replace(newFormatCode, "");
            var withTransformedCurrency = Regex.Replace(withTrimmedUnderscores, s_CurrRegex, m => m.Groups[1].Value.TrimEnd('-'));
            return withTransformedCurrency;
        }

        private static string[] ValidColors = new[] {
            "Black",
            "Blue",
            "Cyan",
            "Green",
            "Magenta",
            "Red",
            "White",
            "Yellow",
        };

        private static string FormatDouble(string formatCode, double dv, out string color)
        {
            color = null;
            var trimmed = formatCode.Trim();
            if (trimmed.StartsWith("[") &&
                trimmed.Contains("]"))
            {
                var colorLen = trimmed.IndexOf(']');
                color = trimmed.Substring(1, colorLen - 1);
                if (ValidColors.Contains(color) ||
                    color.StartsWith("Color"))
                {
                    if (color.StartsWith("Color"))
                    {
                        var idxStr = color.Substring(5);
                        int colorIdx;
                        if (int.TryParse(idxStr, out colorIdx))
                        {
                            if (colorIdx < SmlDataRetriever.IndexedColors.Length)
                                color = SmlDataRetriever.IndexedColors[colorIdx];
                            else
                                color = null;
                        }
                    }
                    formatCode = trimmed.Substring(colorLen + 1);
                }
                else
                    color = null;
            }


            if (formatCode == "General")
                return dv.ToString();
            bool isDate = IsFormatCodeForDate(formatCode);
            var cfc = ConvertFormatCode(formatCode);
            if (isDate)
            {
                DateTime thisDate;
                try
                {
                    thisDate = DateTime.FromOADate(dv);
                }
                catch (ArgumentException)
                {
                    return dv.ToString();
                }
                if (cfc.StartsWith("[h]"))
                {
                    DateTime zeroHour = new DateTime(1899, 12, 30, 0, 0, 0);
                    int deltaInHours = (int)((thisDate - zeroHour).TotalHours);
                    var newCfc = cfc.Substring(3);
                    var s = (deltaInHours.ToString() + thisDate.ToString(newCfc)).Trim();
                    return s;
                }
                if (cfc.EndsWith(".0"))
                {
                    var cfc2 = cfc.Replace(".0", ":fff");
                    var s4 = thisDate.ToString(cfc2).Trim();
                    return s4;
                }
                var s2 = thisDate.ToString(cfc).Trim();
                return s2;
            }
            if (ExcelFormatCodeToNetFormatCodeExceptionMap.ContainsKey(formatCode))
            {
                FormatConfig fc = ExcelFormatCodeToNetFormatCodeExceptionMap[formatCode];
                var s = dv.ToString(fc.FormatCode).Trim();
                return s;
            }
            if ((cfc.Contains('(') && cfc.Contains(')')) || cfc.Contains('-'))
            {
                var s3 = (-dv).ToString(cfc).Trim();
                return s3;
            }
            else
            {
                var s4 = dv.ToString(cfc).Trim();
                return s4;
            }
        }

        private static bool IsFormatCodeForDate(string formatCode)
        {
            if (formatCode == "General")
                return false;
            return formatCode.Contains("m") ||
                formatCode.Contains("d") ||
                formatCode.Contains("y") ||
                formatCode.Contains("h") ||
                formatCode.Contains("s") ||
                formatCode.Contains("AM") ||
                formatCode.Contains("PM");
        }
    }
}
