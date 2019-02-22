// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text;

namespace OpenXmlPowerTools
{
    internal class CorrelatedSequence
    {
#if DEBUG
        public string SourceFile;
        public int SourceLine;
#endif

        public CorrelatedSequence()
        {
#if DEBUG
            SourceFile = new System.Diagnostics.StackTrace(true).GetFrame(1).GetFileName();
            SourceLine = new System.Diagnostics.StackTrace(true).GetFrame(1).GetFileLineNumber();
#endif
        }

        public CorrelationStatus CorrelationStatus { get; set; }

        // if ComparisonUnitList1 == null and ComparisonUnitList2 contains sequence, then inserted content.
        // if ComparisonUnitList2 == null and ComparisonUnitList1 contains sequence, then deleted content.
        // if ComparisonUnitList2 contains sequence and ComparisonUnitList1 contains sequence, then either is Unknown or Equal.
        public ComparisonUnit[] ComparisonUnitArray1 { get; set; }

        public ComparisonUnit[] ComparisonUnitArray2 { get; set; }

        public override string ToString()
        {
            var sb = new StringBuilder();
            const string indentString = "  ";
            const string indentString4 = "    ";
            sb.Append("CorrelatedSequence =====" + Environment.NewLine);
#if DEBUG
            sb.Append(indentString + "Created at Line: " + SourceLine + Environment.NewLine);
#endif
            sb.Append(indentString + "CorrelatedItem =====" + Environment.NewLine);
            sb.Append(indentString4 + "CorrelationStatus: " + CorrelationStatus + Environment.NewLine);
            if (CorrelationStatus == CorrelationStatus.Equal)
            {
                sb.Append(indentString4 + "ComparisonUnitList =====" + Environment.NewLine);
                foreach (ComparisonUnit item in ComparisonUnitArray2)
                    sb.Append(item.ToString(6) + Environment.NewLine);
            }
            else
            {
                if (ComparisonUnitArray1 != null)
                {
                    sb.Append(indentString4 + "ComparisonUnitList1 =====" + Environment.NewLine);
                    foreach (ComparisonUnit item in ComparisonUnitArray1)
                        sb.Append(item.ToString(6) + Environment.NewLine);
                }

                if (ComparisonUnitArray2 != null)
                {
                    sb.Append(indentString4 + "ComparisonUnitList2 =====" + Environment.NewLine);
                    foreach (ComparisonUnit item in ComparisonUnitArray2)
                        sb.Append(item.ToString(6) + Environment.NewLine);
                }
            }

            return sb.ToString();
        }
    }
}
