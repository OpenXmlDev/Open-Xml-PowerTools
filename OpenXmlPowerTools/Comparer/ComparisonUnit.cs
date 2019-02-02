// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenXmlPowerTools
{
    public abstract class ComparisonUnit
    {
        private int? _descendantContentAtomsCount;

        public CorrelationStatus CorrelationStatus { get; set; }

        public List<ComparisonUnit> Contents { get; protected set; }

        public string SHA1Hash { get; protected set; }

        public int DescendantContentAtomsCount
        {
            get
            {
                if (_descendantContentAtomsCount != null) return (int) _descendantContentAtomsCount;

                _descendantContentAtomsCount = DescendantContentAtoms().Count();
                return (int) _descendantContentAtomsCount;
            }
        }

        private IEnumerable<ComparisonUnit> Descendants()
        {
            var comparisonUnitList = new List<ComparisonUnit>();
            DescendantsInternal(this, comparisonUnitList);
            return comparisonUnitList;
        }

        public IEnumerable<ComparisonUnitAtom> DescendantContentAtoms()
        {
            return Descendants().OfType<ComparisonUnitAtom>();
        }

        private static void DescendantsInternal(
            ComparisonUnit comparisonUnit,
            List<ComparisonUnit> comparisonUnitList)
        {
            foreach (ComparisonUnit cu in comparisonUnit.Contents)
            {
                comparisonUnitList.Add(cu);
                if (cu.Contents != null && cu.Contents.Any())
                    DescendantsInternal(cu, comparisonUnitList);
            }
        }

        public abstract string ToString(int indent);

        internal static string ComparisonUnitListToString(ComparisonUnit[] cul)
        {
            var sb = new StringBuilder();
            sb.Append("Dump Comparision Unit List To String" + Environment.NewLine);
            foreach (ComparisonUnit item in cul) sb.Append(item.ToString(2) + Environment.NewLine);

            return sb.ToString();
        }
    }
}
