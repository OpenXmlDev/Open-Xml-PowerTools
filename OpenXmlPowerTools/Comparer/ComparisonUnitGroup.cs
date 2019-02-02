// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    internal class ComparisonUnitGroup : ComparisonUnit
    {
        public ComparisonUnitGroup(
            IEnumerable<ComparisonUnit> comparisonUnitList,
            ComparisonUnitGroupType groupType,
            int level)
        {
            Contents = comparisonUnitList.ToList();
            ComparisonUnitGroupType = groupType;
            ComparisonUnit first = Contents.First();
            ComparisonUnitAtom comparisonUnitAtom = GetFirstComparisonUnitAtomOfGroup(first);

            XElement[] ancestorsToLookAt = comparisonUnitAtom
                .AncestorElements
                .Where(e => e.Name == W.tbl || e.Name == W.tr || e.Name == W.tc || e.Name == W.p || e.Name == W.txbxContent)
                .ToArray();

            XElement ancestor = ancestorsToLookAt[level];
            if (ancestor == null) throw new OpenXmlPowerToolsException("Internal error: ComparisonUnitGroup");

            SHA1Hash = (string) ancestor.Attribute(PtOpenXml.SHA1Hash);
            CorrelatedSHA1Hash = (string) ancestor.Attribute(PtOpenXml.CorrelatedSHA1Hash);
            StructureSHA1Hash = (string) ancestor.Attribute(PtOpenXml.StructureSHA1Hash);
        }

        public ComparisonUnitGroupType ComparisonUnitGroupType { get; }

        public string CorrelatedSHA1Hash { get; }

        public string StructureSHA1Hash { get; }

        private static ComparisonUnitAtom GetFirstComparisonUnitAtomOfGroup(ComparisonUnit group)
        {
            ComparisonUnit thisGroup = group;
            while (true)
            {
                if (thisGroup is ComparisonUnitGroup tg)
                {
                    thisGroup = tg.Contents.First();
                    continue;
                }

                if (!(thisGroup is ComparisonUnitWord tw))
                {
                    throw new OpenXmlPowerToolsException("Internal error: GetFirstComparisonUnitAtomOfGroup");
                }

                var ca = (ComparisonUnitAtom) tw.Contents.First();
                return ca;
            }
        }

        public override string ToString(int indent)
        {
            var sb = new StringBuilder();
            sb.Append("".PadRight(indent) + "Group Type: " + ComparisonUnitGroupType + " SHA1:" + SHA1Hash + Environment.NewLine);

            foreach (ComparisonUnit comparisonUnitAtom in Contents)
            {
                sb.Append(comparisonUnitAtom.ToString(indent + 2));
            }

            return sb.ToString();
        }
    }
}
