// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    internal class ComparisonUnitWord : ComparisonUnit
    {
        public static readonly XName[] ElementsWithRelationshipIds =
        {
            A.blip,
            A.hlinkClick,
            A.relIds,
            C.chart,
            C.externalData,
            C.userShapes,
            DGM.relIds,
            O.OLEObject,
            VML.fill,
            VML.imagedata,
            VML.stroke,
            W.altChunk,
            W.attachedTemplate,
            W.control,
            W.dataSource,
            W.embedBold,
            W.embedBoldItalic,
            W.embedItalic,
            W.embedRegular,
            W.footerReference,
            W.headerReference,
            W.headerSource,
            W.hyperlink,
            W.printerSettings,
            W.recipientData,
            W.saveThroughXslt,
            W.sourceFileName,
            W.src,
            W.subDoc,
            WNE.toolbarData
        };

        public static readonly XName[] RelationshipAttributeNames =
        {
            R.embed,
            R.link,
            R.id,
            R.cs,
            R.dm,
            R.lo,
            R.qs,
            R.href,
            R.pict
        };

        public ComparisonUnitWord(IEnumerable<ComparisonUnitAtom> comparisonUnitAtomList)
        {
            Contents = comparisonUnitAtomList.OfType<ComparisonUnit>().ToList();
            string sha1String = Contents.Select(c => c.SHA1Hash).StringConcatenate();
            SHA1Hash = WmlComparerUtil.SHA1HashStringForUTF8String(sha1String);
        }

        public override string ToString(int indent)
        {
            var sb = new StringBuilder();
            sb.Append("".PadRight(indent) + "Word SHA1:" + SHA1Hash.Substring(0, 8) + Environment.NewLine);

            foreach (ComparisonUnit comparisonUnitAtom in Contents)
            {
                sb.Append(comparisonUnitAtom.ToString(indent + 2) + Environment.NewLine);
            }

            return sb.ToString();
        }
    }
}
