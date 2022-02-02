// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static partial class WmlComparer
    {
#pragma warning disable 414
        private static readonly bool False = false;
        private static readonly bool True = true;
        private static readonly bool SaveIntermediateFilesForDebugging = false;
#pragma warning restore 414

        private static readonly string NewLine = Environment.NewLine;

        private static readonly XAttribute[] NamespaceAttributes =
        {
            new XAttribute(XNamespace.Xmlns + "wpc", WPC.wpc),
            new XAttribute(XNamespace.Xmlns + "mc", MC.mc),
            new XAttribute(XNamespace.Xmlns + "o", O.o),
            new XAttribute(XNamespace.Xmlns + "r", R.r),
            new XAttribute(XNamespace.Xmlns + "m", M.m),
            new XAttribute(XNamespace.Xmlns + "v", VML.vml),
            new XAttribute(XNamespace.Xmlns + "wp14", WP14.wp14),
            new XAttribute(XNamespace.Xmlns + "wp", WP.wp),
            new XAttribute(XNamespace.Xmlns + "w10", W10.w10),
            new XAttribute(XNamespace.Xmlns + "w", W.w),
            new XAttribute(XNamespace.Xmlns + "w14", W14.w14),
            new XAttribute(XNamespace.Xmlns + "wpg", WPG.wpg),
            new XAttribute(XNamespace.Xmlns + "wpi", WPI.wpi),
            new XAttribute(XNamespace.Xmlns + "wne", WNE.wne),
            new XAttribute(XNamespace.Xmlns + "wps", WPS.wps),
            new XAttribute(MC.Ignorable, "w14 wp14")
        };

        private static readonly XName[] RevElementsWithNoText =
        {
            M.oMath,
            M.oMathPara,
            W.drawing
        };

        private static readonly XName[] AttributesToTrimWhenCloning =
        {
            WP14.anchorId,
            WP14.editId,
            "ObjectID",
            "ShapeID",
            "id",
            "type"
        };

        private static int _maxId;

        private static readonly XName[] WordBreakElements =
        {
            W.pPr,
            W.tab,
            W.br,
            W.continuationSeparator,
            W.cr,
            W.dayLong,
            W.dayShort,
            W.drawing,
            W.pict,
            W.endnoteRef,
            W.footnoteRef,
            W.monthLong,
            W.monthShort,
            W.noBreakHyphen,
            W._object,
            W.ptab,
            W.separator,
            W.sym,
            W.yearLong,
            W.yearShort,
            M.oMathPara,
            M.oMath,
            W.footnoteReference,
            W.endnoteReference
        };

        private static readonly XName[] AllowableRunChildren =
        {
            W.br,
            W.drawing,
            W.cr,
            W.dayLong,
            W.dayShort,
            W.footnoteReference,
            W.endnoteReference,
            W.monthLong,
            W.monthShort,
            W.noBreakHyphen,

            //W._object,
            W.pgNum,
            W.ptab,
            W.softHyphen,
            W.sym,
            W.tab,
            W.yearLong,
            W.yearShort,
            M.oMathPara,
            M.oMath,
            W.fldChar,
            W.instrText
        };

        private static readonly XName[] ElementsToThrowAway =
        {
            W.bookmarkStart,
            W.bookmarkEnd,
            W.commentRangeStart,
            W.commentRangeEnd,
            W.lastRenderedPageBreak,
            W.proofErr,
            W.tblPr,
            W.sectPr,
            W.permEnd,
            W.permStart,
            W.footnoteRef,
            W.endnoteRef,
            W.separator,
            W.continuationSeparator
        };

        private static readonly XName[] ElementsToHaveSha1Hash =
        {
            W.p,
            W.tbl,
            W.tr,
            W.tc,
            W.drawing,
            W.pict,
            W.txbxContent
        };

        private static readonly XName[] InvalidElements =
        {
            W.altChunk,
            W.customXml,
            W.customXmlDelRangeEnd,
            W.customXmlDelRangeStart,
            W.customXmlInsRangeEnd,
            W.customXmlInsRangeStart,
            W.customXmlMoveFromRangeEnd,
            W.customXmlMoveFromRangeStart,
            W.customXmlMoveToRangeEnd,
            W.customXmlMoveToRangeStart,
            W.moveFrom,
            W.moveFromRangeStart,
            W.moveFromRangeEnd,
            W.moveTo,
            W.moveToRangeStart,
            W.moveToRangeEnd,
            W.subDoc
        };

        private static readonly RecursionInfo[] RecursionElements =
        {
            new RecursionInfo
            {
                ElementName = W.del,
                ChildElementPropertyNames = null
            },
            new RecursionInfo
            {
                ElementName = W.ins,
                ChildElementPropertyNames = null
            },
            new RecursionInfo
            {
                ElementName = W.tbl,
                ChildElementPropertyNames = new[] { W.tblPr, W.tblGrid, W.tblPrEx }
            },
            new RecursionInfo
            {
                ElementName = W.tr,
                ChildElementPropertyNames = new[] { W.trPr, W.tblPrEx }
            },
            new RecursionInfo
            {
                ElementName = W.tc,
                ChildElementPropertyNames = new[] { W.tcPr, W.tblPrEx }
            },
            new RecursionInfo
            {
                ElementName = W.pict,
                ChildElementPropertyNames = new[] { VML.shapetype }
            },
            new RecursionInfo
            {
                ElementName = VML.group,
                ChildElementPropertyNames = null
            },
            new RecursionInfo
            {
                ElementName = VML.shape,
                ChildElementPropertyNames = null
            },
            new RecursionInfo
            {
                ElementName = VML.rect,
                ChildElementPropertyNames = null
            },
            new RecursionInfo
            {
                ElementName = VML.textbox,
                ChildElementPropertyNames = null
            },
            new RecursionInfo
            {
                ElementName = O._lock,
                ChildElementPropertyNames = null
            },
            new RecursionInfo
            {
                ElementName = W.txbxContent,
                ChildElementPropertyNames = null
            },
            new RecursionInfo
            {
                ElementName = W10.wrap,
                ChildElementPropertyNames = null
            },
            new RecursionInfo
            {
                ElementName = W.sdt,
                ChildElementPropertyNames = new[] { W.sdtPr, W.sdtEndPr }
            },
            new RecursionInfo
            {
                ElementName = W.sdtContent,
                ChildElementPropertyNames = null
            },
            new RecursionInfo
            {
                ElementName = W.hyperlink,
                ChildElementPropertyNames = null
            },
            new RecursionInfo
            {
                ElementName = W.fldSimple,
                ChildElementPropertyNames = null
            },
            new RecursionInfo
            {
                ElementName = VML.shapetype,
                ChildElementPropertyNames = null
            },
            new RecursionInfo
            {
                ElementName = W.smartTag,
                ChildElementPropertyNames = new[] { W.smartTagPr }
            },
            new RecursionInfo
            {
                ElementName = W.ruby,
                ChildElementPropertyNames = new[] { W.rubyPr }
            }
        };

        private static readonly XName[] ComparisonGroupingElements =
        {
            W.p,
            W.tbl,
            W.tr,
            W.tc,
            W.txbxContent
        };
    }
}
