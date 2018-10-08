using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public static class WordprocessingMLUtil
    {
        private static readonly HashSet<string> UnknownFonts = new HashSet<string>();
        private static HashSet<string> KnownFamilies;

        public static int CalcWidthOfRunInTwips(XElement r)
        {
            if (KnownFamilies == null)
            {
                KnownFamilies = new HashSet<string>();
                FontFamily[] families = FontFamily.Families;
                foreach (FontFamily fam in families)
                    KnownFamilies.Add(fam.Name);
            }

            var fontName = (string)r.Attribute(PtOpenXml.pt + "FontName");
            if (fontName == null)
                fontName = (string)r.Ancestors(W.p).First().Attribute(PtOpenXml.pt + "FontName");
            if (fontName == null)
                throw new OpenXmlPowerToolsException("Internal Error, should have FontName attribute");

            if (UnknownFonts.Contains(fontName))
                return 0;

            XElement rPr = r.Element(W.rPr);
            if (rPr == null)
                throw new OpenXmlPowerToolsException("Internal Error, should have run properties");

            var languageType = (string)r.Attribute(PtOpenXml.LanguageType);
            decimal? szn = null;
            if (languageType == "bidi")
                szn = (decimal?)rPr.Elements(W.szCs).Attributes(W.val).FirstOrDefault();
            else
                szn = (decimal?)rPr.Elements(W.sz).Attributes(W.val).FirstOrDefault();
            if (szn == null)
                szn = 22m;

            decimal sz = szn.GetValueOrDefault();

            // unknown font families will throw ArgumentException, in which case just return 0
            if (!KnownFamilies.Contains(fontName))
                return 0;

            // in theory, all unknown fonts are found by the above test, but if not...
            FontFamily ff;
            try
            {
                ff = new FontFamily(fontName);
            }
            catch (ArgumentException)
            {
                UnknownFonts.Add(fontName);

                return 0;
            }

            var fs = FontStyle.Regular;
            bool bold = GetBoolProp(rPr, W.b) || GetBoolProp(rPr, W.bCs);
            bool italic = GetBoolProp(rPr, W.i) || GetBoolProp(rPr, W.iCs);
            if (bold && !italic)
                fs = FontStyle.Bold;
            if (italic && !bold)
                fs = FontStyle.Italic;
            if (bold && italic)
                fs = FontStyle.Bold | FontStyle.Italic;

            string runText = r.DescendantsTrimmed(W.txbxContent)
                .Where(e => e.Name == W.t)
                .Select(t => (string)t)
                .StringConcatenate();

            decimal tabLength = r.DescendantsTrimmed(W.txbxContent)
                .Where(e => e.Name == W.tab)
                .Select(t => (decimal)t.Attribute(PtOpenXml.TabWidth))
                .Sum();

            if (runText.Length == 0 && tabLength == 0)
                return 0;

            var multiplier = 1;
            if (runText.Length <= 2)
                multiplier = 100;
            else if (runText.Length <= 4)
                multiplier = 50;
            else if (runText.Length <= 8)
                multiplier = 25;
            else if (runText.Length <= 16)
                multiplier = 12;
            else if (runText.Length <= 32)
                multiplier = 6;
            if (multiplier != 1)
            {
                var sb = new StringBuilder();
                for (var i = 0; i < multiplier; i++)
                    sb.Append(runText);
                runText = sb.ToString();
            }

            int w = MetricsGetter.GetTextWidth(ff, fs, sz, runText);

            return (int)(w / 96m * 1440m / multiplier + tabLength * 1440m);
        }

        public static bool GetBoolProp(XElement runProps, XName xName)
        {
            XElement p = runProps.Element(xName);
            if (p == null)
                return false;

            XAttribute v = p.Attribute(W.val);
            if (v == null)
                return true;

            string s = v.Value.ToLower();
            if (s == "0" || s == "false")
                return false;
            if (s == "1" || s == "true")
                return true;

            return false;
        }

        public static int StringToTwips(string twipsOrPoints)
        {
            // if the pos value is in points, not twips
            if (twipsOrPoints.EndsWith("pt"))
            {
                decimal decimalValue = decimal.Parse(twipsOrPoints.Substring(0, twipsOrPoints.Length - 2));
                return (int)(decimalValue * 20);
            }

            return int.Parse(twipsOrPoints);
        }

        public static int? AttributeToTwips(XAttribute attribute)
        {
            if (attribute == null)
            {
                return null;
            }

            var twipsOrPoints = (string)attribute;

            // if the pos value is in points, not twips
            if (twipsOrPoints.EndsWith("pt"))
            {
                decimal decimalValue = decimal.Parse(twipsOrPoints.Substring(0, twipsOrPoints.Length - 2));
                return (int)(decimalValue * 20);
            }

            return int.Parse(twipsOrPoints);
        }

        private static readonly List<XName> AdditionalRunContainerNames = new List<XName>
        {
            W.w + "bdo",
            W.customXml,
            W.dir,
            W.fldSimple,
            W.hyperlink,
            W.moveFrom,
            W.moveTo,
            W.sdtContent
        };

        public static XElement CoalesceAdjacentRunsWithIdenticalFormatting(XElement runContainer)
        {
            const string dontConsolidate = "DontConsolidate";

            IEnumerable<IGrouping<string, XElement>> groupedAdjacentRunsWithIdenticalFormatting =
                runContainer
                    .Elements()
                    .GroupAdjacent(ce =>
                    {
                        if (ce.Name == W.r)
                        {
                            if (ce.Elements().Count(e => e.Name != W.rPr) != 1)
                                return dontConsolidate;

                            if (ce.Attribute(PtOpenXml.AbstractNumId) != null)
                                return dontConsolidate;

                            XElement rPr = ce.Element(W.rPr);
                            string rPrString = rPr != null ? rPr.ToString(SaveOptions.None) : string.Empty;

                            if (ce.Element(W.t) != null)
                                return "Wt" + rPrString;

                            if (ce.Element(W.instrText) != null)
                                return "WinstrText" + rPrString;

                            return dontConsolidate;
                        }

                        if (ce.Name == W.ins)
                        {
                            if (ce.Elements(W.del).Any())
                            {
                                return dontConsolidate;
#if false
                                // for w:ins/w:del/w:r/w:delText
                                if ((ce.Elements(W.del).Elements(W.r).Elements().Count(e => e.Name != W.rPr) != 1) ||
                                    !ce.Elements().Elements().Elements(W.delText).Any())
                                    return dontConsolidate;

                                XAttribute dateIns = ce.Attribute(W.date);
                                XElement del = ce.Element(W.del);
                                XAttribute dateDel = del.Attribute(W.date);

                                string authorIns = (string) ce.Attribute(W.author) ?? string.Empty;
                                string dateInsString = dateIns != null
                                    ? ((DateTime) dateIns).ToString("s")
                                    : string.Empty;
                                string authorDel = (string) del.Attribute(W.author) ?? string.Empty;
                                string dateDelString = dateDel != null
                                    ? ((DateTime) dateDel).ToString("s")
                                    : string.Empty;

                                return "Wins" +
                                       authorIns +
                                       dateInsString +
                                       authorDel +
                                       dateDelString +
                                       ce.Elements(W.del)
                                           .Elements(W.r)
                                           .Elements(W.rPr)
                                           .Select(rPr => rPr.ToString(SaveOptions.None))
                                           .StringConcatenate();
#endif
                            }

                            // w:ins/w:r/w:t
                            if (ce.Elements().Elements().Count(e => e.Name != W.rPr) != 1 ||
                                !ce.Elements().Elements(W.t).Any())
                                return dontConsolidate;

                            XAttribute dateIns2 = ce.Attribute(W.date);

                            string authorIns2 = (string)ce.Attribute(W.author) ?? string.Empty;
                            string dateInsString2 = dateIns2 != null
                                ? ((DateTime)dateIns2).ToString("s")
                                : string.Empty;

                            var idIns2 = (string)ce.Attribute(W.id);

                            return "Wins2" +
                                   authorIns2 +
                                   dateInsString2 +
                                   idIns2 +
                                   ce.Elements()
                                       .Elements(W.rPr)
                                       .Select(rPr => rPr.ToString(SaveOptions.None))
                                       .StringConcatenate();
                        }

                        if (ce.Name == W.del)
                        {
                            if (ce.Elements(W.r).Elements().Count(e => e.Name != W.rPr) != 1 ||
                                !ce.Elements().Elements(W.delText).Any())
                                return dontConsolidate;

                            XAttribute dateDel2 = ce.Attribute(W.date);

                            string authorDel2 = (string)ce.Attribute(W.author) ?? string.Empty;
                            string dateDelString2 = dateDel2 != null ? ((DateTime)dateDel2).ToString("s") : string.Empty;

                            return "Wdel" +
                                   authorDel2 +
                                   dateDelString2 +
                                   ce.Elements(W.r)
                                       .Elements(W.rPr)
                                       .Select(rPr => rPr.ToString(SaveOptions.None))
                                       .StringConcatenate();
                        }

                        return dontConsolidate;
                    });

            var runContainerWithConsolidatedRuns = new XElement(runContainer.Name,
                runContainer.Attributes(),
                groupedAdjacentRunsWithIdenticalFormatting.Select(g =>
                {
                    if (g.Key == dontConsolidate)
                        return (object)g;

                    string textValue = g
                        .Select(r =>
                            r.Descendants()
                                .Where(d => d.Name == W.t || d.Name == W.delText || d.Name == W.instrText)
                                .Select(d => d.Value)
                                .StringConcatenate())
                        .StringConcatenate();
                    XAttribute xs = XmlUtil.GetXmlSpaceAttribute(textValue);

                    if (g.First().Name == W.r)
                    {
                        if (g.First().Element(W.t) != null)
                        {
                            IEnumerable<IEnumerable<XAttribute>> statusAtt =
                                g.Select(r => r.Descendants(W.t).Take(1).Attributes(PtOpenXml.Status));
                            return new XElement(W.r,
                                g.First().Attributes(),
                                g.First().Elements(W.rPr),
                                new XElement(W.t, statusAtt, xs, textValue));
                        }

                        if (g.First().Element(W.instrText) != null)
                            return new XElement(W.r,
                                g.First().Attributes(),
                                g.First().Elements(W.rPr),
                                new XElement(W.instrText, xs, textValue));
                    }

                    if (g.First().Name == W.ins)
                    {
                        XElement firstR = g.First().Element(W.r);
                        return new XElement(W.ins,
                            g.First().Attributes(),
                            new XElement(W.r,
                                firstR?.Attributes(),
                                g.First().Elements(W.r).Elements(W.rPr),
                                new XElement(W.t, xs, textValue)));
                    }

                    if (g.First().Name == W.del)
                    {
                        XElement firstR = g.First().Element(W.r);
                        return new XElement(W.del,
                            g.First().Attributes(),
                            new XElement(W.r,
                                firstR?.Attributes(),
                                g.First().Elements(W.r).Elements(W.rPr),
                                new XElement(W.delText, xs, textValue)));
                    }

                    return g;
                }));

            // Process w:txbxContent//w:p
            foreach (XElement txbx in runContainerWithConsolidatedRuns.Descendants(W.txbxContent))
            foreach (XElement txbxPara in txbx.DescendantsTrimmed(W.txbxContent).Where(d => d.Name == W.p))
            {
                XElement newPara = CoalesceAdjacentRunsWithIdenticalFormatting(txbxPara);
                txbxPara.ReplaceWith(newPara);
            }

            // Process additional run containers.
            List<XElement> runContainers = runContainerWithConsolidatedRuns
                .Descendants()
                .Where(d => AdditionalRunContainerNames.Contains(d.Name))
                .ToList();
            foreach (XElement container in runContainers)
            {
                XElement newContainer = CoalesceAdjacentRunsWithIdenticalFormatting(container);
                container.ReplaceWith(newContainer);
            }

            return runContainerWithConsolidatedRuns;
        }

        private static readonly Dictionary<XName, int> Order_settings = new Dictionary<XName, int>
        {
            { W.writeProtection, 10 },
            { W.view, 20 },
            { W.zoom, 30 },
            { W.removePersonalInformation, 40 },
            { W.removeDateAndTime, 50 },
            { W.doNotDisplayPageBoundaries, 60 },
            { W.displayBackgroundShape, 70 },
            { W.printPostScriptOverText, 80 },
            { W.printFractionalCharacterWidth, 90 },
            { W.printFormsData, 100 },
            { W.embedTrueTypeFonts, 110 },
            { W.embedSystemFonts, 120 },
            { W.saveSubsetFonts, 130 },
            { W.saveFormsData, 140 },
            { W.mirrorMargins, 150 },
            { W.alignBordersAndEdges, 160 },
            { W.bordersDoNotSurroundHeader, 170 },
            { W.bordersDoNotSurroundFooter, 180 },
            { W.gutterAtTop, 190 },
            { W.hideSpellingErrors, 200 },
            { W.hideGrammaticalErrors, 210 },
            { W.activeWritingStyle, 220 },
            { W.proofState, 230 },
            { W.formsDesign, 240 },
            { W.attachedTemplate, 250 },
            { W.linkStyles, 260 },
            { W.stylePaneFormatFilter, 270 },
            { W.stylePaneSortMethod, 280 },
            { W.documentType, 290 },
            { W.mailMerge, 300 },
            { W.revisionView, 310 },
            { W.trackRevisions, 320 },
            { W.doNotTrackMoves, 330 },
            { W.doNotTrackFormatting, 340 },
            { W.documentProtection, 350 },
            { W.autoFormatOverride, 360 },
            { W.styleLockTheme, 370 },
            { W.styleLockQFSet, 380 },
            { W.defaultTabStop, 390 },
            { W.autoHyphenation, 400 },
            { W.consecutiveHyphenLimit, 410 },
            { W.hyphenationZone, 420 },
            { W.doNotHyphenateCaps, 430 },
            { W.showEnvelope, 440 },
            { W.summaryLength, 450 },
            { W.clickAndTypeStyle, 460 },
            { W.defaultTableStyle, 470 },
            { W.evenAndOddHeaders, 480 },
            { W.bookFoldRevPrinting, 490 },
            { W.bookFoldPrinting, 500 },
            { W.bookFoldPrintingSheets, 510 },
            { W.drawingGridHorizontalSpacing, 520 },
            { W.drawingGridVerticalSpacing, 530 },
            { W.displayHorizontalDrawingGridEvery, 540 },
            { W.displayVerticalDrawingGridEvery, 550 },
            { W.doNotUseMarginsForDrawingGridOrigin, 560 },
            { W.drawingGridHorizontalOrigin, 570 },
            { W.drawingGridVerticalOrigin, 580 },
            { W.doNotShadeFormData, 590 },
            { W.noPunctuationKerning, 600 },
            { W.characterSpacingControl, 610 },
            { W.printTwoOnOne, 620 },
            { W.strictFirstAndLastChars, 630 },
            { W.noLineBreaksAfter, 640 },
            { W.noLineBreaksBefore, 650 },
            { W.savePreviewPicture, 660 },
            { W.doNotValidateAgainstSchema, 670 },
            { W.saveInvalidXml, 680 },
            { W.ignoreMixedContent, 690 },
            { W.alwaysShowPlaceholderText, 700 },
            { W.doNotDemarcateInvalidXml, 710 },
            { W.saveXmlDataOnly, 720 },
            { W.useXSLTWhenSaving, 730 },
            { W.saveThroughXslt, 740 },
            { W.showXMLTags, 750 },
            { W.alwaysMergeEmptyNamespace, 760 },
            { W.updateFields, 770 },
            { W.footnotePr, 780 },
            { W.endnotePr, 790 },
            { W.compat, 800 },
            { W.docVars, 810 },
            { W.rsids, 820 },
            { M.mathPr, 830 },
            { W.attachedSchema, 840 },
            { W.themeFontLang, 850 },
            { W.clrSchemeMapping, 860 },
            { W.doNotIncludeSubdocsInStats, 870 },
            { W.doNotAutoCompressPictures, 880 },
            { W.forceUpgrade, 890 },

            //{W.captions, 900},
            { W.readModeInkLockDown, 910 },
            { W.smartTagType, 920 },

            //{W.sl:schemaLibrary, 930},
            { W.doNotEmbedSmartTags, 940 },
            { W.decimalSymbol, 950 },
            { W.listSeparator, 960 }
        };

#if false
// from the schema in the standard

writeProtection
view
zoom
removePersonalInformation
removeDateAndTime
doNotDisplayPageBoundaries
displayBackgroundShape
printPostScriptOverText
printFractionalCharacterWidth
printFormsData
embedTrueTypeFonts
embedSystemFonts
saveSubsetFonts
saveFormsData
mirrorMargins
alignBordersAndEdges
bordersDoNotSurroundHeader
bordersDoNotSurroundFooter
gutterAtTop
hideSpellingErrors
hideGrammaticalErrors
activeWritingStyle
proofState
formsDesign
attachedTemplate
linkStyles
stylePaneFormatFilter
stylePaneSortMethod
documentType
mailMerge
revisionView
trackRevisions
doNotTrackMoves
doNotTrackFormatting
documentProtection
autoFormatOverride
styleLockTheme
styleLockQFSet
defaultTabStop
autoHyphenation
consecutiveHyphenLimit
hyphenationZone
doNotHyphenateCaps
showEnvelope
summaryLength
clickAndTypeStyle
defaultTableStyle
evenAndOddHeaders
bookFoldRevPrinting
bookFoldPrinting
bookFoldPrintingSheets
drawingGridHorizontalSpacing
drawingGridVerticalSpacing
displayHorizontalDrawingGridEvery
displayVerticalDrawingGridEvery
doNotUseMarginsForDrawingGridOrigin
drawingGridHorizontalOrigin
drawingGridVerticalOrigin
doNotShadeFormData
noPunctuationKerning
characterSpacingControl
printTwoOnOne
strictFirstAndLastChars
noLineBreaksAfter
noLineBreaksBefore
savePreviewPicture
doNotValidateAgainstSchema
saveInvalidXml
ignoreMixedContent
alwaysShowPlaceholderText
doNotDemarcateInvalidXml
saveXmlDataOnly
useXSLTWhenSaving
saveThroughXslt
showXMLTags
alwaysMergeEmptyNamespace
updateFields
footnotePr
endnotePr
compat
docVars
rsids
m:mathPr
attachedSchema
themeFontLang
clrSchemeMapping
doNotIncludeSubdocsInStats
doNotAutoCompressPictures
forceUpgrade
captions
readModeInkLockDown
smartTagType
sl:schemaLibrary
doNotEmbedSmartTags
decimalSymbol
listSeparator
#endif

        private static readonly Dictionary<XName, int> Order_pPr = new Dictionary<XName, int>
        {
            { W.pStyle, 10 },
            { W.keepNext, 20 },
            { W.keepLines, 30 },
            { W.pageBreakBefore, 40 },
            { W.framePr, 50 },
            { W.widowControl, 60 },
            { W.numPr, 70 },
            { W.suppressLineNumbers, 80 },
            { W.pBdr, 90 },
            { W.shd, 100 },
            { W.tabs, 120 },
            { W.suppressAutoHyphens, 130 },
            { W.kinsoku, 140 },
            { W.wordWrap, 150 },
            { W.overflowPunct, 160 },
            { W.topLinePunct, 170 },
            { W.autoSpaceDE, 180 },
            { W.autoSpaceDN, 190 },
            { W.bidi, 200 },
            { W.adjustRightInd, 210 },
            { W.snapToGrid, 220 },
            { W.spacing, 230 },
            { W.ind, 240 },
            { W.contextualSpacing, 250 },
            { W.mirrorIndents, 260 },
            { W.suppressOverlap, 270 },
            { W.jc, 280 },
            { W.textDirection, 290 },
            { W.textAlignment, 300 },
            { W.textboxTightWrap, 310 },
            { W.outlineLvl, 320 },
            { W.divId, 330 },
            { W.cnfStyle, 340 },
            { W.rPr, 350 },
            { W.sectPr, 360 },
            { W.pPrChange, 370 }
        };

        private static readonly Dictionary<XName, int> Order_rPr = new Dictionary<XName, int>
        {
            { W.moveFrom, 5 },
            { W.moveTo, 7 },
            { W.ins, 10 },
            { W.del, 20 },
            { W.rStyle, 30 },
            { W.rFonts, 40 },
            { W.b, 50 },
            { W.bCs, 60 },
            { W.i, 70 },
            { W.iCs, 80 },
            { W.caps, 90 },
            { W.smallCaps, 100 },
            { W.strike, 110 },
            { W.dstrike, 120 },
            { W.outline, 130 },
            { W.shadow, 140 },
            { W.emboss, 150 },
            { W.imprint, 160 },
            { W.noProof, 170 },
            { W.snapToGrid, 180 },
            { W.vanish, 190 },
            { W.webHidden, 200 },
            { W.color, 210 },
            { W.spacing, 220 },
            { W._w, 230 },
            { W.kern, 240 },
            { W.position, 250 },
            { W.sz, 260 },
            { W14.wShadow, 270 },
            { W14.wTextOutline, 280 },
            { W14.wTextFill, 290 },
            { W14.wScene3d, 300 },
            { W14.wProps3d, 310 },
            { W.szCs, 320 },
            { W.highlight, 330 },
            { W.u, 340 },
            { W.effect, 350 },
            { W.bdr, 360 },
            { W.shd, 370 },
            { W.fitText, 380 },
            { W.vertAlign, 390 },
            { W.rtl, 400 },
            { W.cs, 410 },
            { W.em, 420 },
            { W.lang, 430 },
            { W.eastAsianLayout, 440 },
            { W.specVanish, 450 },
            { W.oMath, 460 }
        };

        private static readonly Dictionary<XName, int> Order_tblPr = new Dictionary<XName, int>
        {
            { W.tblStyle, 10 },
            { W.tblpPr, 20 },
            { W.tblOverlap, 30 },
            { W.bidiVisual, 40 },
            { W.tblStyleRowBandSize, 50 },
            { W.tblStyleColBandSize, 60 },
            { W.tblW, 70 },
            { W.jc, 80 },
            { W.tblCellSpacing, 90 },
            { W.tblInd, 100 },
            { W.tblBorders, 110 },
            { W.shd, 120 },
            { W.tblLayout, 130 },
            { W.tblCellMar, 140 },
            { W.tblLook, 150 },
            { W.tblCaption, 160 },
            { W.tblDescription, 170 }
        };

        private static readonly Dictionary<XName, int> Order_tblBorders = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.left, 20 },
            { W.start, 30 },
            { W.bottom, 40 },
            { W.right, 50 },
            { W.end, 60 },
            { W.insideH, 70 },
            { W.insideV, 80 }
        };

        private static readonly Dictionary<XName, int> Order_tcPr = new Dictionary<XName, int>
        {
            { W.cnfStyle, 10 },
            { W.tcW, 20 },
            { W.gridSpan, 30 },
            { W.hMerge, 40 },
            { W.vMerge, 50 },
            { W.tcBorders, 60 },
            { W.shd, 70 },
            { W.noWrap, 80 },
            { W.tcMar, 90 },
            { W.textDirection, 100 },
            { W.tcFitText, 110 },
            { W.vAlign, 120 },
            { W.hideMark, 130 },
            { W.headers, 140 }
        };

        private static readonly Dictionary<XName, int> Order_tcBorders = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.start, 20 },
            { W.left, 30 },
            { W.bottom, 40 },
            { W.right, 50 },
            { W.end, 60 },
            { W.insideH, 70 },
            { W.insideV, 80 },
            { W.tl2br, 90 },
            { W.tr2bl, 100 }
        };

        private static readonly Dictionary<XName, int> Order_pBdr = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.left, 20 },
            { W.bottom, 30 },
            { W.right, 40 },
            { W.between, 50 },
            { W.bar, 60 }
        };

        public static object WmlOrderElementsPerStandard(XNode node)
        {
            var element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.pPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Order_pPr.ContainsKey(e.Name))
                                return Order_pPr[e.Name];

                            return 999;
                        }));

                if (element.Name == W.rPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Order_rPr.ContainsKey(e.Name))
                                return Order_rPr[e.Name];

                            return 999;
                        }));

                if (element.Name == W.tblPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Order_tblPr.ContainsKey(e.Name))
                                return Order_tblPr[e.Name];

                            return 999;
                        }));

                if (element.Name == W.tcPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Order_tcPr.ContainsKey(e.Name))
                                return Order_tcPr[e.Name];

                            return 999;
                        }));

                if (element.Name == W.tcBorders)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Order_tcBorders.ContainsKey(e.Name))
                                return Order_tcBorders[e.Name];

                            return 999;
                        }));

                if (element.Name == W.tblBorders)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Order_tblBorders.ContainsKey(e.Name))
                                return Order_tblBorders[e.Name];

                            return 999;
                        }));

                if (element.Name == W.pBdr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Order_pBdr.ContainsKey(e.Name))
                                return Order_pBdr[e.Name];

                            return 999;
                        }));

                if (element.Name == W.p)
                {
                    var newP = new XElement(element.Name,
                        element.Attributes(),
                        element.Elements(W.pPr).Select(e => (XElement)WmlOrderElementsPerStandard(e)),
                        element.Elements().Where(e => e.Name != W.pPr).Select(e => (XElement)WmlOrderElementsPerStandard(e)));
                    return newP;
                }

                if (element.Name == W.r)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements(W.rPr).Select(e => (XElement)WmlOrderElementsPerStandard(e)),
                        element.Elements().Where(e => e.Name != W.rPr).Select(e => (XElement)WmlOrderElementsPerStandard(e)));

                if (element.Name == W.settings)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Order_settings.ContainsKey(e.Name))
                                return Order_settings[e.Name];

                            return 999;
                        }));

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => WmlOrderElementsPerStandard(n)));
            }

            return node;
        }

        public static WmlDocument BreakLinkToTemplate(WmlDocument source)
        {
            using (var ms = new MemoryStream())
            {
                ms.Write(source.DocumentByteArray, 0, source.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    ExtendedFilePropertiesPart efpp = wDoc.ExtendedFilePropertiesPart;
                    if (efpp != null)
                    {
                        XDocument xd = efpp.GetXDocument();
                        XElement template = xd.Descendants(EP.Template).FirstOrDefault();
                        if (template != null)
                            template.Value = "";
                        efpp.PutXDocument();
                    }
                }

                var result = new WmlDocument(source.FileName, ms.ToArray());
                return result;
            }
        }
    }
}
