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

/***************************************************************************
 * HTML elements handled in this module:
 * 
 * a
 * b
 * body
 * caption
 * div
 * em
 * h1, h2, h3, h4, h5, h6, h7, h8
 * hr
 * html
 * i
 * blockquote
 * img
 * li
 * ol
 * p
 * s
 * span
 * strong
 * style
 * sub
 * sup
 * table
 * tbody
 * td
 * th
 * tr
 * u
 * ul
 * br
 * tt
 * code
 * kbd
 * samp
 * pre
 * 
 * HTML elements that are handled by recursively processing descedants
 * 
 * article
 * hgroup
 * nav
 * section
 * dd
 * dl
 * dt
 * figure
 * main
 * abbr
 * bdi
 * bdo
 * cite
 * data
 * dfn
 * mark
 * q
 * rp
 * rt
 * ruby
 * small
 * time
 * var
 * wbr
 * 
 * HTML elements ignored in this module
 * 
 * head
 * 
***************************************************************************/

// need to research all of the html attributes that take effect, such as border="1" and somehow work into the rendering system.
// note that some of these 'inherit' so need to implement correct logic.

// this module has not been fully engineered to work with RTL languages.  This is a pending work item.  There are issues involved,
// including that there is RTL content in HTML that can't be represented in WordprocessingML, although this probably is rare.
// The reverse is not true - all RTL WordprocessingML can be represented in HTML, but there is some HTML RTL content that can only
// be approximated in WordprocessingML.

// have I handled all forms of colors? see GetWmlColorFromExpression in HtmlToWmlCssApplier

// min-height and max-height not implemented yet.

// internal hyperlinks are not supported.  I believe it possible - bookmarks can be created, hyperlinks to the bookmark can be created.

// To be supported in future
// page-break-before:always

// ************************************************************************
// ToDo at some point in the future
// I'm not implementing caption exactly correctly.  If caption does not have borders, then there needs to not be a border around the table,
// otherwise it looks as if caption has a border.  See T1200.  If there is, per the markup, a table border, but caption does not have a border,
// then need to make sure that all of the cells below the caption have the border on the appropriate sides so that it looks as if the table
// has a border.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using OpenXmlPowerTools.HtmlToWml;
using OpenXmlPowerTools.HtmlToWml.CSS;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace OpenXmlPowerTools.HtmlToWml
{
    public class ElementToStyleMap
    {
        public string ElementName;
        public string StyleName;
    }

    public static class LocalExtensions
    {
        public static CssExpression GetProp(this XElement element, string propertyName)
        {
            Dictionary<string, CssExpression> d = element.Annotation<Dictionary<string, CssExpression>>();
            if (d != null)
            {
                if (d.ContainsKey(propertyName))
                    return d[propertyName];
            }
            return null;
        }
    }

    public class HtmlToWmlConverterCore
    {
        public static WmlDocument ConvertHtmlToWml(
            string defaultCss,
            string authorCss,
            string userCss,
            XElement xhtml,
            HtmlToWmlConverterSettings settings)
        {
            return ConvertHtmlToWml(defaultCss, authorCss, userCss, xhtml, settings, null, null);
        }

        public static WmlDocument ConvertHtmlToWml(
            string defaultCss,
            string authorCss,
            string userCss,
            XElement xhtml,
            HtmlToWmlConverterSettings settings,
            WmlDocument emptyDocument,
            string annotatedHtmlDumpFileName)
        {
            if (emptyDocument == null)
                emptyDocument = HtmlToWmlConverter.EmptyDocument;

            NextRectId = 1025;

            // clone and transform all element names to lower case
            XElement html = (XElement)TransformToLower(xhtml);

            // add pseudo cells for rowspan
            html = (XElement)AddPseudoCells(html);

            html = (XElement)TransformWhiteSpaceInPreCodeTtKbdSamp(html, false, false);

            CssDocument defaultCssDoc, userCssDoc, authorCssDoc;
            CssApplier.ApplyAllCss(
                defaultCss,
                authorCss,
                userCss,
                html,
                settings,
                out defaultCssDoc,
                out authorCssDoc,
                out userCssDoc,
                annotatedHtmlDumpFileName);

            WmlDocument newWmlDocument;

            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(emptyDocument))
            {
                using (WordprocessingDocument wDoc = streamDoc.GetWordprocessingDocument())
                {
                    AnnotateOlUl(wDoc, html);
                    UpdateMainDocumentPart(wDoc, html, settings);
                    NormalizeMainDocumentPart(wDoc);
                    StylesUpdater.UpdateStylesPart(wDoc, html, settings, defaultCssDoc, authorCssDoc, userCssDoc);
                    HtmlToWmlFontUpdater.UpdateFontsPart(wDoc, html, settings);
                    ThemeUpdater.UpdateThemePart(wDoc, html, settings);
                    NumberingUpdater.UpdateNumberingPart(wDoc, html, settings);
                }
                newWmlDocument = streamDoc.GetModifiedWmlDocument();
            }

            return newWmlDocument;
        }


        private static object TransformToLower(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                XElement e = new XElement(element.Name.LocalName.ToLower(),
                    element.Attributes().Select(a => new XAttribute(a.Name.LocalName.ToLower(), a.Value)),
                    element.Nodes().Select(n => TransformToLower(n)));
                return e;
            }
            return node;
        }

        private static XElement AddPseudoCells(XElement html)
        {
            while (true)
            {
                var rowSpanCell = html
                    .Descendants(XhtmlNoNamespace.td)
                    .FirstOrDefault(td => td.Attribute(XhtmlNoNamespace.rowspan) != null && td.Attribute("HtmlToWmlVMergeRestart") == null);
                if (rowSpanCell == null)
                    break;
                rowSpanCell.Add(
                    new XAttribute("HtmlToWmlVMergeRestart", "true"));
                int colNumber = rowSpanCell.ElementsBeforeSelf(XhtmlNoNamespace.td).Count();
                int numberPseudoToAdd = (int)rowSpanCell.Attribute(XhtmlNoNamespace.rowspan) - 1;
                var tr = rowSpanCell.Ancestors(XhtmlNoNamespace.tr).FirstOrDefault();
                if (tr == null)
                    throw new OpenXmlPowerToolsException("Invalid HTML - td does not have parent tr");
                var rowsToAddTo = tr
                    .ElementsAfterSelf(XhtmlNoNamespace.tr)
                    .Take(numberPseudoToAdd)
                    .ToList();
                foreach (var rowToAddTo in rowsToAddTo)
                {
                    if (colNumber > 0)
                    {
                        var tdToAddAfter = rowToAddTo
                            .Elements(XhtmlNoNamespace.td)
                            .Skip(colNumber - 1)
                            .FirstOrDefault();
                        var td = new XElement(XhtmlNoNamespace.td,
                            rowSpanCell.Attributes(),
                            new XAttribute("HtmlToWmlVMergeNoRestart", "true"));
                        tdToAddAfter.AddAfterSelf(td);
                    }
                    else
                    {
                        var tdToAddBefore = rowToAddTo
                            .Elements(XhtmlNoNamespace.td)
                            .Skip(colNumber)
                            .FirstOrDefault();
                        var td = new XElement(XhtmlNoNamespace.td,
                            rowSpanCell.Attributes(),
                            new XAttribute("HtmlToWmlVMergeNoRestart", "true"));
                        tdToAddBefore.AddBeforeSelf(td);
                    }
                }
            }
            return html;
        }

        public class NumberedItemAnnotation
        {
            public int numId;
            public int ilvl;
            public string listStyleType;
        }

        private static void AnnotateOlUl(WordprocessingDocument wDoc, XElement html)
        {
            int numId;
            NumberingUpdater.GetNextNumId(wDoc, out numId);
            foreach (var item in html.DescendantsAndSelf().Where(d => d.Name == XhtmlNoNamespace.ol || d.Name == XhtmlNoNamespace.ul))
            {
                XElement parentOlUl = item.Ancestors().Where(a => a.Name == XhtmlNoNamespace.ol || a.Name == XhtmlNoNamespace.ul).LastOrDefault();
                int numIdToUse;
                if (parentOlUl != null)
                    numIdToUse = parentOlUl.Annotation<NumberedItemAnnotation>().numId;
                else
                    numIdToUse = numId++;
                string lst = CssApplier.GetComputedPropertyValue(null, item, "list-style-type", null).ToString();
                item.AddAnnotation(new NumberedItemAnnotation
                {
                    numId = numIdToUse,
                    ilvl = item.Ancestors().Where(a => a.Name == XhtmlNoNamespace.ol || a.Name == XhtmlNoNamespace.ul).Count(),
                    listStyleType = lst,
                });
            }
        }

        private static void UpdateMainDocumentPart(WordprocessingDocument wDoc, XElement html, HtmlToWmlConverterSettings settings)
        {
            XDocument xDoc = XDocument.Parse(
@"<w:document xmlns:wpc='http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'
            xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006'
            xmlns:o='urn:schemas-microsoft-com:office:office'
            xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'
            xmlns:v='urn:schemas-microsoft-com:vml'
            xmlns:wp14='http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'
            xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
            xmlns:w10='urn:schemas-microsoft-com:office:word'
            xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            xmlns:w14='http://schemas.microsoft.com/office/word/2010/wordml'
            xmlns:wpg='http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
            xmlns:wpi='http://schemas.microsoft.com/office/word/2010/wordprocessingInk'
            xmlns:wne='http://schemas.microsoft.com/office/word/2006/wordml'
            xmlns:wps='http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
            mc:Ignorable='w14 wp14'/>");

            XElement body = new XElement(W.body,
                        Transform(html, settings, wDoc, NextExpected.Paragraph, false),
                        settings.SectPr);

            AddNonBreakingSpacesForSpansWithWidth(wDoc, body);
            body = (XElement)TransformAndOrderElements(body);

            foreach (var d in body.Descendants())
                d.Attributes().Where(a => a.Name.Namespace == PtOpenXml.pt).Remove();
            xDoc.Root.Add(body);
            wDoc.MainDocumentPart.PutXDocument(xDoc);
        }

        private static object TransformWhiteSpaceInPreCodeTtKbdSamp(XNode node, bool inPre, bool inOther)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == XhtmlNoNamespace.pre)
                {
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Nodes().Select(n => TransformWhiteSpaceInPreCodeTtKbdSamp(n, true, false)));
                }
                if (element.Name == XhtmlNoNamespace.code ||
                    element.Name == XhtmlNoNamespace.tt ||
                    element.Name == XhtmlNoNamespace.kbd ||
                    element.Name == XhtmlNoNamespace.samp)
                {
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Nodes().Select(n => TransformWhiteSpaceInPreCodeTtKbdSamp(n, false, true)));
                }
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => TransformWhiteSpaceInPreCodeTtKbdSamp(n, false, false)));
            }
            XText xt = node as XText;
            if (xt != null && inPre)
            {
                var val = xt.Value.TrimStart('\r', '\n').TrimEnd('\r', '\n');
                var groupedCharacters = val.GroupAdjacent(c => c == '\r' || c == '\n');
                var newNodes = groupedCharacters.Select(g =>
                {
                    if (g.Key == true)
                        return (object)(new XElement(XhtmlNoNamespace.br));
                    string x = g.Select(c => c.ToString()).StringConcatenate();
                    return new XText(x);
                });
                return newNodes;
            }
            if (xt != null && inOther)
            {
                var val = xt.Value.TrimStart('\r', '\n', '\t', ' ').TrimEnd('\r', '\n', '\t', ' ');
                return new XText(val);
            }
            return node;
        }

        private static Dictionary<XName, int> Order_pPr = new Dictionary<XName, int>
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
            { W.pPrChange, 370 },
        };

        private static Dictionary<XName, int> Order_rPr = new Dictionary<XName, int>
        {
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
            { W.oMath, 460 },
        };

        private static Dictionary<XName, int> Order_tblPr = new Dictionary<XName, int>
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
            { W.tblDescription, 170 },
        };

        private static Dictionary<XName, int> Order_tblBorders = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.left, 20 },
            { W.start, 30 },
            { W.bottom, 40 },
            { W.right, 50 },
            { W.end, 60 },
            { W.insideH, 70 },
            { W.insideV, 80 },
        };

        private static Dictionary<XName, int> Order_tcPr = new Dictionary<XName, int>
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
            { W.headers, 140 },
        };

        private static Dictionary<XName, int> Order_tcBorders = new Dictionary<XName, int>
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
            { W.tr2bl, 100 },
        };

        private static Dictionary<XName, int> Order_pBdr = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.left, 20 },
            { W.bottom, 30 },
            { W.right, 40 },
            { W.between, 50 },
            { W.bar, 60 },
        };

        private static object TransformAndOrderElements(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.pPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)TransformAndOrderElements(e)).OrderBy(e =>
                        {
                            if (Order_pPr.ContainsKey(e.Name))
                                return Order_pPr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.rPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)TransformAndOrderElements(e)).OrderBy(e =>
                        {
                            if (Order_rPr.ContainsKey(e.Name))
                                return Order_rPr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.tblPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)TransformAndOrderElements(e)).OrderBy(e =>
                        {
                            if (Order_tblPr.ContainsKey(e.Name))
                                return Order_tblPr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.tcPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)TransformAndOrderElements(e)).OrderBy(e =>
                        {
                            if (Order_tcPr.ContainsKey(e.Name))
                                return Order_tcPr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.tcBorders)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)TransformAndOrderElements(e)).OrderBy(e =>
                        {
                            if (Order_tcBorders.ContainsKey(e.Name))
                                return Order_tcBorders[e.Name];
                            return 999;
                        }));

                if (element.Name == W.tblBorders)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)TransformAndOrderElements(e)).OrderBy(e =>
                        {
                            if (Order_tblBorders.ContainsKey(e.Name))
                                return Order_tblBorders[e.Name];
                            return 999;
                        }));

                if (element.Name == W.pBdr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)TransformAndOrderElements(e)).OrderBy(e =>
                        {
                            if (Order_pBdr.ContainsKey(e.Name))
                                return Order_pBdr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.p)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements(W.pPr).Select(e => (XElement)TransformAndOrderElements(e)),
                        element.Elements().Where(e => e.Name != W.pPr).Select(e => (XElement)TransformAndOrderElements(e)));

                if (element.Name == W.r)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements(W.rPr).Select(e => (XElement)TransformAndOrderElements(e)),
                        element.Elements().Where(e => e.Name != W.rPr).Select(e => (XElement)TransformAndOrderElements(e)));

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => TransformAndOrderElements(n)));
            }
            return node;
        }

        private static void AddNonBreakingSpacesForSpansWithWidth(WordprocessingDocument wDoc, XElement body)
        {
            var runsWithWidth = body.Descendants(W.r).Where(r => r.Attribute(PtOpenXml.HtmlToWmlCssWidth) != null).ToList();
            foreach (var run in runsWithWidth)
            {
                var rPr = run.Element(W.rPr);
                XElement pPr = null;
                var p = run.Ancestors(W.p).FirstOrDefault();
                if (p != null)
                    pPr = p.Element(W.pPr);
                XElement rFonts = rPr.Element(W.rFonts);
                var str = run.Descendants(W.t).Select(t => (string)t).StringConcatenate();
                if (pPr != null && rPr != null && rFonts != null && str != "")
                {
                    AdjustFontAttributes(wDoc, run, pPr, rPr);
                    var csa = new CharStyleAttributes(pPr, rPr);
                    var charToExamine = str.FirstOrDefault(c => !WeakAndNeutralDirectionalCharacters.Contains(c));
                    if (charToExamine == '\0')
                        charToExamine = str[0];

                    var ft = DetermineFontTypeFromCharacter(charToExamine, csa);
                    string fontType = null;
                    string languageType = null;
                    switch (ft)
                    {
                        case FontType.Ascii:
                            fontType = (string)rFonts.Attribute(W.ascii);
                            languageType = "western";
                            break;
                        case FontType.HAnsi:
                            fontType = (string)rFonts.Attribute(W.hAnsi);
                            languageType = "western";
                            break;
                        case FontType.EastAsia:
                            fontType = (string)rFonts.Attribute(W.eastAsia);
                            languageType = "eastAsia";
                            break;
                        case FontType.CS:
                            fontType = (string)rFonts.Attribute(W.cs);
                            languageType = "bidi";
                            break;
                    }

                    if (fontType != null)
                    {
                        if (run.Attribute(PtOpenXml.FontName) == null)
                        {
                            XAttribute fta = new XAttribute(PtOpenXml.FontName, fontType.ToString());
                            run.Add(fta);
                        }
                        else
                        {
                            run.Attribute(PtOpenXml.FontName).Value = fontType.ToString();
                        }
                    }
                    if (languageType != null)
                    {
                        if (run.Attribute(PtOpenXml.LanguageType) == null)
                        {
                            XAttribute lta = new XAttribute(PtOpenXml.LanguageType, languageType);
                            run.Add(lta);
                        }
                        else
                        {
                            run.Attribute(PtOpenXml.LanguageType).Value = languageType;
                        }
                    }

                    var pixWidth = CalcWidthOfRunInPixels(run);
                    // calc width of non breaking spaces
                    var npSpRun = new XElement(W.r,
                        run.Attributes(),
                        run.Elements(W.rPr),
                        new XElement(W.t, "\u00a0"));
                    var nbSpWidth = CalcWidthOfRunInPixels(npSpRun);
                    if (nbSpWidth == 0)
                        continue;
                    // get HtmlToWmlCssWidth attribute, convert to pixels
                    var cssWidth = (string)run.Attribute(PtOpenXml.HtmlToWmlCssWidth);
                    if (cssWidth.EndsWith("pt"))
                    {
                        cssWidth = cssWidth.Substring(0, cssWidth.Length - 2);
                        decimal cssWidthInDecimal;
                        if (decimal.TryParse(cssWidth, out cssWidthInDecimal))
                        {
                            decimal cssWidthInPixels = (cssWidthInDecimal / 72) * 96;
                            // calculate the number of non-breaking spaces to add
                            var numberOfNpSpToAdd = (cssWidthInPixels - pixWidth) / nbSpWidth;
                            // then add them.
                            run.Add(new XElement(W.t, "".PadRight((int)numberOfNpSpToAdd, '\u00a0')));
                        }
                    }
                }
            }
        }

        private static void NormalizeMainDocumentPart(WordprocessingDocument wDoc)
        {
            XDocument mainXDoc = wDoc.MainDocumentPart.GetXDocument();
            XElement newRoot = (XElement)NormalizeTransform(mainXDoc.Root);
            mainXDoc.Root.ReplaceWith(newRoot);
            wDoc.MainDocumentPart.PutXDocument();
        }

        private static object NormalizeTransform(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.p && element.Elements().Any(c => c.Name == W.p || c.Name == W.tbl))
                {
                    var groupedChildren = element.Elements()
                        .GroupAdjacent(e => e.Name == W.p || e.Name == W.tbl);
                    var newContent = groupedChildren
                        .Select(g =>
                        {
                            if (g.Key == false)
                            {
                                XElement paragraph = new XElement(W.p,
                                    element.Elements(W.pPr),
                                    g.Where(gc => gc.Name != W.pPr));
                                return (object)paragraph;
                            }
                            return g.Select(n => NormalizeTransform(n));
                        });
                    return newContent;
                }
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => NormalizeTransform(n)));
            }
            return node;
        }

        private enum NextExpected
        {
            Paragraph,
            Run,
            SubRun,
        }

        private static object Transform(XNode node, HtmlToWmlConverterSettings settings, WordprocessingDocument wDoc, NextExpected nextExpected, bool preserveWhiteSpace)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == XhtmlNoNamespace.a)
                {
                    string rId = "R" + Guid.NewGuid().ToString().Replace("-", "");
                    string href = (string)element.Attribute(NoNamespace.href);
                    if (href != null)
                    {
                        Uri uri = null;
                        try
                        {
                            uri = new Uri(href);
                        }
                        catch (UriFormatException)
                        {
                            return null;
                        }

                        if (uri != null)
                        {
                            wDoc.MainDocumentPart.AddHyperlinkRelationship(uri, true, rId);
                            XElement rPr = GetRunProperties(element, settings);
                            XElement hyperlink = new XElement(W.hyperlink,
                                new XAttribute(R.id, rId),
                                new XElement(W.r,
                                    rPr,
                                    new XElement(W.t, element.Value)));
                            return new[] { hyperlink };
                        }
                    }
                    return null;
                }

                if (element.Name == XhtmlNoNamespace.b)
                    return element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.body)
                    return element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.caption)
                {
                    return new XElement(W.tr,
                        GetTableRowProperties(element),
                        new XElement(W.tc,
                            GetCellPropertiesForCaption(element),
                            element.Nodes().Select(n => Transform(n, settings, wDoc, NextExpected.Paragraph, preserveWhiteSpace))));
                }

                if (element.Name == XhtmlNoNamespace.div)
                {
                    if (nextExpected == NextExpected.Paragraph)
                    {
                        if (element.Descendants().Any(d => d.Name == XhtmlNoNamespace.h1 ||
                            d.Name == XhtmlNoNamespace.li ||
                            d.Name == XhtmlNoNamespace.p ||
                            d.Name == XhtmlNoNamespace.table))
                        {
                            return element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace));
                        }
                        else
                        {
                            return GenerateNextExpected(element, settings, wDoc, null, nextExpected, false);
                        }
                    }
                    else
                    {
                        return element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace));
                    }
                }

                if (element.Name == XhtmlNoNamespace.em)
                    return element.Nodes().Select(n => Transform(n, settings, wDoc, NextExpected.Run, preserveWhiteSpace));

                HeadingInfo hi = HeadingTagMap.FirstOrDefault(htm => htm.Name == element.Name);
                if (hi != null)
                {
                    return GenerateNextExpected(element, settings, wDoc, hi.StyleName, NextExpected.Paragraph, false);
                }

                if (element.Name == XhtmlNoNamespace.hr)
                {
                    int i = GetNextRectId();
                    XElement hr = XElement.Parse(
                      @"<w:p xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                               xmlns:o='urn:schemas-microsoft-com:office:office'
                               xmlns:w14='http://schemas.microsoft.com/office/word/2010/wordml'
                               xmlns:v='urn:schemas-microsoft-com:vml'>
                          <w:r>
                            <w:pict w14:anchorId='0DBC9ADE'>
                              <v:rect id='_x0000_i" + i + @"'
                                      style='width:0;height:1.5pt'
                                      o:hralign='center'
                                      o:hrstd='t'
                                      o:hr='t'
                                      fillcolor='#a0a0a0'
                                      stroked='f'/>
                            </w:pict>
                          </w:r>
                        </w:p>");
                    hr.Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
                    return hr;
                }

                if (element.Name == XhtmlNoNamespace.html)
                    return element.Nodes().Select(n => Transform(n, settings, wDoc, NextExpected.Paragraph, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.i)
                    return element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.blockquote)
                    return element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.img)
                {
                    if (element.Parent.Name == XhtmlNoNamespace.body)
                    {
                        XElement para = new XElement(W.p,
                            GetParagraphPropertiesForImage(),
                            TransformImageToWml(element, settings, wDoc));
                        return para;
                    }
                    else
                    {
                        XElement content = TransformImageToWml(element, settings, wDoc);
                        return content;
                    }
                }

                if (element.Name == XhtmlNoNamespace.li)
                {
                    return GenerateNextExpected(element, settings, wDoc, null, NextExpected.Paragraph, false);
                }

                if (element.Name == XhtmlNoNamespace.ol)
                    return element.Nodes().Select(n => Transform(n, settings, wDoc, NextExpected.Paragraph, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.p)
                {
                    return GenerateNextExpected(element, settings, wDoc, null, NextExpected.Paragraph, false);
                }

                if (element.Name == XhtmlNoNamespace.s)
                    return element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace));

                /****************************************** SharePoint Specific ********************************************/
                // todo sharepoint specific
                //if (element.Name == Xhtml.div && (string)element.Attribute(Xhtml._class) == "ms-rteElement-Callout1")
                //{
                //    return new XElement(W.p,
                //        // todo need a style for class
                //        new XElement(W.pPr,
                //            new XElement(W.pStyle,
                //                new XAttribute(W.val, "ms-rteElement-Callout1"))),
                //        new XElement(W.r,
                //            new XElement(W.t, element.Value)));
                //}
                if (element.Name == XhtmlNoNamespace.span && (string)element.Attribute(XhtmlNoNamespace.id) == "layoutsData")
                    return null;
                /****************************************** End SharePoint Specific ********************************************/

                if (element.Name == XhtmlNoNamespace.span)
                {
                    var spanReplacement = element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace));
                    var dummyElement = new XElement("dummy", spanReplacement);
                    var firstChild = dummyElement.Elements().FirstOrDefault();
                    XElement run = null;
                    if (firstChild != null && firstChild.Name == W.r)
                        run = firstChild;
                    if (run != null)
                    {
                        Dictionary<string, CssExpression> computedProperties = element.Annotation<Dictionary<string, CssExpression>>();
                        if (computedProperties != null && computedProperties.ContainsKey("width"))
                        {
                            string width = computedProperties["width"];
                            if (width != "auto")
                                run.Add(new XAttribute(PtOpenXml.HtmlToWmlCssWidth, width));
                            var rFontsLocal = run.Element(W.rFonts);
                            XElement rFontsGlobal = null;
                            var styleDefPart = wDoc.MainDocumentPart.StyleDefinitionsPart;
                            if (styleDefPart != null)
                            {
                                rFontsGlobal = styleDefPart.GetXDocument().Root.Elements(W.docDefaults).Elements(W.rPrDefault).Elements(W.rPr).Elements(W.rFonts).FirstOrDefault();
                            }
                            var rFontsNew = FontMerge(rFontsLocal, rFontsGlobal);
                            var rPr = run.Element(W.rPr);
                            if (rPr != null)
                            {
                                var rFontsExisting = rPr.Element(W.rFonts);
                                if (rFontsExisting == null)
                                    rPr.AddFirst(rFontsGlobal);
                                else
                                    rFontsExisting.ReplaceWith(rFontsGlobal);
                            }
                        }
                        return dummyElement.Elements();
                    }

                    return spanReplacement;
                }

                if (element.Name == XhtmlNoNamespace.strong)
                    return element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.style)
                    return null;

                if (element.Name == XhtmlNoNamespace.sub)
                    return element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.sup)
                    return element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.table)
                {
                    XElement wmlTable = new XElement(W.tbl,
                        GetTableProperties(element),
                        GetTableGrid(element, settings),
                        element.Nodes().Select(n => Transform(n, settings, wDoc, NextExpected.Paragraph, preserveWhiteSpace)));
                    return wmlTable;
                }

                if (element.Name == XhtmlNoNamespace.tbody)
                    return element.Nodes().Select(n => Transform(n, settings, wDoc, NextExpected.Paragraph, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.td)
                {
                    var tdText = element.DescendantNodes().OfType<XText>().Select(t => t.Value).StringConcatenate().Trim();
                    var hasOtherThanSpansAndParas = element.Descendants().Any(d => d.Name != XhtmlNoNamespace.span && d.Name != XhtmlNoNamespace.p);
                    if (tdText != "" || hasOtherThanSpansAndParas)
                    {
                        return new XElement(W.tc,
                            GetCellProperties(element),
                            element.Nodes().Select(n => Transform(n, settings, wDoc, NextExpected.Paragraph, preserveWhiteSpace)));
                    }
                    else
                    {
                        XElement p;
                        p = new XElement(W.p,
                            GetParagraphProperties(element, null, settings),
                            new XElement(W.r,
                                GetRunProperties(element, settings),
                                new XElement(W.t, "")));
                        return new XElement(W.tc,
                            GetCellProperties(element), p);
                    }
                }

                if (element.Name == XhtmlNoNamespace.th)
                {
                    return new XElement(W.tc,
                        GetCellHeaderProperties(element),
                        element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace)));
                }

                if (element.Name == XhtmlNoNamespace.tr)
                {
                    return new XElement(W.tr,
                        GetTableRowProperties(element),
                        element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace)));
                }

                if (element.Name == XhtmlNoNamespace.u)
                    return element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.ul)
                    return element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.br)
                    if (nextExpected == NextExpected.Paragraph)
                    {
                        return new XElement(W.p,
                            new XElement(W.r,
                                new XElement(W.t)));
                    }
                    else
                    {
                        return new XElement(W.r, new XElement(W.br));
                    }

                if (element.Name == XhtmlNoNamespace.tt || element.Name == XhtmlNoNamespace.code || element.Name == XhtmlNoNamespace.kbd || element.Name == XhtmlNoNamespace.samp)
                    return element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.pre)
                    return GenerateNextExpected(element, settings, wDoc, null, NextExpected.Paragraph, true);

                // if no match up to this point, then just recursively process descendants
                return element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace));
            }

            if (node.Parent.Name != XhtmlNoNamespace.title)
                return GenerateNextExpected(node, settings, wDoc, null, nextExpected, preserveWhiteSpace);

            return null;

        }

        private static XElement FontMerge(XElement higherPriorityFont, XElement lowerPriorityFont)
        {
            XElement rFonts;

            if (higherPriorityFont == null)
                return lowerPriorityFont;
            if (lowerPriorityFont == null)
                return higherPriorityFont;
            if (higherPriorityFont == null && lowerPriorityFont == null)
                return null;

            rFonts = new XElement(W.rFonts,
                (higherPriorityFont.Attribute(W.ascii) != null || higherPriorityFont.Attribute(W.asciiTheme) != null) ?
                    new[] { higherPriorityFont.Attribute(W.ascii), higherPriorityFont.Attribute(W.asciiTheme) } :
                    new[] { lowerPriorityFont.Attribute(W.ascii), lowerPriorityFont.Attribute(W.asciiTheme) },
                (higherPriorityFont.Attribute(W.hAnsi) != null || higherPriorityFont.Attribute(W.hAnsiTheme) != null) ?
                    new[] { higherPriorityFont.Attribute(W.hAnsi), higherPriorityFont.Attribute(W.hAnsiTheme) } :
                    new[] { lowerPriorityFont.Attribute(W.hAnsi), lowerPriorityFont.Attribute(W.hAnsiTheme) },
                (higherPriorityFont.Attribute(W.eastAsia) != null || higherPriorityFont.Attribute(W.eastAsiaTheme) != null) ?
                    new[] { higherPriorityFont.Attribute(W.eastAsia), higherPriorityFont.Attribute(W.eastAsiaTheme) } :
                    new[] { lowerPriorityFont.Attribute(W.eastAsia), lowerPriorityFont.Attribute(W.eastAsiaTheme) },
                (higherPriorityFont.Attribute(W.cs) != null || higherPriorityFont.Attribute(W.cstheme) != null) ?
                    new[] { higherPriorityFont.Attribute(W.cs), higherPriorityFont.Attribute(W.cstheme) } :
                    new[] { lowerPriorityFont.Attribute(W.cs), lowerPriorityFont.Attribute(W.cstheme) },
                (higherPriorityFont.Attribute(W.hint) != null ? higherPriorityFont.Attribute(W.hint) :
                    lowerPriorityFont.Attribute(W.hint))
            );

            return rFonts;
        }

        private static int? CalcWidthOfRunInPixels(XElement r)
        {
            var fontName = (string)r.Attribute(PtOpenXml.FontName) ??
               (string)r.Ancestors(W.p).First().Attribute(PtOpenXml.FontName);
            if (fontName == null)
                throw new OpenXmlPowerToolsException("Internal Error, should have FontName attribute");
            if (UnknownFonts.Contains(fontName))
                return 0;

            if (UnknownFonts.Contains(fontName))
                return null;

            var rPr = r.Element(W.rPr);
            if (rPr == null)
                return null;

            var sz = GetFontSize(r) ?? 22m;

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
            if (Util.GetBoolProp(rPr, W.b) == true || Util.GetBoolProp(rPr, W.bCs) == true)
                fs |= FontStyle.Bold;
            if (Util.GetBoolProp(rPr, W.i) == true || Util.GetBoolProp(rPr, W.iCs) == true)
                fs |= FontStyle.Italic;

            // Appended blank as a quick fix to accommodate &nbsp; that will get
            // appended to some layout-critical runs such as list item numbers.
            // In some cases, this might not be required or even wrong, so this
            // must be revisited.
            // TODO: Revisit.
            var runText = r.DescendantsTrimmed(W.txbxContent)
                .Where(e => e.Name == W.t)
                .Select(t => (string)t)
                .StringConcatenate() + " ";

            var tabLength = r.DescendantsTrimmed(W.txbxContent)
                .Where(e => e.Name == W.tab)
                .Select(t => (decimal)t.Attribute(PtOpenXml.TabWidth))
                .Sum();

            if (runText.Length == 0 && tabLength == 0)
                return 0;

            int multiplier = 1;
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
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < multiplier; i++)
                    sb.Append(runText);
                runText = sb.ToString();
            }

            try
            {
                using (Font f = new Font(ff, (float)sz / 2f, fs))
                {
                    const TextFormatFlags tff = TextFormatFlags.NoPadding;
                    var proposedSize = new Size(int.MaxValue, int.MaxValue);
                    var sf = TextRenderer.MeasureText(runText, f, proposedSize, tff);
                    // sf returns size in pixels
                    return sf.Width / multiplier;
                }
            }
            catch (ArgumentException)
            {
                try
                {
                    const FontStyle fs2 = FontStyle.Regular;
                    using (Font f = new Font(ff, (float)sz / 2f, fs2))
                    {
                        const TextFormatFlags tff = TextFormatFlags.NoPadding;
                        var proposedSize = new Size(int.MaxValue, int.MaxValue);
                        var sf = TextRenderer.MeasureText(runText, f, proposedSize, tff);
                        return sf.Width / multiplier;
                    }
                }
                catch (ArgumentException)
                {
                    const FontStyle fs2 = FontStyle.Bold;
                    try
                    {
                        using (var f = new Font(ff, (float)sz / 2f, fs2))
                        {
                            const TextFormatFlags tff = TextFormatFlags.NoPadding;
                            var proposedSize = new Size(int.MaxValue, int.MaxValue);
                            var sf = TextRenderer.MeasureText(runText, f, proposedSize, tff);
                            // sf returns size in pixels
                            return sf.Width / multiplier;
                        }
                    }
                    catch (ArgumentException)
                    {
                        // if both regular and bold fail, then get metrics for Times New Roman
                        // use the original FontStyle (in fs)
                        var ff2 = new FontFamily("Times New Roman");
                        using (var f = new Font(ff2, (float)sz / 2f, fs))
                        {
                            const TextFormatFlags tff = TextFormatFlags.NoPadding;
                            var proposedSize = new Size(int.MaxValue, int.MaxValue);
                            var sf = TextRenderer.MeasureText(runText, f, proposedSize, tff);
                            // sf returns size in pixels
                            return sf.Width / multiplier;
                        }
                    }
                }
            }
            catch (OverflowException)
            {
                // This happened on Azure but interestingly enough not while testing locally.
                return 0;
            }
        }

        // The algorithm for this method comes from the implementer notes in [MS-OI29500].pdf
        // section 2.1.87

        // The implementer notes are at:
        // http://msdn.microsoft.com/en-us/library/ee908652.aspx

        public enum FontType
        {
            Ascii,
            HAnsi,
            EastAsia,
            CS
        };

        public class CharStyleAttributes
        {
            public string AsciiFont;
            public string HAnsiFont;
            public string EastAsiaFont;
            public string CsFont;
            public string Hint;
            public bool Rtl;

            public string LatinLang;
            public string BidiLang;
            public string EastAsiaLang;

            public Dictionary<XName, bool?> ToggleProperties;
            public Dictionary<XName, XElement> Properties;

            public CharStyleAttributes(XElement pPr, XElement rPr)
            {
                ToggleProperties = new Dictionary<XName, bool?>();
                Properties = new Dictionary<XName, XElement>();

                if (rPr == null)
                    return;
                foreach (XName xn in TogglePropertyNames)
                {
                    ToggleProperties[xn] = Util.GetBoolProp(rPr, xn);
                }
                foreach (XName xn in PropertyNames)
                {
                    Properties[xn] = GetXmlProperty(rPr, xn);
                }
                var rFonts = rPr.Element(W.rFonts);
                if (rFonts == null)
                {
                    this.AsciiFont = null;
                    this.HAnsiFont = null;
                    this.EastAsiaFont = null;
                    this.CsFont = null;
                    this.Hint = null;
                }
                else
                {
                    this.AsciiFont = (string)(rFonts.Attribute(W.ascii));
                    this.HAnsiFont = (string)(rFonts.Attribute(W.hAnsi));
                    this.EastAsiaFont = (string)(rFonts.Attribute(W.eastAsia));
                    this.CsFont = (string)(rFonts.Attribute(W.cs));
                    this.Hint = (string)(rFonts.Attribute(W.hint));
                }
                XElement csel = this.Properties[W.cs];
                bool cs = csel != null && (csel.Attribute(W.val) == null || csel.Attribute(W.val).ToBoolean() == true);
                XElement rtlel = this.Properties[W.rtl];
                bool rtl = rtlel != null && (rtlel.Attribute(W.val) == null || rtlel.Attribute(W.val).ToBoolean() == true);
                var bidi = false;
                if (pPr != null)
                {
                    XElement bidiel = pPr.Element(W.bidi);
                    bidi = bidiel != null && (bidiel.Attribute(W.val) == null || bidiel.Attribute(W.val).ToBoolean() == true);
                }
                Rtl = cs || rtl || bidi;
                var lang = rPr.Element(W.lang);
                if (lang != null)
                {
                    LatinLang = (string)lang.Attribute(W.val);
                    BidiLang = (string)lang.Attribute(W.bidi);
                    EastAsiaLang = (string)lang.Attribute(W.eastAsia);
                }
            }

            private static XElement GetXmlProperty(XElement rPr, XName propertyName)
            {
                return rPr.Element(propertyName);
            }

            private static XName[] TogglePropertyNames = new[] {
                W.b,
                W.bCs,
                W.caps,
                W.emboss,
                W.i,
                W.iCs,
                W.imprint,
                W.outline,
                W.shadow,
                W.smallCaps,
                W.strike,
                W.vanish
            };

            private static XName[] PropertyNames = new[] {
                W.cs,
                W.rtl,
                W.u,
                W.color,
                W.highlight,
                W.shd
            };

        }

        public static FontType DetermineFontTypeFromCharacter(char ch, CharStyleAttributes csa)
        {
            // If the run has the cs element ("[ISO/IEC-29500-1] §17.3.2.7; cs") or the rtl element ("[ISO/IEC-29500-1] §17.3.2.30; rtl"),
            // then the cs (or cstheme if defined) font is used, regardless of the Unicode character values of the run’s content.
            if (csa.Rtl)
            {
                return FontType.CS;
            }

            // A large percentage of characters will fall in the following rule.

            // Unicode Block: Basic Latin
            if (ch >= 0x00 && ch <= 0x7f)
            {
                return FontType.Ascii;
            }

            // If the eastAsia (or eastAsiaTheme if defined) attribute’s value is “Times New Roman” and the ascii (or asciiTheme if defined)
            // and hAnsi (or hAnsiTheme if defined) attributes are equal, then the ascii (or asciiTheme if defined) font is used.
            if (csa.EastAsiaFont == "Times New Roman" &&
                csa.AsciiFont == csa.HAnsiFont)
            {
                return FontType.Ascii;
            }

            // Unicode BLock: Latin-1 Supplement
            if (ch >= 0xA0 && ch <= 0xFF)
            {
                if (csa.Hint == "eastAsia")
                {
                    if (ch == 0xA1 ||
                        ch == 0xA4 ||
                        ch == 0xA7 ||
                        ch == 0xA8 ||
                        ch == 0xAA ||
                        ch == 0xAD ||
                        ch == 0xAF ||
                        (ch >= 0xB0 && ch <= 0xB4) ||
                        (ch >= 0xB6 && ch <= 0xBA) ||
                        (ch >= 0xBC && ch <= 0xBF) ||
                        ch == 0xD7 ||
                        ch == 0xF7)
                    {
                        return FontType.EastAsia;
                    }
                    if (csa.EastAsiaLang == "zh-hant" ||
                        csa.EastAsiaLang == "zh-hans")
                    {
                        if (ch == 0xE0 ||
                            ch == 0xE1 ||
                            (ch >= 0xE8 && ch <= 0xEA) ||
                            (ch >= 0xEC && ch <= 0xED) ||
                            (ch >= 0xF2 && ch <= 0xF3) ||
                            (ch >= 0xF9 && ch <= 0xFA) ||
                            ch == 0xFC)
                        {
                            return FontType.EastAsia;
                        }
                    }
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Latin Extended-A
            if (ch >= 0x0100 && ch <= 0x017F)
            {
                if (csa.Hint == "eastAsia")
                {
                    if (csa.EastAsiaLang == "zh-hant" ||
                        csa.EastAsiaLang == "zh-hans"
                        /* || the character set of the east Asia (or east Asia theme) font is Chinese5 || GB2312 todo */)
                    {
                        return FontType.EastAsia;
                    }
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Latin Extended-B
            if (ch >= 0x0180 && ch <= 0x024F)
            {
                if (csa.Hint == "eastAsia")
                {
                    if (csa.EastAsiaLang == "zh-hant" ||
                        csa.EastAsiaLang == "zh-hans"
                        /* || the character set of the east Asia (or east Asia theme) font is Chinese5 || GB2312 todo */)
                    {
                        return FontType.EastAsia;
                    }
                }
                return FontType.HAnsi;
            }

            // Unicode Block: IPA Extensions
            if (ch >= 0x0250 && ch <= 0x02AF)
            {
                if (csa.Hint == "eastAsia")
                {
                    if (csa.EastAsiaLang == "zh-hant" ||
                        csa.EastAsiaLang == "zh-hans"
                        /* || the character set of the east Asia (or east Asia theme) font is Chinese5 || GB2312 todo */)
                    {
                        return FontType.EastAsia;
                    }
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Spacing Modifier Letters
            if (ch >= 0x02B0 && ch <= 0x02FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Combining Diacritic Marks
            if (ch >= 0x0300 && ch <= 0x036F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Greek
            if (ch >= 0x0370 && ch <= 0x03CF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Cyrillic
            if (ch >= 0x0400 && ch <= 0x04FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Hebrew
            if (ch >= 0x0590 && ch <= 0x05FF)
            {
                return FontType.Ascii;
            }

            // Unicode Block: Arabic
            if (ch >= 0x0600 && ch <= 0x06FF)
            {
                return FontType.Ascii;
            }

            // Unicode Block: Syriac
            if (ch >= 0x0700 && ch <= 0x074F)
            {
                return FontType.Ascii;
            }

            // Unicode Block: Arabic Supplement
            if (ch >= 0x0750 && ch <= 0x077F)
            {
                return FontType.Ascii;
            }

            // Unicode Block: Thanna
            if (ch >= 0x0780 && ch <= 0x07BF)
            {
                return FontType.Ascii;
            }

            // Unicode Block: Hangul Jamo
            if (ch >= 0x1100 && ch <= 0x11FF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Latin Extended Additional
            if (ch >= 0x1E00 && ch <= 0x1EFF)
            {
                if (csa.Hint == "eastAsia" &&
                    (csa.EastAsiaLang == "zh-hant" ||
                    csa.EastAsiaLang == "zh-hans"))
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: General Punctuation
            if (ch >= 0x2000 && ch <= 0x206F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Superscripts and Subscripts
            if (ch >= 0x2070 && ch <= 0x209F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Currency Symbols
            if (ch >= 0x20A0 && ch <= 0x20CF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Combining Diacritical Marks for Symbols
            if (ch >= 0x20D0 && ch <= 0x20FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Letter-like Symbols
            if (ch >= 0x2100 && ch <= 0x214F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Number Forms
            if (ch >= 0x2150 && ch <= 0x218F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Arrows
            if (ch >= 0x2190 && ch <= 0x21FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Mathematical Operators
            if (ch >= 0x2200 && ch <= 0x22FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Miscellaneous Technical
            if (ch >= 0x2300 && ch <= 0x23FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Control Pictures
            if (ch >= 0x2400 && ch <= 0x243F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Optical Character Recognition
            if (ch >= 0x2440 && ch <= 0x245F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Enclosed Alphanumerics
            if (ch >= 0x2460 && ch <= 0x24FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Box Drawing
            if (ch >= 0x2500 && ch <= 0x257F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Block Elements
            if (ch >= 0x2580 && ch <= 0x259F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Geometric Shapes
            if (ch >= 0x25A0 && ch <= 0x25FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Miscellaneous Symbols
            if (ch >= 0x2600 && ch <= 0x26FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Dingbats
            if (ch >= 0x2700 && ch <= 0x27BF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: CJK Radicals Supplement
            if (ch >= 0x2E80 && ch <= 0x2EFF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Kangxi Radicals
            if (ch >= 0x2F00 && ch <= 0x2FDF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Ideographic Description Characters
            if (ch >= 0x2FF0 && ch <= 0x2FFF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: CJK Symbols and Punctuation
            if (ch >= 0x3000 && ch <= 0x303F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Hiragana
            if (ch >= 0x3040 && ch <= 0x309F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Katakana
            if (ch >= 0x30A0 && ch <= 0x30FF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Bopomofo
            if (ch >= 0x3100 && ch <= 0x312F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Hangul Compatibility Jamo
            if (ch >= 0x3130 && ch <= 0x318F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Kanbun
            if (ch >= 0x3190 && ch <= 0x319F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Enclosed CJK Letters and Months
            if (ch >= 0x3200 && ch <= 0x32FF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: CJK Compatibility
            if (ch >= 0x3300 && ch <= 0x33FF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: CJK Unified Ideographs Extension A
            if (ch >= 0x3400 && ch <= 0x4DBF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: CJK Unified Ideographs
            if (ch >= 0x4E00 && ch <= 0x9FAF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Yi Syllables
            if (ch >= 0xA000 && ch <= 0xA48F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Yi Radicals
            if (ch >= 0xA490 && ch <= 0xA4CF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Hangul Syllables
            if (ch >= 0xAC00 && ch <= 0xD7AF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: High Surrogates
            if (ch >= 0xD800 && ch <= 0xDB7F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: High Private Use Surrogates
            if (ch >= 0xDB80 && ch <= 0xDBFF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Low Surrogates
            if (ch >= 0xDC00 && ch <= 0xDFFF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Private Use Area
            if (ch >= 0xE000 && ch <= 0xF8FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: CJK Compatibility Ideographs
            if (ch >= 0xF900 && ch <= 0xFAFF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Alphabetic Presentation Forms
            if (ch >= 0xFB00 && ch <= 0xFB4F)
            {
                if (csa.Hint == "eastAsia")
                {
                    if (ch >= 0xFB00 && ch <= 0xFB1C)
                        return FontType.EastAsia;
                    if (ch >= 0xFB1D && ch <= 0xFB4F)
                        return FontType.Ascii;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Arabic Presentation Forms-A
            if (ch >= 0xFB50 && ch <= 0xFDFF)
            {
                return FontType.Ascii;
            }

            // Unicode Block: CJK Compatibility Forms
            if (ch >= 0xFE30 && ch <= 0xFE4F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Small Form Variants
            if (ch >= 0xFE50 && ch <= 0xFE6F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Arabic Presentation Forms-B
            if (ch >= 0xFE70 && ch <= 0xFEFE)
            {
                return FontType.Ascii;
            }

            // Unicode Block: Halfwidth and Fullwidth Forms
            if (ch >= 0xFF00 && ch <= 0xFFEF)
            {
                return FontType.EastAsia;
            }
            return FontType.HAnsi;
        }

        private static readonly HashSet<string> UnknownFonts = new HashSet<string>();
        private static HashSet<string> _knownFamilies;

        private static HashSet<string> KnownFamilies
        {
            get
            {
                if (_knownFamilies == null)
                {
                    _knownFamilies = new HashSet<string>();
                    var families = FontFamily.Families;
                    foreach (var fam in families)
                        _knownFamilies.Add(fam.Name);
                }
                return _knownFamilies;
            }
        }

        private static HashSet<char> WeakAndNeutralDirectionalCharacters = new HashSet<char>() {
            '0',
            '1',
            '2',
            '3',
            '4',
            '5',
            '6',
            '7',
            '8',
            '9',
            '+',
            '-',
            ':',
            ',',
            '.',
            '|',
            '\t',
            '\r',
            '\n',
            ' ',
            '\x00A0', // non breaking space

            '\x00B0', // degree sign
            '\x066B', // arabic decimal separator
            '\x066C', // arabic thousands separator

            '\x0627', // arabic pipe 

            '\x20A0', // start currency symbols
            '\x20A1',
            '\x20A2',
            '\x20A3',
            '\x20A4',
            '\x20A5',
            '\x20A6',
            '\x20A7',
            '\x20A8',
            '\x20A9',
            '\x20AA',
            '\x20AB',
            '\x20AC',
            '\x20AD',
            '\x20AE',
            '\x20AF',
            '\x20B0',
            '\x20B1',
            '\x20B2',
            '\x20B3',
            '\x20B4',
            '\x20B5',
            '\x20B6',
            '\x20B7',
            '\x20B8',
            '\x20B9',
            '\x20BA',
            '\x20BB',
            '\x20BC',
            '\x20BD',
            '\x20BE',
            '\x20BF',
            '\x20C0',
            '\x20C1',
            '\x20C2',
            '\x20C3',
            '\x20C4',
            '\x20C5',
            '\x20C6',
            '\x20C7',
            '\x20C8',
            '\x20C9',
            '\x20CA',
            '\x20CB',
            '\x20CC',
            '\x20CD',
            '\x20CE',
            '\x20CF',  // end currency symbols

            '\x0660', // "Arabic" Indic Numeral Forms Iraq and West
            '\x0661',
            '\x0662',
            '\x0663',
            '\x0664',
            '\x0665',
            '\x0666',
            '\x0667',
            '\x0668',
            '\x0669',

            '\x06F0', // "Arabic" Indic Numberal Forms Iran and East
            '\x06F1',
            '\x06F2',
            '\x06F3',
            '\x06F4',
            '\x06F5',
            '\x06F6',
            '\x06F7',
            '\x06F8',
            '\x06F9',
        };

        private static void AdjustFontAttributes(WordprocessingDocument wDoc, XElement paraOrRun, XElement pPr, XElement rPr)
        {
            XDocument themeXDoc = null;
            if (wDoc.MainDocumentPart.ThemePart != null)
                themeXDoc = wDoc.MainDocumentPart.ThemePart.GetXDocument();

            XElement fontScheme = null;
            XElement majorFont = null;
            XElement minorFont = null;
            if (themeXDoc != null)
            {
                fontScheme = themeXDoc.Root.Element(A.themeElements).Element(A.fontScheme);
                majorFont = fontScheme.Element(A.majorFont);
                minorFont = fontScheme.Element(A.minorFont);
            }
            var rFonts = rPr.Element(W.rFonts);
            if (rFonts == null)
            {
                return;
            }
            var asciiTheme = (string)rFonts.Attribute(W.asciiTheme);
            var hAnsiTheme = (string)rFonts.Attribute(W.hAnsiTheme);
            var eastAsiaTheme = (string)rFonts.Attribute(W.eastAsiaTheme);
            var cstheme = (string)rFonts.Attribute(W.cstheme);
            string ascii = null;
            string hAnsi = null;
            string eastAsia = null;
            string cs = null;

            XElement minorLatin = null;
            string minorLatinTypeface = null;
            XElement majorLatin = null;
            string majorLatinTypeface = null;

            if (minorFont != null)
            {
                minorLatin = minorFont.Element(A.latin);
                minorLatinTypeface = (string)minorLatin.Attribute("typeface");
            }

            if (majorFont != null)
            {
                majorLatin = majorFont.Element(A.latin);
                majorLatinTypeface = (string)majorLatin.Attribute("typeface");
            }
            if (asciiTheme != null)
            {
                if (asciiTheme.StartsWith("minor") && minorLatinTypeface != null)
                {
                    ascii = minorLatinTypeface;
                }
                else if (asciiTheme.StartsWith("major") && majorLatinTypeface != null)
                {
                    ascii = majorLatinTypeface;
                }
            }
            if (hAnsiTheme != null)
            {
                if (hAnsiTheme.StartsWith("minor") && minorLatinTypeface != null)
                {
                    hAnsi = minorLatinTypeface;
                }
                else if (hAnsiTheme.StartsWith("major") && majorLatinTypeface != null)
                {
                    hAnsi = majorLatinTypeface;
                }
            }
            if (eastAsiaTheme != null)
            {
                if (eastAsiaTheme.StartsWith("minor") && minorLatinTypeface != null)
                {
                    eastAsia = minorLatinTypeface;
                }
                else if (eastAsiaTheme.StartsWith("major") && majorLatinTypeface != null)
                {
                    eastAsia = majorLatinTypeface;
                }
            }
            if (cstheme != null)
            {
                if (cstheme.StartsWith("minor") && minorFont != null)
                {
                    cs = (string)minorFont.Element(A.cs).Attribute("typeface");
                }
                else if (cstheme.StartsWith("major") && majorFont != null)
                {
                    cs = (string)majorFont.Element(A.cs).Attribute("typeface");
                }
            }

            if (ascii != null)
            {
                rFonts.SetAttributeValue(W.ascii, ascii);
            }
            if (hAnsi != null)
            {
                rFonts.SetAttributeValue(W.hAnsi, hAnsi);
            }
            if (eastAsia != null)
            {
                rFonts.SetAttributeValue(W.eastAsia, eastAsia);
            }
            if (cs != null)
            {
                rFonts.SetAttributeValue(W.cs, cs);
            }

            var firstTextNode = paraOrRun.Descendants(W.t).FirstOrDefault(t => t.Value.Length > 0);
            string str = " ";

            // if there is a run with no text in it, then no need to do any of the rest of this method.
            if (firstTextNode == null && paraOrRun.Name == W.r)
                return;

            if (firstTextNode != null)
                str = firstTextNode.Value;

            var csa = new CharStyleAttributes(pPr, rPr);

            // This module determines the font based on just the first character.
            // Technically, a run can contain characters from different Unicode code blocks, and hence should be rendered with different fonts.
            // However, Word breaks up runs that use more than one font into multiple runs.  Other producers of WordprocessingML may not, so in
            // that case, this routine may need to be augmented to look at all characters in a run.

            /*
            old code
            var fontFamilies = str.select(function (c) {
                var ft = Pav.DetermineFontTypeFromCharacter(c, csa);
                switch (ft) {
                    case Pav.FontType.Ascii:
                        return cast(rFonts.attribute(W.ascii));
                    case Pav.FontType.HAnsi:
                        return cast(rFonts.attribute(W.hAnsi));
                    case Pav.FontType.EastAsia:
                        return cast(rFonts.attribute(W.eastAsia));
                    case Pav.FontType.CS:
                        return cast(rFonts.attribute(W.cs));
                    default:
                        return null;
                }
            })
                .where(function (f) { return f != null && f != ""; })
                .distinct()
                .select(function (f) { return new Pav.FontFamily(f); })
                .toArray();
            */

            var charToExamine = str.FirstOrDefault(c => !WeakAndNeutralDirectionalCharacters.Contains(c));
            if (charToExamine == '\0')
                charToExamine = str[0];

            var ft = DetermineFontTypeFromCharacter(charToExamine, csa);
            string fontType = null;
            string languageType = null;
            switch (ft)
            {
                case FontType.Ascii:
                    fontType = (string)rFonts.Attribute(W.ascii);
                    languageType = "western";
                    break;
                case FontType.HAnsi:
                    fontType = (string)rFonts.Attribute(W.hAnsi);
                    languageType = "western";
                    break;
                case FontType.EastAsia:
                    fontType = (string)rFonts.Attribute(W.eastAsia);
                    languageType = "eastAsia";
                    break;
                case FontType.CS:
                    fontType = (string)rFonts.Attribute(W.cs);
                    languageType = "bidi";
                    break;
            }

            if (fontType != null)
            {
                if (paraOrRun.Attribute(PtOpenXml.FontName) == null)
                {
                    XAttribute fta = new XAttribute(PtOpenXml.FontName, fontType.ToString());
                    paraOrRun.Add(fta);
                }
                else
                {
                    paraOrRun.Attribute(PtOpenXml.FontName).Value = fontType.ToString();
                }
            }
            if (languageType != null)
            {
                if (paraOrRun.Attribute(PtOpenXml.LanguageType) == null)
                {
                    XAttribute lta = new XAttribute(PtOpenXml.LanguageType, languageType);
                    paraOrRun.Add(lta);
                }
                else
                {
                    paraOrRun.Attribute(PtOpenXml.LanguageType).Value = languageType;
                }
            }
        }

        private static decimal? GetFontSize(XElement e)
        {
            var languageType = (string)e.Attribute(PtOpenXml.LanguageType);
            if (e.Name == W.p)
            {
                return GetFontSize(languageType, e.Elements(W.pPr).Elements(W.rPr).FirstOrDefault());
            }
            if (e.Name == W.r)
            {
                return GetFontSize(languageType, e.Element(W.rPr));
            }
            return null;
        }

        private static decimal? GetFontSize(string languageType, XElement rPr)
        {
            if (rPr == null) return null;
            return languageType == "bidi"
                ? (decimal?)rPr.Elements(W.szCs).Attributes(W.val).FirstOrDefault()
                : (decimal?)rPr.Elements(W.sz).Attributes(W.val).FirstOrDefault();
        }

        private static int NextRectId = 1025;

        private static int GetNextRectId()
        {
            return NextRectId++;
        }

        private static object GenerateNextExpected(XNode node, HtmlToWmlConverterSettings settings, WordprocessingDocument wDoc,
            string styleName, NextExpected nextExpected, bool preserveWhiteSpace)
        {
            if (nextExpected == NextExpected.Paragraph)
            {
                XElement element = node as XElement;
                if (element != null)
                {
                    return new XElement(W.p,
                        GetParagraphProperties(element, styleName, settings),
                        element.Nodes().Select(n => Transform(n, settings, wDoc, NextExpected.Run, preserveWhiteSpace)));
                }
                else
                {
                    XText xTextNode = node as XText;
                    if (xTextNode != null)
                    {
                        string textNodeString = GetDisplayText(xTextNode, preserveWhiteSpace);
                        XElement p;
                        p = new XElement(W.p,
                            GetParagraphProperties(node.Parent, null, settings),
                            new XElement(W.r,
                                GetRunProperties((XText)node, settings),
                                new XElement(W.t,
                                    GetXmlSpaceAttribute(textNodeString),
                                    textNodeString)));
                        return p;
                    }
                    return null;
                }
            }
            else
            {
                XElement element = node as XElement;
                if (element != null)
                {
                    return element.Nodes().Select(n => Transform(n, settings, wDoc, nextExpected, preserveWhiteSpace));
                }
                else
                {
                    string textNodeString = GetDisplayText((XText)node, preserveWhiteSpace);
                    XElement rPr = GetRunProperties((XText)node, settings);
                    XElement r = new XElement(W.r,
                        rPr,
                        new XElement(W.t,
                            GetXmlSpaceAttribute(textNodeString),
                            textNodeString));
                    return r;
                }
            }
        }

        private static XElement TransformImageToWml(XElement element, HtmlToWmlConverterSettings settings, WordprocessingDocument wDoc)
        {
            string imageName = (string)element.Attribute(XhtmlNoNamespace.src);
            Bitmap bmp;
            try
            {
                bmp = new Bitmap(settings.BaseUriForImages + "/" + imageName);
            }
            catch (ArgumentException)
            {
                return null;
            }
            catch (NotSupportedException)
            {
                return null;
            }
            MemoryStream ms = new MemoryStream();
            bmp.Save(ms, bmp.RawFormat);
            byte[] ba = ms.ToArray();
            MainDocumentPart mdp = wDoc.MainDocumentPart;
            string rId = "R" + Guid.NewGuid().ToString().Replace("-", "");
            ImagePartType ipt = ImagePartType.Png;
            ImagePart newPart = mdp.AddImagePart(ipt, rId);
            using (Stream s = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                s.Write(ba, 0, ba.GetUpperBound(0) + 1);

            PictureId pid = wDoc.Annotation<PictureId>();
            if (pid == null)
            {
                pid = new PictureId
                {
                    Id = 1,
                };
                wDoc.AddAnnotation(pid);
            }
            int pictureId = pid.Id;
            ++pid.Id;

            string pictureDescription = "Picture " + pictureId.ToString();

            string floatValue = element.GetProp("float").ToString();
            if (floatValue == "none")
            {
                XElement run = new XElement(W.r,
                    GetRunPropertiesForImage(),
                    new XElement(W.drawing,
                        GetImageAsInline(element, settings, wDoc, bmp, rId, pictureId, pictureDescription)));
                return run;
            }
            if (floatValue == "left" || floatValue == "right")
            {
                XElement run = new XElement(W.r,
                    GetRunPropertiesForImage(),
                    new XElement(W.drawing,
                        GetImageAsAnchor(element, settings, wDoc, bmp, rId, floatValue, pictureId, pictureDescription)));
                return run;
            }
            return null;
        }

        private static XElement GetImageAsInline(XElement element, HtmlToWmlConverterSettings settings, WordprocessingDocument wDoc, Bitmap bmp,
            string rId, int pictureId, string pictureDescription)
        {
            XElement inline = new XElement(WP.inline, // 20.4.2.8
                new XAttribute(XNamespace.Xmlns + "wp", WP.wp.NamespaceName),
                new XAttribute(NoNamespace.distT, 0),  // distance from top of image to text, in EMUs, no effect if the parent is inline
                new XAttribute(NoNamespace.distB, 0),  // bottom
                new XAttribute(NoNamespace.distL, 0),  // left
                new XAttribute(NoNamespace.distR, 0),  // right
                GetImageExtent(element, bmp),
                GetEffectExtent(),
                GetDocPr(element, pictureId, pictureDescription),
                GetCNvGraphicFramePr(),
                GetGraphicForImage(element, rId, bmp, pictureId, pictureDescription));
            return inline;
        }

        private static XElement GetImageAsAnchor(XElement element, HtmlToWmlConverterSettings settings, WordprocessingDocument wDoc, Bitmap bmp,
            string rId, string floatValue, int pictureId, string pictureDescription)
        {
            Emu minDistFromEdge = (long)(0.125 * Emu.s_EmusPerInch);
            long relHeight = 251658240;  // z-order

            CssExpression marginTopProp = element.GetProp("margin-top");
            CssExpression marginLeftProp = element.GetProp("margin-left");
            CssExpression marginBottomProp = element.GetProp("margin-bottom");
            CssExpression marginRightProp = element.GetProp("margin-right");

            Emu marginTopInEmus = 0;
            Emu marginBottomInEmus = 0;
            Emu marginLeftInEmus = 0;
            Emu marginRightInEmus = 0;

            if (marginTopProp.IsNotAuto)
                marginTopInEmus = (Emu)marginTopProp;

            if (marginBottomProp.IsNotAuto)
                marginBottomInEmus = (Emu)marginBottomProp;

            if (marginLeftProp.IsNotAuto)
                marginLeftInEmus = (Emu)marginLeftProp;

            if (marginRightProp.IsNotAuto)
                marginRightInEmus = (Emu)marginRightProp;

            Emu relativeFromColumn = 0;
            if (floatValue == "left")
            {
                relativeFromColumn = marginLeftInEmus;
                CssExpression parentMarginLeft = element.Parent.GetProp("margin-left");
                if (parentMarginLeft.IsNotAuto)
                    relativeFromColumn += (long)(Emu)parentMarginLeft;
                marginRightInEmus = Math.Max(marginRightInEmus, minDistFromEdge);
            }
            else if (floatValue == "right")
            {
                Emu printWidth = (long)settings.PageWidthEmus - (long)settings.PageMarginLeftEmus - (long)settings.PageMarginRightEmus;
                SizeEmu sl = GetImageSizeInEmus(element, bmp);
                relativeFromColumn = printWidth - sl.m_Width;
                if (marginRightProp.IsNotAuto)
                    relativeFromColumn -= (long)(Emu)marginRightInEmus;
                CssExpression parentMarginRight = element.Parent.GetProp("margin-right");
                if (parentMarginRight.IsNotAuto)
                    relativeFromColumn -= (long)(Emu)parentMarginRight;
                marginLeftInEmus = Math.Max(marginLeftInEmus, minDistFromEdge);
            }

            Emu relativeFromParagraph = marginTopInEmus;
            CssExpression parentMarginTop = element.Parent.GetProp("margin-top");
            if (parentMarginTop.IsNotAuto)
                relativeFromParagraph += (long)(Emu)parentMarginTop;

            XElement anchor = new XElement(WP.anchor,
                new XAttribute(XNamespace.Xmlns + "wp", WP.wp.NamespaceName),
                new XAttribute(NoNamespace.distT, (long)marginTopInEmus),     // distance from top of image to text, in EMUs, no effect if the parent is inline
                new XAttribute(NoNamespace.distB, (long)marginBottomInEmus),  // bottom
                new XAttribute(NoNamespace.distL, (long)marginLeftInEmus),    // left
                new XAttribute(NoNamespace.distR, (long)marginRightInEmus),   // right
                new XAttribute(NoNamespace.simplePos, 0),
                new XAttribute(NoNamespace.relativeHeight, relHeight),
                new XAttribute(NoNamespace.behindDoc, 0),
                new XAttribute(NoNamespace.locked, 0),
                new XAttribute(NoNamespace.layoutInCell, 1),
                new XAttribute(NoNamespace.allowOverlap, 1),
                new XElement(WP.simplePos, new XAttribute(NoNamespace.x, 0), new XAttribute(NoNamespace.y, 0)),
                new XElement(WP.positionH, new XAttribute(NoNamespace.relativeFrom, "column"),
                    new XElement(WP.posOffset, (long)relativeFromColumn)),
                new XElement(WP.positionV, new XAttribute(NoNamespace.relativeFrom, "paragraph"),
                    new XElement(WP.posOffset, (long)relativeFromParagraph)),
                GetImageExtent(element, bmp),
                GetEffectExtent(),
                new XElement(WP.wrapSquare, new XAttribute(NoNamespace.wrapText, "bothSides")),
                GetDocPr(element, pictureId, pictureDescription),
                GetCNvGraphicFramePr(),
                GetGraphicForImage(element, rId, bmp, pictureId, pictureDescription),
                new XElement(WP14.sizeRelH, new XAttribute(NoNamespace.relativeFrom, "page"),
                    new XElement(WP14.pctWidth, 0)),
                new XElement(WP14.sizeRelV, new XAttribute(NoNamespace.relativeFrom, "page"),
                    new XElement(WP14.pctHeight, 0))
            );
            return anchor;
        }
#if false
          <wp:anchor distT="0"
                     distB="0"
                     distL="114300"
                     distR="114300"
                     simplePos="0"
                     relativeHeight="251658240"
                     behindDoc="0"
                     locked="0"
                     layoutInCell="1"
                     allowOverlap="1">
            <wp:simplePos x="0"
                          y="0"/>
            <wp:positionH relativeFrom="column">
              <wp:posOffset>0</wp:posOffset>
            </wp:positionH>
            <wp:positionV relativeFrom="paragraph">
              <wp:posOffset>0</wp:posOffset>
            </wp:positionV>
            <wp:extent cx="1713865"
                       cy="1656715"/>
            <wp:effectExtent l="0"
                             t="0"
                             r="635"
                             b="635"/>
            <wp:wrapSquare wrapText="bothSides"/>
            <wp:docPr id="1"
                      name="Picture 1"
                      descr="img.png"/>
            <wp:cNvGraphicFramePr>
              <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                                   noChangeAspect="1"/>
            </wp:cNvGraphicFramePr>
            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:nvPicPr>
                    <pic:cNvPr id="0"
                               name="Picture 1"
                               descr="img.png"/>
                    <pic:cNvPicPr>
                      <a:picLocks noChangeAspect="1"
                                  noChangeArrowheads="1"/>
                    </pic:cNvPicPr>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="rId5">
                      <a:extLst>
                        <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                          <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
                                           val="0"/>
                        </a:ext>
                      </a:extLst>
                    </a:blip>
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </pic:blipFill>
                  <pic:spPr bwMode="auto">
                    <a:xfrm>
                      <a:off x="0"
                             y="0"/>
                      <a:ext cx="1713865"
                             cy="1656715"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                      <a:avLst/>
                    </a:prstGeom>
                    <a:noFill/>
                    <a:ln>
                      <a:noFill/>
                    </a:ln>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
            <wp14:sizeRelH relativeFrom="page">
              <wp14:pctWidth>0</wp14:pctWidth>
            </wp14:sizeRelH>
            <wp14:sizeRelV relativeFrom="page">
              <wp14:pctHeight>0</wp14:pctHeight>
            </wp14:sizeRelV>
          </wp:anchor>
#endif

        private static XElement GetParagraphPropertiesForImage()
        {
            return null;
        }

        private static XElement GetRunPropertiesForImage()
        {
            return new XElement(W.rPr,
                new XElement(W.noProof));
        }

        private static SizeEmu GetImageSizeInEmus(XElement img, Bitmap bmp)
        {
            double hres = bmp.HorizontalResolution;
            double vres = bmp.VerticalResolution;
            Size s = bmp.Size;
            Emu cx = (long)((double)(s.Width / hres) * (double)Emu.s_EmusPerInch);
            Emu cy = (long)((double)(s.Height / vres) * (double)Emu.s_EmusPerInch);

            CssExpression width = img.GetProp("width");
            CssExpression height = img.GetProp("height");
            if (width.IsNotAuto && height.IsAuto)
            {
                Emu widthInEmus = (Emu)width;
                double percentChange = widthInEmus / cx;
                cx = widthInEmus;
                cy = (long)(cy * percentChange);
                return new SizeEmu(cx, cy);
            }
            if (width.IsAuto && height.IsNotAuto)
            {
                Emu heightInEmus = (Emu)height;
                double percentChange = (float)heightInEmus / (float)cy;
                cy = heightInEmus;
                cx = (long)(cx * percentChange);
                return new SizeEmu(cx, cy);
            }
            if (width.IsNotAuto && height.IsNotAuto)
            {
                return new SizeEmu((Emu)width, (Emu)height);
            }
            return new SizeEmu(cx, cy);
        }

        private static XElement GetImageExtent(XElement img, Bitmap bmp)
        {
            SizeEmu szEmu = GetImageSizeInEmus(img, bmp);
            return new XElement(WP.extent,
                new XAttribute(NoNamespace.cx, (long)szEmu.m_Width),   // in EMUs
                new XAttribute(NoNamespace.cy, (long)szEmu.m_Height)); // in EMUs
        }

        private static XElement GetEffectExtent()
        {
            return new XElement(WP.effectExtent,
                new XAttribute(NoNamespace.l, 0),
                new XAttribute(NoNamespace.t, 0),
                new XAttribute(NoNamespace.r, 0),
                new XAttribute(NoNamespace.b, 0));
        }

        private static XElement GetDocPr(XElement element, int pictureId, string pictureDescription)
        {
            return new XElement(WP.docPr,
                new XAttribute(NoNamespace.id, pictureId),
                new XAttribute(NoNamespace.name, pictureDescription),
                new XAttribute(NoNamespace.descr, (string)element.Attribute(NoNamespace.src)));
        }

        private static XElement GetCNvGraphicFramePr()
        {
            return new XElement(WP.cNvGraphicFramePr,
                new XElement(A.graphicFrameLocks,
                    new XAttribute(XNamespace.Xmlns + "a", A.a.NamespaceName),
                    new XAttribute(NoNamespace.noChangeAspect, 1)));
        }

        private static XElement GetGraphicForImage(XElement element, string rId, Bitmap bmp, int pictureId, string pictureDescription)
        {
            SizeEmu szEmu = GetImageSizeInEmus(element, bmp);
            XElement graphic = new XElement(A.graphic,
                new XAttribute(XNamespace.Xmlns + "a", A.a.NamespaceName),
                new XElement(A.graphicData,
                    new XAttribute(NoNamespace.uri, Pic.pic.NamespaceName),
                    new XElement(Pic._pic,
                        new XAttribute(XNamespace.Xmlns + "pic", Pic.pic.NamespaceName),
                        new XElement(Pic.nvPicPr,
                            new XElement(Pic.cNvPr,
                                new XAttribute(NoNamespace.id, pictureId),
                                new XAttribute(NoNamespace.name, pictureDescription),
                                new XAttribute(NoNamespace.descr, (string)element.Attribute(NoNamespace.src))),
                            new XElement(Pic.cNvPicPr,
                                new XElement(A.picLocks,
                                    new XAttribute(NoNamespace.noChangeAspect, 1),
                                    new XAttribute(NoNamespace.noChangeArrowheads, 1)))),
                        new XElement(Pic.blipFill,
                            new XElement(A.blip,
                                new XAttribute(R.embed, rId),
                                new XElement(A.extLst,
                                    new XElement(A.ext,
                                        new XAttribute(NoNamespace.uri, "{28A0092B-C50C-407E-A947-70E740481C1C}"),
                                        new XElement(A14.useLocalDpi,
                                            new XAttribute(NoNamespace.val, "0"))))),
                            new XElement(A.stretch,
                                new XElement(A.fillRect))),
                        new XElement(Pic.spPr,
                            new XAttribute(NoNamespace.bwMode, "auto"),
                            new XElement(A.xfrm,
                                new XElement(A.off, new XAttribute(NoNamespace.x, 0), new XAttribute(NoNamespace.y, 0)),
                                new XElement(A.ext, new XAttribute(NoNamespace.cx, (long)szEmu.m_Width), new XAttribute(NoNamespace.cy, (long)szEmu.m_Height))),
                            new XElement(A.prstGeom, new XAttribute(NoNamespace.prst, "rect"),
                                new XElement(A.avLst)),
                            new XElement(A.noFill),
                            new XElement(A.ln,
                                new XElement(A.noFill))))));
            return graphic;
        }

#if false
            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:nvPicPr>
                    <pic:cNvPr id="0" name="Picture 1" descr="img.png"/>
                    <pic:cNvPicPr>
                      <a:picLocks noChangeAspect="1" noChangeArrowheads="1"/>
                    </pic:cNvPicPr>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:link="rId5">
                      <a:extLst>
                        <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                          <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
                        </a:ext>
                      </a:extLst>
                    </a:blip>
                    <a:srcRect/>
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </pic:blipFill>
                  <pic:spPr bwMode="auto">
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="1781175" cy="1781175"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                      <a:avLst/>
                    </a:prstGeom>
                    <a:noFill/>
                    <a:ln>
                      <a:noFill/>
                    </a:ln>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
#endif
        private static XElement GetParagraphProperties(XElement blockLevelElement, string styleName, HtmlToWmlConverterSettings settings)
        {
            XElement paragraphMarkRunProperties = GetRunProperties(blockLevelElement, settings);
            XElement backgroundProperty = GetBackgroundProperty(blockLevelElement);
            XElement[] spacingProperty = GetSpacingProperties(blockLevelElement, settings); // spacing, ind, contextualSpacing
            XElement jc = GetJustification(blockLevelElement, settings);
            XElement pStyle = styleName != null ? new XElement(W.pStyle, new XAttribute(W.val, styleName)) : null;
            XElement numPr = GetNumberingProperties(blockLevelElement, settings);
            XElement pBdr = GetBlockContentBorders(blockLevelElement, W.pBdr, true);

            XElement bidi = null;
            string direction = GetDirection(blockLevelElement);
            if (direction == "rtl")
                bidi = new XElement(W.bidi);

            XElement pPr = new XElement(W.pPr,
                pStyle,
                numPr,
                pBdr,
                backgroundProperty,
                bidi,
                spacingProperty,
                jc,
                paragraphMarkRunProperties
            );
            return pPr;
        }

        // vertical-align doesn't really work in the Word rendering - puts space above, but not below.  There really are no
        // options in WordprocessingML to specify vertical alignment.  I think that the only possible way that this could be
        // implemented would be to specifically calculate the space before and space after.  I'm not completely sure that
        // this could be possible.  I am pretty sure that this is not worth the effort.

        // Returns the spacing, ind, and contextualSpacing elements
        private static XElement[] GetSpacingProperties(XElement paragraph, HtmlToWmlConverterSettings settings)
        {
            CssExpression marginLeftProperty = paragraph.GetProp("margin-left");
            CssExpression marginRightProperty = paragraph.GetProp("margin-right");
            CssExpression marginTopProperty = paragraph.GetProp("margin-top");
            CssExpression marginBottomProperty = paragraph.GetProp("margin-bottom");
            CssExpression lineHeightProperty = paragraph.GetProp("line-height");
            CssExpression leftPaddingProperty = paragraph.GetProp("padding-left");
            CssExpression rightPaddingProperty = paragraph.GetProp("padding-right");

            /*****************************************************************************************/
            // leftIndent, rightIndent, firstLine

            Twip leftIndent = 0;
            Twip rightIndent = 0;
            Twip firstLine = 0;

#if false
            // this code is here for some reason.  What is it?
            double leftBorderSize = GetBorderSize(paragraph, "left"); // in 1/8 point
            double rightBorderSize = GetBorderSize(paragraph, "right"); // in 1/8 point
            leftIndent += (long)((leftBorderSize / 8d) * 20d);
            rightIndent += (long)((rightBorderSize / 8d) * 20d);
#endif

            if (leftPaddingProperty != null)
                leftIndent += (Twip)leftPaddingProperty;
            if (rightPaddingProperty != null)
                rightIndent += (Twip)rightPaddingProperty;

            if (paragraph.Name == XhtmlNoNamespace.li)
            {
                leftIndent += 180;
                rightIndent += 180;
            }
            XElement listElement = null;
            NumberedItemAnnotation numberedItemAnnotation = null;
            listElement = paragraph.Ancestors().FirstOrDefault(a => a.Name == XhtmlNoNamespace.ol || a.Name == XhtmlNoNamespace.ul);
            if (listElement != null)
            {
                numberedItemAnnotation = listElement.Annotation<NumberedItemAnnotation>();
                leftIndent += 600 * (numberedItemAnnotation.ilvl + 1);
            }

            int blockQuoteCount = paragraph.Ancestors(XhtmlNoNamespace.blockquote).Count();
            leftIndent += blockQuoteCount * 720;
            if (blockQuoteCount == 0)
            {
                if (marginLeftProperty != null && marginLeftProperty.IsNotAuto && marginLeftProperty.IsNotNormal)
                    leftIndent += (Twip)marginLeftProperty;
                if (marginRightProperty != null && marginRightProperty.IsNotAuto && marginRightProperty.IsNotNormal)
                    rightIndent += (Twip)marginRightProperty;
            }
            CssExpression textIndentProperty = paragraph.GetProp("text-indent");
            if (textIndentProperty != null)
            {
                Twip twips = (Twip)textIndentProperty;
                firstLine = twips;
            }

            XElement ind = null;
            if (leftIndent > 0 || rightIndent > 0 || firstLine != 0)
            {
                if (firstLine < 0)
                    ind = new XElement(W.ind,
                        leftIndent != 0 ? new XAttribute(W.left, (long)leftIndent) : null,
                        rightIndent != 0 ? new XAttribute(W.right, (long)rightIndent) : null,
                        firstLine != 0 ? new XAttribute(W.hanging, -(long)firstLine) : null);
                else
                    ind = new XElement(W.ind,
                        leftIndent != 0 ? new XAttribute(W.left, (long)leftIndent) : null,
                        rightIndent != 0 ? new XAttribute(W.right, (long)rightIndent) : null,
                        firstLine != 0 ? new XAttribute(W.firstLine, (long)firstLine) : null);
            }

            /*****************************************************************************************/
            // spacing

            long line = 240;
            string lineRule = "auto";
            string beforeAutospacing = null;
            string afterAutospacing = null;
            long? before = null;
            long? after = null;

            if (paragraph.Name == XhtmlNoNamespace.td || paragraph.Name == XhtmlNoNamespace.th || paragraph.Name == XhtmlNoNamespace.caption)
            {
                line = (long)settings.DefaultSpacingElementForParagraphsInTables.Attribute(W.line);
                lineRule = (string)settings.DefaultSpacingElementForParagraphsInTables.Attribute(W.lineRule);
                before = (long?)settings.DefaultSpacingElementForParagraphsInTables.Attribute(W.before);
                beforeAutospacing = (string)settings.DefaultSpacingElementForParagraphsInTables.Attribute(W.beforeAutospacing);
                after = (long?)settings.DefaultSpacingElementForParagraphsInTables.Attribute(W.after);
                afterAutospacing = (string)settings.DefaultSpacingElementForParagraphsInTables.Attribute(W.afterAutospacing);
            }

            // todo should check based on display property
            bool numItem = paragraph.Name == XhtmlNoNamespace.li;

            if (numItem && marginTopProperty.IsAuto)
                beforeAutospacing = "1";
            if (numItem && marginBottomProperty.IsAuto)
                afterAutospacing = "1";
            if (marginTopProperty != null && marginTopProperty.IsNotAuto)
            {
                before = (long)(Twip)marginTopProperty;
                beforeAutospacing = "0";
            }
            if (marginBottomProperty != null && marginBottomProperty.IsNotAuto)
            {
                after = (long)(Twip)marginBottomProperty;
                afterAutospacing = "0";
            }
            if (lineHeightProperty != null && lineHeightProperty.IsNotAuto && lineHeightProperty.IsNotNormal)
            {
                // line is in twips if lineRule == "atLeast"
                line = (long)(Twip)lineHeightProperty;
                lineRule = "atLeast";
            }

            XElement spacing = new XElement(W.spacing,
                before != null ? new XAttribute(W.before, before) : null,
                beforeAutospacing != null ? new XAttribute(W.beforeAutospacing, beforeAutospacing) : null,
                after != null ? new XAttribute(W.after, after) : null,
                afterAutospacing != null ? new XAttribute(W.afterAutospacing, afterAutospacing) : null,
                new XAttribute(W.line, line),
                new XAttribute(W.lineRule, lineRule));

            /*****************************************************************************************/
            // contextualSpacing

            XElement contextualSpacing = null;
            if (paragraph.Name == XhtmlNoNamespace.li)
            {
                NumberedItemAnnotation thisNumberedItemAnnotation = null;
                XElement listElement2 = paragraph.Ancestors().FirstOrDefault(a => a.Name == XhtmlNoNamespace.ol || a.Name == XhtmlNoNamespace.ul);
                if (listElement2 != null)
                {
                    thisNumberedItemAnnotation = listElement2.Annotation<NumberedItemAnnotation>();
                    XElement next = paragraph.ElementsAfterSelf().FirstOrDefault();
                    if (next != null && next.Name == XhtmlNoNamespace.li)
                    {
                        XElement nextListElement = next.Ancestors().FirstOrDefault(a => a.Name == XhtmlNoNamespace.ol || a.Name == XhtmlNoNamespace.ul);
                        NumberedItemAnnotation nextNumberedItemAnnotation = nextListElement.Annotation<NumberedItemAnnotation>();
                        if (nextNumberedItemAnnotation != null && thisNumberedItemAnnotation.numId == nextNumberedItemAnnotation.numId)
                            contextualSpacing = new XElement(W.contextualSpacing);
                    }
                }
            }

            return new XElement[] { spacing, ind, contextualSpacing };
        }

        private static XElement GetRunProperties(XText textNode, HtmlToWmlConverterSettings settings)
        {
            XElement parent = textNode.Parent;
            XElement rPr = GetRunProperties(parent, settings);
            return rPr;
        }

        private static XElement GetRunProperties(XElement element, HtmlToWmlConverterSettings settings)
        {
            CssExpression colorProperty = element.GetProp("color");
            CssExpression fontFamilyProperty = element.GetProp("font-family");
            CssExpression fontSizeProperty = element.GetProp("font-size");
            CssExpression textDecorationProperty = element.GetProp("text-decoration");
            CssExpression fontStyleProperty = element.GetProp("font-style");
            CssExpression fontWeightProperty = element.GetProp("font-weight");
            CssExpression backgroundColorProperty = element.GetProp("background-color");
            CssExpression letterSpacingProperty = element.GetProp("letter-spacing");
            CssExpression directionProp = element.GetProp("direction");

            string colorPropertyString = colorProperty.ToString();
            string fontFamilyString = GetUsedFontFromFontFamilyProperty(fontFamilyProperty);
            TPoint? fontSizeTPoint = GetUsedSizeFromFontSizeProperty(fontSizeProperty);
            string textDecorationString = textDecorationProperty.ToString();
            string fontStyleString = fontStyleProperty.ToString();
            string fontWeightString = fontWeightProperty.ToString().ToLower();
            string backgroundColorString = backgroundColorProperty.ToString().ToLower();
            string letterSpacingString = letterSpacingProperty.ToString().ToLower();
            string directionString = directionProp.ToString().ToLower();

            bool subAncestor = element.AncestorsAndSelf(XhtmlNoNamespace.sub).Any();
            bool supAncestor = element.AncestorsAndSelf(XhtmlNoNamespace.sup).Any();
            bool bAncestor = element.AncestorsAndSelf(XhtmlNoNamespace.b).Any();
            bool iAncestor = element.AncestorsAndSelf(XhtmlNoNamespace.i).Any();
            bool strongAncestor = element.AncestorsAndSelf(XhtmlNoNamespace.strong).Any();
            bool emAncestor = element.AncestorsAndSelf(XhtmlNoNamespace.em).Any();
            bool uAncestor = element.AncestorsAndSelf(XhtmlNoNamespace.u).Any();
            bool sAncestor = element.AncestorsAndSelf(XhtmlNoNamespace.s).Any();

            XAttribute dirAttribute = element.Attribute(XhtmlNoNamespace.dir);
            string dirAttributeString = "";
            if (dirAttribute != null)
                dirAttributeString = dirAttribute.Value.ToLower();

            XElement shd = null;
            if (backgroundColorString != "transparent")
                shd = new XElement(W.shd, new XAttribute(W.val, "clear"),
                    new XAttribute(W.color, "auto"),
                    new XAttribute(W.fill, backgroundColorString));

            XElement subSuper = null;
            if (subAncestor)
                subSuper = new XElement(W.vertAlign, new XAttribute(W.val, "subscript"));
            else
                if (supAncestor)
                    subSuper = new XElement(W.vertAlign, new XAttribute(W.val, "superscript"));

            XElement rFonts = null;
            if (fontFamilyString != null)
            {
                rFonts = new XElement(W.rFonts,
                    fontFamilyString != settings.MinorLatinFont ? new XAttribute(W.ascii, fontFamilyString) : null,
                    fontFamilyString != settings.MajorLatinFont ? new XAttribute(W.hAnsi, fontFamilyString) : null,
                    new XAttribute(W.cs, fontFamilyString));
            }

            // todo I think this puts a color on every element.
            XElement color = colorPropertyString != null ?
                new XElement(W.color, new XAttribute(W.val, colorPropertyString)) : null;

            XElement sz = null;
            XElement szCs = null;
            if (fontSizeTPoint != null)
            {
                sz = new XElement(W.sz, new XAttribute(W.val, (int)((double)fontSizeTPoint * 2)));
                szCs = new XElement(W.szCs, new XAttribute(W.val, (int)((double)fontSizeTPoint * 2)));
            }

            XElement strike = null;
            if (textDecorationString == "line-through" || sAncestor)
                strike = new XElement(W.strike);

            XElement bold = null;
            XElement boldCs = null;
            if (bAncestor || strongAncestor || fontWeightString == "bold" || fontWeightString == "bolder" || fontWeightString == "600" || fontWeightString == "700" || fontWeightString == "800" || fontWeightString == "900")
            {
                bold = new XElement(W.b);
                boldCs = new XElement(W.bCs);
            }

            XElement italic = null;
            XElement italicCs = null;
            if (iAncestor || emAncestor || fontStyleString == "italic")
            {
                italic = new XElement(W.i);
                italicCs = new XElement(W.iCs);
            }

            XElement underline = null;
            if (uAncestor || textDecorationString == "underline")
                underline = new XElement(W.u, new XAttribute(W.val, "single"));

            XElement rStyle = null;
            if (element.Name == XhtmlNoNamespace.a)
                rStyle = new XElement(W.rStyle,
                    new XAttribute(W.val, "Hyperlink"));

            XElement spacing = null;
            if (letterSpacingProperty.IsNotNormal)
                spacing = new XElement(W.spacing,
                    new XAttribute(W.val, (long)(Twip)letterSpacingProperty));

            XElement rtl = null;
            if (dirAttributeString == "rtl" || directionString == "rtl")
                rtl = new XElement(W.rtl);

            XElement runProps = new XElement(W.rPr,
                rStyle,
                rFonts,
                bold,
                boldCs,
                italic,
                italicCs,
                strike,
                color,
                spacing,
                sz,
                szCs,
                underline,
                shd,
                subSuper,
                rtl);

            if (runProps.Elements().Any())
                return runProps;

            return null;
        }

        // todo can make this faster
        // todo this is not right - needs to be rationalized for all characters in an entire paragraph.
        // if there is text like <p>abc <em> def </em> ghi</p> then there needs to be just one space between abc and def, and between
        // def and ghi.
        private static string GetDisplayText(XText node, bool preserveWhiteSpace)
        {
            string textTransform = node.Parent.GetProp("text-transform").ToString();
            bool isFirst = node.Parent.Name == XhtmlNoNamespace.p && node == node.Parent.FirstNode;
            bool isLast = node.Parent.Name == XhtmlNoNamespace.p && node == node.Parent.LastNode;

            IEnumerable<IGrouping<bool, char>> groupedCharacters = null;
            if (preserveWhiteSpace)
                groupedCharacters = node.Value.GroupAdjacent(c => c == '\r' || c == '\n');
            else
                groupedCharacters = node.Value.GroupAdjacent(c => c == ' ' || c == '\r' || c == '\n');

            string newString = groupedCharacters.Select(g =>
            {
                if (g.Key == true)
                    return " ";
                string x = g.Select(c => c.ToString()).StringConcatenate();
                return x;
            })
                .StringConcatenate();
            if (!preserveWhiteSpace)
            {
                if (isFirst)
                    newString = newString.TrimStart();
                if (isLast)
                    newString = newString.TrimEnd();
            }
            if (textTransform == "uppercase")
                newString = newString.ToUpper();
            else if (textTransform == "lowercase")
                newString = newString.ToLower();
            else if (textTransform == "capitalize")
                newString = newString.Substring(0, 1).ToUpper() + newString.Substring(1).ToLower();
            return newString;
        }

        private static XElement GetNumberingProperties(XElement paragraph, HtmlToWmlConverterSettings settings)
        {
            // Numbering properties ******************************************************
            NumberedItemAnnotation numberedItemAnnotation = null;
            XElement listElement = paragraph.Ancestors().FirstOrDefault(a => a.Name == XhtmlNoNamespace.ol || a.Name == XhtmlNoNamespace.ul);
            if (listElement != null)
            {
                numberedItemAnnotation = listElement.Annotation<NumberedItemAnnotation>();
            }
            XElement numPr = null;
            if (paragraph.Name == XhtmlNoNamespace.li)
                numPr = new XElement(W.numPr,
                    new XElement(W.ilvl, new XAttribute(W.val, numberedItemAnnotation.ilvl)),
                    new XElement(W.numId, new XAttribute(W.val, numberedItemAnnotation.numId)));
            return numPr;
        }

        private static XElement GetJustification(XElement blockLevelElement, HtmlToWmlConverterSettings settings)
        {
            // Justify ******************************************************
            CssExpression textAlignProperty = blockLevelElement.GetProp("text-align");
            string textAlign;
            if (blockLevelElement.Name == XhtmlNoNamespace.caption || blockLevelElement.Name == XhtmlNoNamespace.th)
                textAlign = "center";
            else
                textAlign = "left";
            if (textAlignProperty != null)
                textAlign = textAlignProperty.ToString();
            string jc = null;
            if (textAlign == "center")
                jc = "center";
            else
            {
                if (textAlign == "right")
                    jc = "right";
                else
                {
                    if (textAlign == "justify")
                        jc = "both";
                }
            }
            string direction = GetDirection(blockLevelElement);
            if (direction == "rtl")
            {
                if (jc == "left")
                    jc = "right";
                else if (jc == "right")
                    jc = "left";
            }
            XElement jcElement = null;
            if (jc != null)
                jcElement = new XElement(W.jc, new XAttribute(W.val, jc));
            return jcElement;
        }

        private class HeadingInfo
        {
            public XName Name;
            public string StyleName;
        };

        private static HeadingInfo[] HeadingTagMap = new[]
            {
                new HeadingInfo { Name = XhtmlNoNamespace.h1, StyleName = "Heading1" },
                new HeadingInfo { Name = XhtmlNoNamespace.h2, StyleName = "Heading2" },
                new HeadingInfo { Name = XhtmlNoNamespace.h3, StyleName = "Heading3" },
                new HeadingInfo { Name = XhtmlNoNamespace.h4, StyleName = "Heading4" },
                new HeadingInfo { Name = XhtmlNoNamespace.h5, StyleName = "Heading5" },
                new HeadingInfo { Name = XhtmlNoNamespace.h6, StyleName = "Heading6" },
                new HeadingInfo { Name = XhtmlNoNamespace.h7, StyleName = "Heading7" },
                new HeadingInfo { Name = XhtmlNoNamespace.h8, StyleName = "Heading8" },
            };

        private static string GetDirection(XElement element)
        {
            string retValue = "ltr";
            string dirString = (string)element.Attribute(XhtmlNoNamespace.dir);
            if (dirString != null && dirString.ToLower() == "rtl")
                retValue = "rtl";
            CssExpression directionProp = element.GetProp("direction");
            if (directionProp != null)
            {
                string directionValue = directionProp.ToString();
                if (directionValue.ToLower() == "rtl")
                    retValue = "rtl";
            }
            return retValue;
        }

        private static XElement GetTableProperties(XElement element)
        {

            XElement bidiVisual = null;
            string direction = GetDirection(element);
            if (direction == "rtl")
                bidiVisual = new XElement(W.bidiVisual);

            XElement tblPr = new XElement(W.tblPr,
                bidiVisual,
                GetTableWidth(element),
                GetTableCellSpacing(element),
                GetBlockContentBorders(element, W.tblBorders, false),
                GetTableShading(element),
                GetTableCellMargins(element),
                GetTableLook(element));
            return tblPr;
        }

        private static XElement GetTableShading(XElement element)
        {
            // todo this is not done.
            // needs to work for W.tbl and W.tc
            //XElement shd = new XElement(W.shd,
            //    new XAttribute(W.val, "clear"),
            //    new XAttribute(W.color, "auto"),
            //    new XAttribute(W.fill, "ffffff"));
            //return shd;
            return null;
        }

        private static XElement GetTableWidth(XElement element)
        {
            CssExpression width = element.GetProp("width");
            if (width.IsAuto)
            {
                return new XElement(W.tblW,
                    new XAttribute(W._w, "0"),
                    new XAttribute(W.type, "auto"));
            }
            XElement widthElement = new XElement(W.tblW,
                new XAttribute(W._w, (long)(Twip)width),
                new XAttribute(W.type, "dxa"));
            return widthElement;
        }

        private static XElement GetCellWidth(XElement element)
        {
            CssExpression width = element.GetProp("width");
            if (width.IsAuto)
            {
                return new XElement(W.tcW,
                    new XAttribute(W._w, "0"),
                    new XAttribute(W.type, "auto"));
            }
            XElement widthElement = new XElement(W.tcW,
                new XAttribute(W._w, (long)(Twip)width),
                new XAttribute(W.type, "dxa"));
            return widthElement;
        }

        private static XElement GetBlockContentBorders(XElement element, XName borderXName, bool forParagraph)
        {
            if ((element.Name == XhtmlNoNamespace.td || element.Name == XhtmlNoNamespace.th || element.Name == XhtmlNoNamespace.caption) && forParagraph)
                return null;
            XElement borders = new XElement(borderXName,
                new XElement(W.top, GetBorderAttributes(element, "top")),
                new XElement(W.left, GetBorderAttributes(element, "left")),
                new XElement(W.bottom, GetBorderAttributes(element, "bottom")),
                new XElement(W.right, GetBorderAttributes(element, "right")));
            if (borders.Elements().Attributes(W.val).Where(v => (string)v == "none").Count() == 4)
                return null;
            return borders;
        }

        private static Dictionary<string, string> BorderStyleMap = new Dictionary<string, string>()
        {
            { "none", "none" },
            { "hidden", "none" },
            { "dotted", "dotted" },
            { "dashed", "dashed" },
            { "solid", "single" },
            { "double", "double" },
            { "groove", "inset" },
            { "ridge", "outset" },
            { "inset", "inset" },
            { "outset", "outset" },
        };

        private static List<XAttribute> GetBorderAttributes(XElement element, string whichBorder)
        {
            //if (whichBorder == "right")
            //    Console.WriteLine(1);
            CssExpression styleProp = element.GetProp(string.Format("border-{0}-style", whichBorder));
            CssExpression colorProp = element.GetProp(string.Format("border-{0}-color", whichBorder));
            CssExpression paddingProp = element.GetProp(string.Format("padding-{0}", whichBorder));
            CssExpression marginProp = element.GetProp(string.Format("margin-{0}", whichBorder));

            // The space attribute is equivalent to the margin properties of CSS
            // the ind element of the parent is more or less equivalent to the padding properties of CSS, except that ind takes space
            // AWAY from the space attribute, therefore ind needs to be increased by the amount of padding.

            // if there is no border, and yet there is padding, then need to create a thin border so that word will display the background
            // color of the paragraph properly (including padding).

            XAttribute val = null;
            XAttribute sz = null;
            XAttribute space = null;
            XAttribute color = null;

            if (styleProp != null)
            {
                if (BorderStyleMap.ContainsKey(styleProp.ToString()))
                    val = new XAttribute(W.val, BorderStyleMap[styleProp.ToString()]);
                else
                    val = new XAttribute(W.val, "none");
            }

            double borderSizeInTwips = GetBorderSize(element, whichBorder);

            double borderSizeInOneEighthPoint = borderSizeInTwips / 20 * 8;
            sz = new XAttribute(W.sz, (int)borderSizeInOneEighthPoint);

            if (element.Name == XhtmlNoNamespace.td || element.Name == XhtmlNoNamespace.th)
            {
                space = new XAttribute(W.space, "0");
#if false
                // 2012-05-14 todo alternative algorithm for margin for cells
                if (marginProp != null)
                {
                    // space is specified in points, not twips
                    TPoint points = 0;
                    if (marginProp.IsNotAuto)
                        points = (TPoint)marginProp;
                    space = new XAttribute(W.space, Math.Min(31, (double)points));
                }
#endif
            }
            else
            {
                space = new XAttribute(W.space, "0");
                if (paddingProp != null)
                {
                    // space is specified in points, not twips
                    TPoint points = (TPoint)paddingProp;
                    space = new XAttribute(W.space, (int)(Math.Min(31, (double)points)));
                }
            }

            if (colorProp != null)
                color = new XAttribute(W.color, colorProp.ToString());
            // no default yet

            if ((string)val == "none" && (double)space != 0d)
            {
                val.Value = "single";
                sz.Value = "0";
                //color.Value = "FF0000";
            }

            // sz is in 1/8 of a point
            // space is in 1/20 of a point

            List<XAttribute> attList = new List<XAttribute>()
            {
                val,
                sz,
                space,
                color,
            };
            return attList;
        }

        private static Twip GetBorderSize(XElement element, string whichBorder)
        {
            CssExpression widthProp = element.GetProp(string.Format("border-{0}-width", whichBorder));

            if (widthProp != null && widthProp.Terms.Count() == 1)
            {
                CssTerm term = widthProp.Terms.First();
                Twip twips = (Twip)widthProp;
                return twips;
            }
            return 12;
        }

        private static XElement GetTableLook(XElement element)
        {
            XElement tblLook = XElement.Parse(
                //@"<w:tblLook w:val='0600'
                //  w:firstRow='0'
                //  w:lastRow='0'
                //  w:firstColumn='0'
                //  w:lastColumn='0'
                //  w:noHBand='1'
                //  w:noVBand='1'
                //  xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'/>"

@"<w:tblLook w:val='0600' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'/>"

);
            tblLook.Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
            return tblLook;
        }

        private static XElement GetTableGrid(XElement element, HtmlToWmlConverterSettings settings)
        {
            Twip? pageWidthInTwips = (int?)settings.SectPr.Elements(W.pgSz).Attributes(W._w).FirstOrDefault();
            Twip? marginLeft = (int?)settings.SectPr.Elements(W.pgMar).Attributes(W.left).FirstOrDefault();
            Twip? marginRight = (int?)settings.SectPr.Elements(W.pgMar).Attributes(W.right).FirstOrDefault();
            Twip printable = (long)pageWidthInTwips - (long)marginLeft - (long)marginRight;
            XElement[][] tableArray = GetTableArray(element);
            int numberColumns = tableArray[0].Length;
            CssExpression[] columnWidths = new CssExpression[numberColumns];
            for (int c = 0; c < numberColumns; c++)
            {
                CssExpression columnWidth = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "auto" } } };
                for (int r = 0; r < tableArray.Length; ++r)
                {
                    if (tableArray[r][c] != null)
                    {
                        XElement cell = tableArray[r][c];
                        CssExpression width = cell.GetProp("width");
                        XAttribute colSpan = cell.Attribute(XhtmlNoNamespace.colspan);
                        if (colSpan == null && columnWidth.ToString() == "auto" && width.ToString() != "auto")
                        {
                            columnWidth = width;
                            break;
                        }
                    }
                }
                columnWidths[c] = columnWidth;
            }

            XElement tblGrid = new XElement(W.tblGrid,
                columnWidths.Select(cw => new XElement(W.gridCol,
                    new XAttribute(W._w, (long)GetTwipWidth(cw, (int)printable)))));
            return tblGrid;
        }

        private static Twip GetTwipWidth(CssExpression columnWidth, int printable)
        {
            Twip defaultTwipWidth = 1440;
            if (columnWidth.Terms.Count() == 1)
            {
                CssTerm term = columnWidth.Terms.First();
                if (term.Unit == CssUnit.PT)
                {
                    Double ptValue;
                    if (Double.TryParse(term.Value, out ptValue))
                    {
                        Twip twips = (long)(ptValue * 20);
                        return twips;
                    }
                    return defaultTwipWidth;
                }
            }
            return defaultTwipWidth;
        }

        private static XElement[][] GetTableArray(XElement table)
        {
            List<XElement> rowList = table.DescendantsTrimmed(XhtmlNoNamespace.table).Where(e => e.Name == XhtmlNoNamespace.tr).ToList();
            int numberColumns = rowList.Select(r => r.Elements().Where(e => e.Name == XhtmlNoNamespace.td || e.Name == XhtmlNoNamespace.th).Count()).Max();
            XElement[][] tableArray = new XElement[rowList.Count()][];
            int rowNumber = 0;
            foreach (var row in rowList)
            {
                tableArray[rowNumber] = new XElement[numberColumns];
                int columnNumber = 0;
                foreach (var cell in row.Elements(XhtmlNoNamespace.td))
                {
                    tableArray[rowNumber][columnNumber] = cell;
                    columnNumber++;
                }
                rowNumber++;
            }
            return tableArray;
        }

        private static XElement GetCellPropertiesForCaption(XElement element)
        {
            XElement gridSpan = new XElement(W.gridSpan,
                    new XAttribute(W.val, 3));

            XElement tcBorders = GetBlockContentBorders(element, W.tcBorders, false);
            if (tcBorders == null)
                tcBorders = new XElement(W.tcBorders,
                    new XElement(W.top, new XAttribute(W.val, "nil")),
                    new XElement(W.left, new XAttribute(W.val, "nil")),
                    new XElement(W.bottom, new XAttribute(W.val, "nil")),
                    new XElement(W.right, new XAttribute(W.val, "nil")));

            XElement shd = GetCellShading(element);

            //XElement hideMark = new XElement(W.hideMark);
            XElement hideMark = null;

            XElement tcMar = GetCellMargins(element);

            XElement vAlign = new XElement(W.vAlign, new XAttribute(W.val, "center"));

            return new XElement(W.tcPr,
                gridSpan,
                tcBorders,
                shd,
                tcMar,
                vAlign,
                hideMark);
        }

        private static XElement GetCellProperties(XElement element)
        {
            int? colspan = (int?)element.Attribute(XhtmlNoNamespace.colspan);
            XElement gridSpan = null;
            if (colspan != null)
                gridSpan = new XElement(W.gridSpan,
                    new XAttribute(W.val, colspan));

            XElement tblW = GetCellWidth(element);

            XElement tcBorders = GetBlockContentBorders(element, W.tcBorders, false);

            XElement shd = GetCellShading(element);

            //XElement hideMark = new XElement(W.hideMark);
            XElement hideMark = null;

            XElement tcMar = GetCellMargins(element);

            XElement vAlign = new XElement(W.vAlign, new XAttribute(W.val, "center"));

            XElement vMerge = null;
            if (element.Attribute("HtmlToWmlVMergeNoRestart") != null)
                vMerge = new XElement(W.vMerge);
            else
                if (element.Attribute("HtmlToWmlVMergeRestart") != null)
                    vMerge = new XElement(W.vMerge,
                        new XAttribute(W.val, "restart"));

            string vAlignValue = (string)element.Attribute(XhtmlNoNamespace.valign);
            CssExpression verticalAlignmentProp = element.GetProp("vertical-align");
            if (verticalAlignmentProp != null && verticalAlignmentProp.ToString() != "inherit")
                vAlignValue = verticalAlignmentProp.ToString();
            if (vAlignValue != null)
            {
                if (vAlignValue == "middle" || (vAlignValue != "top" && vAlignValue != "bottom"))
                    vAlignValue = "center";
                vAlign = new XElement(W.vAlign, new XAttribute(W.val, vAlignValue));
            }

            return new XElement(W.tcPr,
                tblW,
                gridSpan,
                vMerge,
                tcBorders,
                shd,
                tcMar,
                vAlign,
                hideMark);
        }

        private static XElement GetCellHeaderProperties(XElement element)
        {
            //int? colspan = (int?)element.Attribute(Xhtml.colspan);
            //XElement gridSpan = null;
            //if (colspan != null)
            //    gridSpan = new XElement(W.gridSpan,
            //        new XAttribute(W.val, colspan));

            XElement tblW = GetCellWidth(element);

            XElement tcBorders = GetBlockContentBorders(element, W.tcBorders, false);

            XElement shd = GetCellShading(element);

            //XElement hideMark = new XElement(W.hideMark);
            XElement hideMark = null;

            XElement tcMar = GetCellMargins(element);

            XElement vAlign = new XElement(W.vAlign, new XAttribute(W.val, "center"));

            return new XElement(W.tcPr,
                tblW,
                tcBorders,
                shd,
                tcMar,
                vAlign,
                hideMark);
        }

        private static XElement GetCellShading(XElement element)
        {
            CssExpression backgroundColorProp = element.GetProp("background-color");
            if (backgroundColorProp != null && (string)backgroundColorProp != "transparent")
            {
                XElement shd = new XElement(W.shd,
                    new XAttribute(W.val, "clear"),
                    new XAttribute(W.color, "auto"),
                    new XAttribute(W.fill, backgroundColorProp));
                return shd;
            }
            return null;
        }

        private static XElement GetCellMargins(XElement element)
        {
            CssExpression topProp = element.GetProp("padding-top");
            CssExpression leftProp = element.GetProp("padding-left");
            CssExpression bottomProp = element.GetProp("padding-bottom");
            CssExpression rightProp = element.GetProp("padding-right");
            if ((long)topProp == 0 &&
                (long)leftProp == 0 &&
                (long)bottomProp == 0 &&
                (long)rightProp == 0)
                return null;
            XElement top = null;
            if (topProp != null)
                top = new XElement(W.top,
                    new XAttribute(W._w, (long)(Twip)topProp),
                    new XAttribute(W.type, "dxa"));
            XElement left = null;
            if (leftProp != null)
                left = new XElement(W.left,
                    new XAttribute(W._w, (long)(Twip)leftProp),
                    new XAttribute(W.type, "dxa"));
            XElement bottom = null;
            if (bottomProp != null)
                bottom = new XElement(W.bottom,
                    new XAttribute(W._w, (long)(Twip)bottomProp),
                    new XAttribute(W.type, "dxa"));
            XElement right = null;
            if (rightProp != null)
                right = new XElement(W.right,
                    new XAttribute(W._w, (long)(Twip)rightProp),
                    new XAttribute(W.type, "dxa"));
            XElement tcMar = new XElement(W.tcMar,
                top, left, bottom, right);
            if (tcMar.Elements().Any())
                return tcMar;
            return null;
        }

#if false
            <w:tcMar>
              <w:top w:w="720"
                     w:type="dxa" />
              <w:left w:w="720"
                      w:type="dxa" />
              <w:bottom w:w="720"
                        w:type="dxa" />
              <w:right w:w="720"
                       w:type="dxa" />
            </w:tcMar>
#endif

        private static XElement GetTableCellSpacing(XElement element)
        {
            XElement table = element.AncestorsAndSelf(XhtmlNoNamespace.table).FirstOrDefault();
            XElement tblCellSpacing = null;
            if (table != null)
            {
                CssExpression borderCollapse = table.GetProp("border-collapse");
                if (borderCollapse == null || (string)borderCollapse != "collapse")
                {
                    // todo very incomplete
                    CssExpression borderSpacing = table.GetProp("border-spacing");
                    CssExpression marginTopProperty = element.GetProp("margin-top");
                    if (marginTopProperty == null)
                        marginTopProperty = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "0", Type = CssTermType.Number, Unit = CssUnit.PT } } };
                    CssExpression marginBottomProperty = element.GetProp("margin-bottom");
                    if (marginBottomProperty == null)
                        marginBottomProperty = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "0", Type = CssTermType.Number, Unit = CssUnit.PT } } };
                    Twip twips1 = (Twip)marginTopProperty;
                    Twip twips2 = (Twip)marginBottomProperty;
                    Twip minTwips = 15;
                    if (borderSpacing != null)
                        minTwips = (Twip)borderSpacing;
                    long twipToUse = Math.Max((long)twips1, (long)twips2);
                    twipToUse = Math.Max(twipToUse, (long)minTwips);
                    // have to divide twipToUse by 2 because border-spacing specifies the space between the border of once cell and its adjacent.
                    // tblCellSpacing specifies the distance between the border and the half way point between two cells.
                    long twipToUseOverTwo = (long)twipToUse / 2;
                    tblCellSpacing = new XElement(W.tblCellSpacing, new XAttribute(W._w, twipToUseOverTwo),
                        new XAttribute(W.type, "dxa"));
                }

            }
            return tblCellSpacing;
        }

        private static XElement GetTableCellMargins(XElement element)
        {
            // todo very incomplete
            XElement tblCellMar = XElement.Parse(
@"<w:tblCellMar xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:top w:w='15'
          w:type='dxa'/>
  <w:left w:w='15'
          w:type='dxa'/>
  <w:bottom w:w='15'
            w:type='dxa'/>
  <w:right w:w='15'
            w:type='dxa'/>
</w:tblCellMar>");
            tblCellMar.Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
            return tblCellMar;
        }

        private static XElement GetTableRowProperties(XElement element)
        {
            XElement trPr = null;
            XElement table = element.AncestorsAndSelf(XhtmlNoNamespace.table).FirstOrDefault();
            if (table != null)
            {
                CssExpression heightProperty = element.GetProp("height");
                //long? maxCellHeight = element.Elements(Xhtml.td).Aggregate((long?)null,
                //    (XElement td, long? last) =>
                //    {
                //        Expression heightProp2 = td.GetProp("height");
                //        if (heightProp2 == null)
                //            return last;
                //        if (last == null)
                //            return (long)(Twip)heightProp2;
                //        return last + (long?)(long)(Twip)heightProp2;
                //    });
                var cellHeights = element
                    .Elements(XhtmlNoNamespace.td)
                    .Select(td => td.GetProp("height"))
                    .Concat(new[] { heightProperty })
                    .Where(d => d != null)
                    .Select(e => (long)(Twip)e)
                    .ToList();
                XElement trHeight = null;
                if (cellHeights.Any())
                {
                    long max = cellHeights.Max();
                    trHeight = new XElement(W.trHeight,
                        new XAttribute(W.val, max));
                }

                CssExpression borderCollapseProperty = table.GetProp("border-collapse");
                XElement borderCollapse = null;
                if (borderCollapseProperty != null && (string)borderCollapseProperty != "collapse")
                    borderCollapse = GetTableCellSpacing(element);

                trPr = new XElement(W.trPr,
                    GetTableCellSpacing(element),
                    trHeight);
                if (trPr.Elements().Any())
                    return trPr;
            }
            return trPr;
        }

        private static XAttribute GetXmlSpaceAttribute(string value)
        {
            if (value.StartsWith(" ") || value.EndsWith(" "))
                return new XAttribute(XNamespace.Xml + "space", "preserve");
            return null;
        }


        private static Dictionary<string, string> InstalledFonts = new Dictionary<string, string>
            {
                {"serif", "Times New Roman"},
                {"sans-serif", "Arial"},
                {"cursive", "Kunstler Script"},
                {"fantasy", "Curlz MT"},
                {"monospace", "Courier New"},

                {"agency fb", "Agency FB"},
                {"agencyfb", "Agency FB"},
                {"aharoni", "Aharoni"},
                {"algerian", "Algerian"},
                {"andalus", "Andalus"},
                {"angsana new", "Angsana New"},
                {"angsananew", "Angsana New"},
                {"angsanaupc", "AngsanaUPC"},
                {"aparajita", "Aparajita"},
                {"arabic typesetting", "Arabic Typesetting"},
                {"arabictypesetting", "Arabic Typesetting"},
                {"arial", "Arial"},
                {"arial black", "Arial Black"},
                {"arial narrow", "Arial Narrow"},
                {"arial rounded mt bold", "Arial Rounded MT Bold"},
                {"arial unicode ms", "Arial Unicode MS"},
                {"arialblack", "Arial Black"},
                {"arialnarrow", "Arial Narrow"},
                {"arialroundedmtbold", "Arial Rounded MT Bold"},
                {"arialunicodems", "Arial Unicode MS"},
                {"baskerville old face", "Baskerville Old Face"},
                {"baskervilleoldface", "Baskerville Old Face"},
                {"batang", "Batang"},
                {"batangche", "BatangChe"},
                {"bauhaus 93", "Bauhaus 93"},
                {"bauhaus93", "Bauhaus 93"},
                {"bell mt", "Bell MT"},
                {"bellmt", "Bell MT"},
                {"berlin sans fb", "Berlin Sans FB"},
                {"berlin sans fb demi", "Berlin Sans FB Demi"},
                {"berlinsansfb", "Berlin Sans FB"},
                {"berlinsansfbdemi", "Berlin Sans FB Demi"},
                {"bernard mt condensed", "Bernard MT Condensed"},
                {"bernardmtcondensed", "Bernard MT Condensed"},
                {"blackadder itc", "Blackadder ITC"},
                {"blackadderitc", "Blackadder ITC"},
                {"bodoni mt", "Bodoni MT"},
                {"bodoni mt black", "Bodoni MT Black"},
                {"bodoni mt condensed", "Bodoni MT Condensed"},
                {"bodoni mt poster compressed", "Bodoni MT Poster Compressed"},
                {"bodonimt", "Bodoni MT"},
                {"bodonimtblack", "Bodoni MT Black"},
                {"bodonimtcondensed", "Bodoni MT Condensed"},
                {"bodonimtpostercompressed", "Bodoni MT Poster Compressed"},
                {"book antiqua", "Book Antiqua"},
                {"bookantiqua", "Book Antiqua"},
                {"bookman old style", "Bookman Old Style"},
                {"bookmanoldstyle", "Bookman Old Style"},
                {"bookshelf symbol 7", "Bookshelf Symbol 7"},
                {"bookshelfsymbol7", "Bookshelf Symbol 7"},
                {"bradley hand itc", "Bradley Hand ITC"},
                {"bradleyhanditc", "Bradley Hand ITC"},
                {"britannic bold", "Britannic Bold"},
                {"britannicbold", "Britannic Bold"},
                {"broadway", "Broadway"},
                {"browallia new", "Browallia New"},
                {"browallianew", "Browallia New"},
                {"browalliaupc", "BrowalliaUPC"},
                {"brush script mt", "Brush Script MT"},
                {"brushscriptmt", "Brush Script MT"},
                {"calibri", "Calibri"},
                {"californian fb", "Californian FB"},
                {"californianfb", "Californian FB"},
                {"calisto mt", "Calisto MT"},
                {"calistomt", "Calisto MT"},
                {"cambria", "Cambria"},
                {"cambria math", "Cambria Math"},
                {"cambriamath", "Cambria Math"},
                {"candara", "Candara"},
                {"castellar", "Castellar"},
                {"centaur", "Centaur"},
                {"century", "Century"},
                {"century gothic", "Century Gothic"},
                {"century schoolbook", "Century Schoolbook"},
                {"centurygothic", "Century Gothic"},
                {"centuryschoolbook", "Century Schoolbook"},
                {"chiller", "Chiller"},
                {"colonna mt", "Colonna MT"},
                {"colonnamt", "Colonna MT"},
                {"comic sans ms", "Comic Sans MS"},
                {"comicsansms", "Comic Sans MS"},
                {"consolas", "Consolas"},
                {"constantia", "Constantia"},
                {"cooper black", "Cooper Black"},
                {"cooperblack", "Cooper Black"},
                {"copperplate gothic bold", "Copperplate Gothic Bold"},
                {"copperplate gothic light", "Copperplate Gothic Light"},
                {"copperplategothicbold", "Copperplate Gothic Bold"},
                {"copperplategothiclight", "Copperplate Gothic Light"},
                {"corbel", "Corbel"},
                {"cordia new", "Cordia New"},
                {"cordianew", "Cordia New"},
                {"cordiaupc", "CordiaUPC"},
                {"courier new", "Courier New"},
                {"couriernew", "Courier New"},
                {"curlz mt", "Curlz MT"},
                {"curlzmt", "Curlz MT"},
                {"daunpenh", "DaunPenh"},
                {"david", "David"},
                {"dfkai-sb", "DFKai-SB"},
                {"dilleniaupc", "DilleniaUPC"},
                {"dokchampa", "DokChampa"},
                {"dotum", "Dotum"},
                {"dotumche", "DotumChe"},
                {"ebrima", "Ebrima"},
                {"edwardian script itc", "Edwardian Script ITC"},
                {"edwardianscriptitc", "Edwardian Script ITC"},
                {"elephant", "Elephant"},
                {"engravers mt", "Engravers MT"},
                {"engraversmt", "Engravers MT"},
                {"eras bold itc", "Eras Bold ITC"},
                {"eras demi itc", "Eras Demi ITC"},
                {"eras light itc", "Eras Light ITC"},
                {"eras medium itc", "Eras Medium ITC"},
                {"erasbolditc", "Eras Bold ITC"},
                {"erasdemiitc", "Eras Demi ITC"},
                {"eraslightitc", "Eras Light ITC"},
                {"erasmediumitc", "Eras Medium ITC"},
                {"estrangelo edessa", "Estrangelo Edessa"},
                {"estrangeloedessa", "Estrangelo Edessa"},
                {"eucrosiaupc", "EucrosiaUPC"},
                {"euphemia", "Euphemia"},
                {"fangsong", "FangSong"},
                {"felix titling", "Felix Titling"},
                {"felixtitling", "Felix Titling"},
                {"footlight mt light", "Footlight MT Light"},
                {"footlightmtlight", "Footlight MT Light"},
                {"forte", "Forte"},
                {"franklin gothic book", "Franklin Gothic Book"},
                {"franklin gothic demi", "Franklin Gothic Demi"},
                {"franklin gothic demi cond", "Franklin Gothic Demi Cond"},
                {"franklin gothic heavy", "Franklin Gothic Heavy"},
                {"franklin gothic medium", "Franklin Gothic Medium"},
                {"franklin gothic medium cond", "Franklin Gothic Medium Cond"},
                {"franklingothicbook", "Franklin Gothic Book"},
                {"franklingothicdemi", "Franklin Gothic Demi"},
                {"franklingothicdemicond", "Franklin Gothic Demi Cond"},
                {"franklingothicheavy", "Franklin Gothic Heavy"},
                {"franklingothicmedium", "Franklin Gothic Medium"},
                {"franklingothicmediumcond", "Franklin Gothic Medium Cond"},
                {"frankruehl", "FrankRuehl"},
                {"freesiaupc", "FreesiaUPC"},
                {"freestyle script", "Freestyle Script"},
                {"freestylescript", "Freestyle Script"},
                {"french script mt", "French Script MT"},
                {"frenchscriptmt", "French Script MT"},
                {"gabriola", "Gabriola"},
                {"garamond", "Garamond"},
                {"gautami", "Gautami"},
                {"georgia", "Georgia"},
                {"gigi", "Gigi"},
                {"gill sans mt", "Gill Sans MT"},
                {"gill sans mt condensed", "Gill Sans MT Condensed"},
                {"gill sans mt ext condensed bold", "Gill Sans MT Ext Condensed Bold"},
                {"gill sans ultra bold", "Gill Sans Ultra Bold"},
                {"gill sans ultra bold condensed", "Gill Sans Ultra Bold Condensed"},
                {"gillsansmt", "Gill Sans MT"},
                {"gillsansmtcondensed", "Gill Sans MT Condensed"},
                {"gillsansmtextcondensedbold", "Gill Sans MT Ext Condensed Bold"},
                {"gillsansultrabold", "Gill Sans Ultra Bold"},
                {"gillsansultraboldcondensed", "Gill Sans Ultra Bold Condensed"},
                {"gisha", "Gisha"},
                {"gloucester mt extra condensed", "Gloucester MT Extra Condensed"},
                {"gloucestermtextracondensed", "Gloucester MT Extra Condensed"},
                {"goudy old style", "Goudy Old Style"},
                {"goudy stout", "Goudy Stout"},
                {"goudyoldstyle", "Goudy Old Style"},
                {"goudystout", "Goudy Stout"},
                {"gulim", "Gulim"},
                {"gulimche", "GulimChe"},
                {"gungsuh", "Gungsuh"},
                {"gungsuhche", "GungsuhChe"},
                {"haettenschweiler", "Haettenschweiler"},
                {"harlow solid italic", "Harlow Solid Italic"},
                {"harlowsoliditalic", "Harlow Solid Italic"},
                {"harrington", "Harrington"},
                {"high tower text", "High Tower Text"},
                {"hightowertext", "High Tower Text"},
                {"impact", "Impact"},
                {"imprint mt shadow", "Imprint MT Shadow"},
                {"imprintmtshadow", "Imprint MT Shadow"},
                {"informal roman", "Informal Roman"},
                {"informalroman", "Informal Roman"},
                {"irisupc", "IrisUPC"},
                {"iskoola pota", "Iskoola Pota"},
                {"iskoolapota", "Iskoola Pota"},
                {"jasmineupc", "JasmineUPC"},
                {"jokerman", "Jokerman"},
                {"juice itc", "Juice ITC"},
                {"juiceitc", "Juice ITC"},
                {"kaiti", "KaiTi"},
                {"kalinga", "Kalinga"},
                {"kartika", "Kartika"},
                {"khmer ui", "Khmer UI"},
                {"khmerui", "Khmer UI"},
                {"kodchiangupc", "KodchiangUPC"},
                {"kokila", "Kokila"},
                {"kristen itc", "Kristen ITC"},
                {"kristenitc", "Kristen ITC"},
                {"kunstler script", "Kunstler Script"},
                {"kunstlerscript", "Kunstler Script"},
                {"lao ui", "Lao UI"},
                {"laoui", "Lao UI"},
                {"latha", "Latha"},
                {"leelawadee", "Leelawadee"},
                {"levenim mt", "Levenim MT"},
                {"levenimmt", "Levenim MT"},
                {"lilyupc", "LilyUPC"},
                {"lucida bright", "Lucida Bright"},
                {"lucida calligraphy", "Lucida Calligraphy"},
                {"lucida console", "Lucida Console"},
                {"lucida fax", "Lucida Fax"},
                {"lucida handwriting", "Lucida Handwriting"},
                {"lucida sans", "Lucida Sans"},
                {"lucida sans typewriter", "Lucida Sans Typewriter"},
                {"lucida sans unicode", "Lucida Sans Unicode"},
                {"lucidabright", "Lucida Bright"},
                {"lucidacalligraphy", "Lucida Calligraphy"},
                {"lucidaconsole", "Lucida Console"},
                {"lucidafax", "Lucida Fax"},
                {"lucidahandwriting", "Lucida Handwriting"},
                {"lucidasans", "Lucida Sans"},
                {"lucidasanstypewriter", "Lucida Sans Typewriter"},
                {"lucidasansunicode", "Lucida Sans Unicode"},
                {"magneto", "Magneto"},
                {"maiandra gd", "Maiandra GD"},
                {"maiandragd", "Maiandra GD"},
                {"malgun gothic", "Malgun Gothic"},
                {"malgungothic", "Malgun Gothic"},
                {"mangal", "Mangal"},
                {"marlett", "Marlett"},
                {"matura mt script capitals", "Matura MT Script Capitals"},
                {"maturamtscriptcapitals", "Matura MT Script Capitals"},
                {"meiryo", "Meiryo"},
                {"meiryo ui", "Meiryo UI"},
                {"meiryoui", "Meiryo UI"},
                {"microsoft himalaya", "Microsoft Himalaya"},
                {"microsoft jhenghei", "Microsoft JhengHei"},
                {"microsoft new tai lue", "Microsoft New Tai Lue"},
                {"microsoft phagspa", "Microsoft PhagsPa"},
                {"microsoft sans serif", "Microsoft Sans Serif"},
                {"microsoft tai le", "Microsoft Tai Le"},
                {"microsoft uighur", "Microsoft Uighur"},
                {"microsoft yahei", "Microsoft YaHei"},
                {"microsoft yi baiti", "Microsoft Yi Baiti"},
                {"microsofthimalaya", "Microsoft Himalaya"},
                {"microsoftjhenghei", "Microsoft JhengHei"},
                {"microsoftnewtailue", "Microsoft New Tai Lue"},
                {"microsoftphagspa", "Microsoft PhagsPa"},
                {"microsoftsansserif", "Microsoft Sans Serif"},
                {"microsofttaile", "Microsoft Tai Le"},
                {"microsoftuighur", "Microsoft Uighur"},
                {"microsoftyahei", "Microsoft YaHei"},
                {"microsoftyibaiti", "Microsoft Yi Baiti"},
                {"mingliu", "MingLiU"},
                {"mingliu_hkscs", "MingLiU_HKSCS"},
                {"mingliu_hkscs-extb", "MingLiU_HKSCS-ExtB"},
                {"mingliu-extb", "MingLiU-ExtB"},
                {"miriam", "Miriam"},
                {"miriam fixed", "Miriam Fixed"},
                {"miriamfixed", "Miriam Fixed"},
                {"mistral", "Mistral"},
                {"modern no. 20", "Modern No. 20"},
                {"modernno.20", "Modern No. 20"},
                {"mongolian baiti", "Mongolian Baiti"},
                {"mongolianbaiti", "Mongolian Baiti"},
                {"monotype corsiva", "Monotype Corsiva"},
                {"monotypecorsiva", "Monotype Corsiva"},
                {"moolboran", "MoolBoran"},
                {"ms gothic", "MS Gothic"},
                {"ms mincho", "MS Mincho"},
                {"ms pgothic", "MS PGothic"},
                {"ms pmincho", "MS PMincho"},
                {"ms reference sans serif", "MS Reference Sans Serif"},
                {"ms reference specialty", "MS Reference Specialty"},
                {"ms ui gothic", "MS UI Gothic"},
                {"msgothic", "MS Gothic"},
                {"msmincho", "MS Mincho"},
                {"mspgothic", "MS PGothic"},
                {"mspmincho", "MS PMincho"},
                {"msreferencesansserif", "MS Reference Sans Serif"},
                {"msreferencespecialty", "MS Reference Specialty"},
                {"msuigothic", "MS UI Gothic"},
                {"mt extra", "MT Extra"},
                {"mtextra", "MT Extra"},
                {"mv boli", "MV Boli"},
                {"mvboli", "MV Boli"},
                {"narkisim", "Narkisim"},
                {"niagara engraved", "Niagara Engraved"},
                {"niagara solid", "Niagara Solid"},
                {"niagaraengraved", "Niagara Engraved"},
                {"niagarasolid", "Niagara Solid"},
                {"nsimsun", "NSimSun"},
                {"nyala", "Nyala"},
                {"ocr a extended", "OCR A Extended"},
                {"ocraextended", "OCR A Extended"},
                {"old english text mt", "Old English Text MT"},
                {"oldenglishtextmt", "Old English Text MT"},
                {"onyx", "Onyx"},
                {"palace script mt", "Palace Script MT"},
                {"palacescriptmt", "Palace Script MT"},
                {"palatino linotype", "Palatino Linotype"},
                {"palatinolinotype", "Palatino Linotype"},
                {"papyrus", "Papyrus"},
                {"parchment", "Parchment"},
                {"perpetua", "Perpetua"},
                {"perpetua titling mt", "Perpetua Titling MT"},
                {"perpetuatitlingmt", "Perpetua Titling MT"},
                {"plantagenet cherokee", "Plantagenet Cherokee"},
                {"plantagenetcherokee", "Plantagenet Cherokee"},
                {"playbill", "Playbill"},
                {"pmingliu", "PMingLiU"},
                {"pmingliu-extb", "PMingLiU-ExtB"},
                {"poor richard", "Poor Richard"},
                {"poorrichard", "Poor Richard"},
                {"pristina", "Pristina"},
                {"raavi", "Raavi"},
                {"rage italic", "Rage Italic"},
                {"rageitalic", "Rage Italic"},
                {"ravie", "Ravie"},
                {"rockwell", "Rockwell"},
                {"rockwell condensed", "Rockwell Condensed"},
                {"rockwell extra bold", "Rockwell Extra Bold"},
                {"rockwellcondensed", "Rockwell Condensed"},
                {"rockwellextrabold", "Rockwell Extra Bold"},
                {"rod", "Rod"},
                {"sakkal majalla", "Sakkal Majalla"},
                {"sakkalmajalla", "Sakkal Majalla"},
                {"script mt bold", "Script MT Bold"},
                {"scriptmtbold", "Script MT Bold"},
                {"segoe print", "Segoe Print"},
                {"segoe script", "Segoe Script"},
                {"segoe ui", "Segoe UI"},
                {"segoe ui light", "Segoe UI Light"},
                {"segoe ui semibold", "Segoe UI Semibold"},
                {"segoe ui symbol", "Segoe UI Symbol"},
                {"segoeprint", "Segoe Print"},
                {"segoescript", "Segoe Script"},
                {"segoeui", "Segoe UI"},
                {"segoeuilight", "Segoe UI Light"},
                {"segoeuisemibold", "Segoe UI Semibold"},
                {"segoeuisymbol", "Segoe UI Symbol"},
                {"shonar bangla", "Shonar Bangla"},
                {"shonarbangla", "Shonar Bangla"},
                {"showcard gothic", "Showcard Gothic"},
                {"showcardgothic", "Showcard Gothic"},
                {"shruti", "Shruti"},
                {"simhei", "SimHei"},
                {"simplified arabic", "Simplified Arabic"},
                {"simplified arabic fixed", "Simplified Arabic Fixed"},
                {"simplifiedarabic", "Simplified Arabic"},
                {"simplifiedarabicfixed", "Simplified Arabic Fixed"},
                {"simsun", "SimSun"},
                {"simsun-extb", "SimSun-ExtB"},
                {"snap itc", "Snap ITC"},
                {"snapitc", "Snap ITC"},
                {"stencil", "Stencil"},
                {"swgamekeys mt", "SWGamekeys MT"},
                {"swgamekeysmt", "SWGamekeys MT"},
                {"swmacro", "SWMacro"},
                {"sylfaen", "Sylfaen"},
                {"symbol", "Symbol"},
                {"tahoma", "Tahoma"},
                {"tempus sans itc", "Tempus Sans ITC"},
                {"tempussansitc", "Tempus Sans ITC"},
                {"times new roman", "Times New Roman"},
                {"timesnewroman", "Times New Roman"},
                {"traditional arabic", "Traditional Arabic"},
                {"traditionalarabic", "Traditional Arabic"},
                {"trebuchet ms", "Trebuchet MS"},
                {"trebuchetms", "Trebuchet MS"},
                {"tunga", "Tunga"},
                {"tw cen mt", "Tw Cen MT"},
                {"tw cen mt condensed", "Tw Cen MT Condensed"},
                {"tw cen mt condensed extra bold", "Tw Cen MT Condensed Extra Bold"},
                {"twcenmt", "Tw Cen MT"},
                {"twcenmtcondensed", "Tw Cen MT Condensed"},
                {"twcenmtcondensedextrabold", "Tw Cen MT Condensed Extra Bold"},
                {"utsaah", "Utsaah"},
                {"vani", "Vani"},
                {"verdana", "Verdana"},
                {"vijaya", "Vijaya"},
                {"viner hand itc", "Viner Hand ITC"},
                {"vinerhanditc", "Viner Hand ITC"},
                {"vivaldi", "Vivaldi"},
                {"vladimir script", "Vladimir Script"},
                {"vladimirscript", "Vladimir Script"},
                {"vrinda", "Vrinda"},
                {"webdings", "Webdings"},
                {"wide latin", "Wide Latin"},
                {"widelatin", "Wide Latin"},
                {"wingdings", "Wingdings"},
                {"wingdings 2", "Wingdings 2"},
                {"wingdings 3", "Wingdings 3"},
                {"wingdings2", "Wingdings 2"},
                {"wingdings3", "Wingdings 3"},
            };

        private static TPoint? GetUsedSizeFromFontSizeProperty(CssExpression fontSize)
        {
            if (fontSize == null)
                return null;
            if (fontSize.Terms.Count() == 1)
            {
                CssTerm term = fontSize.Terms.First();
                double size = 0;
                if (term.Unit == CssUnit.PT)
                {
                    if (double.TryParse(term.Value, out size))
                        return new TPoint(size);
                    return null;
                }
                return null;
            }
            return null;
        }

        private static string GetUsedFontFromFontFamilyProperty(CssExpression fontFamily)
        {
            if (fontFamily == null)
                return null;
            string fullFontFamily = fontFamily.Terms.Select(t => t + " ").StringConcatenate().Trim();
            string lcfont = fullFontFamily.ToLower();
            if (InstalledFonts.ContainsKey(lcfont))
                return InstalledFonts[lcfont];
            return null;
        }

        private static XElement GetBackgroundProperty(XElement element)
        {
            CssExpression color = element.GetProp("background-color");

            // todo this really should test against default background color
            if (color.ToString() != "transparent")
            {
                string hexString = color.ToString();
                XElement shd = new XElement(W.shd,
                    new XAttribute(W.val, "clear"),
                    new XAttribute(W.color, "auto"),
                    new XAttribute(W.fill, hexString));
                return shd;
            }
            return null;
        }

    }

    public class PictureId
    {
        public int Id;
    }

    class HtmlToWmlFontUpdater
    {
        public static void UpdateFontsPart(WordprocessingDocument wDoc, XElement html, HtmlToWmlConverterSettings settings)
        {
            XDocument fontXDoc = wDoc.MainDocumentPart.FontTablePart.GetXDocument();

            PtUtils.AddElementIfMissing(fontXDoc,
                fontXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "paragraph" && (string)e.Attribute(W.styleId) == "Heading1")
                    .FirstOrDefault(),
@"<w:font w:name='Verdana' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:panose1 w:val='020B0604030504040204'/>
  <w:charset w:val='00'/>
  <w:family w:val='swiss'/>
  <w:pitch w:val='variable'/>
  <w:sig w:usb0='A10006FF'
          w:usb1='4000205B'
          w:usb2='00000010'
          w:usb3='00000000'
          w:csb0='0000019F'
          w:csb1='00000000'/>
</w:font>");

            wDoc.MainDocumentPart.FontTablePart.PutXDocument();
        }
    }

    class NumberingUpdater
    {
        public static void InitializeNumberingPart(WordprocessingDocument wDoc)
        {
            NumberingDefinitionsPart numberingPart = wDoc.MainDocumentPart.NumberingDefinitionsPart;
            if (numberingPart == null)
            {
                wDoc.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                XDocument npXDoc = new XDocument(
                    new XElement(W.numbering,
                        new XAttribute(XNamespace.Xmlns + "w", W.w)));
                wDoc.MainDocumentPart.NumberingDefinitionsPart.PutXDocument(npXDoc);
            }
        }

        public static void GetNextNumId(WordprocessingDocument wDoc, out int nextNumId)
        {
            InitializeNumberingPart(wDoc);
            NumberingDefinitionsPart numberingPart = wDoc.MainDocumentPart.NumberingDefinitionsPart;
            XDocument numberingXDoc = numberingPart.GetXDocument();
            nextNumId = numberingXDoc.Root.Elements(W.num).Attributes(W.numId).Select(ni => (int)ni).Concat(new[] { 1 }).Max();
        }

        // decimal, lowerLetter
        private static string OrderedListAbstractXml =
@"<w:abstractNum w:abstractNumId='{0}' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
    <w:multiLevelType w:val='multilevel'/>
    <w:tmpl w:val='7D26959A'/>
    <w:lvl w:ilvl='0'>
      <w:start w:val='1'/>
      <w:numFmt w:val='{1}'/>
      <w:lvlText w:val='%1.'/>
      <w:lvlJc w:val='{2}'/>
      <w:pPr>
        <w:tabs>
          <w:tab w:val='num'
                 w:pos='720'/>
        </w:tabs>
        <w:ind w:left='720'
               w:hanging='360'/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl='1'
           w:tentative='1'>
      <w:start w:val='1'/>
      <w:numFmt w:val='{3}'/>
      <w:lvlText w:val='%2.'/>
      <w:lvlJc w:val='{4}'/>
      <w:pPr>
        <w:tabs>
          <w:tab w:val='num'
                 w:pos='1440'/>
        </w:tabs>
        <w:ind w:left='1440'
               w:hanging='360'/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl='2'
           w:tentative='1'>
      <w:start w:val='1'/>
      <w:numFmt w:val='{5}'/>
      <w:lvlText w:val='%3.'/>
      <w:lvlJc w:val='{6}'/>
      <w:pPr>
        <w:tabs>
          <w:tab w:val='num'
                 w:pos='2160'/>
        </w:tabs>
        <w:ind w:left='2160'
               w:hanging='360'/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl='3'
           w:tentative='1'>
      <w:start w:val='1'/>
      <w:numFmt w:val='{7}'/>
      <w:lvlText w:val='%4.'/>
      <w:lvlJc w:val='{8}'/>
      <w:pPr>
        <w:tabs>
          <w:tab w:val='num'
                 w:pos='2880'/>
        </w:tabs>
        <w:ind w:left='2880'
               w:hanging='360'/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl='4'
           w:tentative='1'>
      <w:start w:val='1'/>
      <w:numFmt w:val='{9}'/>
      <w:lvlText w:val='%5.'/>
      <w:lvlJc w:val='{10}'/>
      <w:pPr>
        <w:tabs>
          <w:tab w:val='num'
                 w:pos='3600'/>
        </w:tabs>
        <w:ind w:left='3600'
               w:hanging='360'/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl='5'
           w:tentative='1'>
      <w:start w:val='1'/>
      <w:numFmt w:val='{11}'/>
      <w:lvlText w:val='%6.'/>
      <w:lvlJc w:val='{12}'/>
      <w:pPr>
        <w:tabs>
          <w:tab w:val='num'
                 w:pos='4320'/>
        </w:tabs>
        <w:ind w:left='4320'
               w:hanging='360'/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl='6'
           w:tentative='1'>
      <w:start w:val='1'/>
      <w:numFmt w:val='{13}'/>
      <w:lvlText w:val='%7.'/>
      <w:lvlJc w:val='{14}'/>
      <w:pPr>
        <w:tabs>
          <w:tab w:val='num'
                 w:pos='5040'/>
        </w:tabs>
        <w:ind w:left='5040'
               w:hanging='360'/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl='7'
           w:tentative='1'>
      <w:start w:val='1'/>
      <w:numFmt w:val='{15}'/>
      <w:lvlText w:val='%8.'/>
      <w:lvlJc w:val='{16}'/>
      <w:pPr>
        <w:tabs>
          <w:tab w:val='num'
                 w:pos='5760'/>
        </w:tabs>
        <w:ind w:left='5760'
               w:hanging='360'/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl='8'
           w:tentative='1'>
      <w:start w:val='1'/>
      <w:numFmt w:val='{17}'/>
      <w:lvlText w:val='%9.'/>
      <w:lvlJc w:val='{18}'/>
      <w:pPr>
        <w:tabs>
          <w:tab w:val='num'
                 w:pos='6480'/>
        </w:tabs>
        <w:ind w:left='6480'
               w:hanging='360'/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>";

        private static string BulletAbstractXml =
@"<w:abstractNum w:abstractNumId='{0}' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:multiLevelType w:val='multilevel' />
  <w:tmpl w:val='02BEA0DA' />
  <w:lvl w:ilvl='0'>
    <w:start w:val='1' />
    <w:numFmt w:val='bullet' />
    <w:lvlText w:val='' />
    <w:lvlJc w:val='left' />
    <w:pPr>
      <w:tabs>
        <w:tab w:val='num' w:pos='720' />
      </w:tabs>
      <w:ind w:left='720' w:hanging='360' />
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii='Symbol' w:hAnsi='Symbol' w:hint='default' />
      <w:sz w:val='20' />
    </w:rPr>
  </w:lvl>
  <w:lvl w:ilvl='1'>
    <w:start w:val='1' />
    <w:numFmt w:val='bullet' />
    <w:lvlText w:val='o' />
    <w:lvlJc w:val='left' />
    <w:pPr>
      <w:tabs>
        <w:tab w:val='num' w:pos='1440' />
      </w:tabs>
      <w:ind w:left='1440' w:hanging='360' />
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii='Courier New' w:hAnsi='Courier New' w:hint='default' />
      <w:sz w:val='20' />
    </w:rPr>
  </w:lvl>
  <w:lvl w:ilvl='2' w:tentative='1'>
    <w:start w:val='1' />
    <w:numFmt w:val='bullet' />
    <w:lvlText w:val='' />
    <w:lvlJc w:val='left' />
    <w:pPr>
      <w:tabs>
        <w:tab w:val='num' w:pos='2160' />
      </w:tabs>
      <w:ind w:left='2160' w:hanging='360' />
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii='Wingdings' w:hAnsi='Wingdings' w:hint='default' />
      <w:sz w:val='20' />
    </w:rPr>
  </w:lvl>
  <w:lvl w:ilvl='3' w:tentative='1'>
    <w:start w:val='1' />
    <w:numFmt w:val='bullet' />
    <w:lvlText w:val='' />
    <w:lvlJc w:val='left' />
    <w:pPr>
      <w:tabs>
        <w:tab w:val='num' w:pos='2880' />
      </w:tabs>
      <w:ind w:left='2880' w:hanging='360' />
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii='Wingdings' w:hAnsi='Wingdings' w:hint='default' />
      <w:sz w:val='20' />
    </w:rPr>
  </w:lvl>
  <w:lvl w:ilvl='4' w:tentative='1'>
    <w:start w:val='1' />
    <w:numFmt w:val='bullet' />
    <w:lvlText w:val='' />
    <w:lvlJc w:val='left' />
    <w:pPr>
      <w:tabs>
        <w:tab w:val='num' w:pos='3600' />
      </w:tabs>
      <w:ind w:left='3600' w:hanging='360' />
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii='Wingdings' w:hAnsi='Wingdings' w:hint='default' />
      <w:sz w:val='20' />
    </w:rPr>
  </w:lvl>
  <w:lvl w:ilvl='5' w:tentative='1'>
    <w:start w:val='1' />
    <w:numFmt w:val='bullet' />
    <w:lvlText w:val='' />
    <w:lvlJc w:val='left' />
    <w:pPr>
      <w:tabs>
        <w:tab w:val='num' w:pos='4320' />
      </w:tabs>
      <w:ind w:left='4320' w:hanging='360' />
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii='Wingdings' w:hAnsi='Wingdings' w:hint='default' />
      <w:sz w:val='20' />
    </w:rPr>
  </w:lvl>
  <w:lvl w:ilvl='6' w:tentative='1'>
    <w:start w:val='1' />
    <w:numFmt w:val='bullet' />
    <w:lvlText w:val='' />
    <w:lvlJc w:val='left' />
    <w:pPr>
      <w:tabs>
        <w:tab w:val='num' w:pos='5040' />
      </w:tabs>
      <w:ind w:left='5040' w:hanging='360' />
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii='Wingdings' w:hAnsi='Wingdings' w:hint='default' />
      <w:sz w:val='20' />
    </w:rPr>
  </w:lvl>
  <w:lvl w:ilvl='7' w:tentative='1'>
    <w:start w:val='1' />
    <w:numFmt w:val='bullet' />
    <w:lvlText w:val='' />
    <w:lvlJc w:val='left' />
    <w:pPr>
      <w:tabs>
        <w:tab w:val='num' w:pos='5760' />
      </w:tabs>
      <w:ind w:left='5760' w:hanging='360' />
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii='Wingdings' w:hAnsi='Wingdings' w:hint='default' />
      <w:sz w:val='20' />
    </w:rPr>
  </w:lvl>
  <w:lvl w:ilvl='8' w:tentative='1'>
    <w:start w:val='1' />
    <w:numFmt w:val='bullet' />
    <w:lvlText w:val='' />
    <w:lvlJc w:val='left' />
    <w:pPr>
      <w:tabs>
        <w:tab w:val='num' w:pos='6480' />
      </w:tabs>
      <w:ind w:left='6480' w:hanging='360' />
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii='Wingdings' w:hAnsi='Wingdings' w:hint='default' />
      <w:sz w:val='20' />
    </w:rPr>
  </w:lvl>
</w:abstractNum>";

        public static void UpdateNumberingPart(WordprocessingDocument wDoc, XElement html, HtmlToWmlConverterSettings settings)
        {
            InitializeNumberingPart(wDoc);
            NumberingDefinitionsPart numberingPart = wDoc.MainDocumentPart.NumberingDefinitionsPart;
            XDocument numberingXDoc = numberingPart.GetXDocument();
            int nextAbstractId, nextNumId;
            nextNumId = numberingXDoc.Root.Elements(W.num).Attributes(W.numId).Select(ni => (int)ni).Concat(new[] { 1 }).Max();
            nextAbstractId = numberingXDoc.Root.Elements(W.abstractNum).Attributes(W.abstractNumId).Select(ani => (int)ani).Concat(new[] { 0 }).Max();
            var numberingElements = html.DescendantsAndSelf().Where(d => d.Name == XhtmlNoNamespace.ol || d.Name == XhtmlNoNamespace.ul).ToList();

            Dictionary<int, int> numToAbstractNum = new Dictionary<int, int>();

            // add abstract numbering elements
            int currentNumId = nextNumId;
            int currentAbstractId = nextAbstractId;
            foreach (var list in numberingElements)
            {
                HtmlToWmlConverterCore.NumberedItemAnnotation nia = list.Annotation<HtmlToWmlConverterCore.NumberedItemAnnotation>();
                if (!numToAbstractNum.ContainsKey(nia.numId))
                {
                    numToAbstractNum.Add(nia.numId, currentAbstractId);
                    if (list.Name == XhtmlNoNamespace.ul)
                    {
                        XElement bulletAbstract = XElement.Parse(String.Format(BulletAbstractXml, currentAbstractId++));
                        numberingXDoc.Root.Add(bulletAbstract);
                    }
                    if (list.Name == XhtmlNoNamespace.ol)
                    {
                        string[] numFmt = new string[9];
                        string[] just = new string[9];
                        for (int i = 0; i < numFmt.Length; ++i)
                        {
                            numFmt[i] = "decimal";
                            just[i] = "left";
                            XElement itemAtLevel = numberingElements
                                .FirstOrDefault(nf =>
                                {
                                    HtmlToWmlConverterCore.NumberedItemAnnotation n = nf.Annotation<HtmlToWmlConverterCore.NumberedItemAnnotation>();
                                    if (n != null && n.numId == nia.numId && n.ilvl == i)
                                        return true;
                                    return false;
                                });
                            if (itemAtLevel != null)
                            {
                                HtmlToWmlConverterCore.NumberedItemAnnotation thisLevelNia = itemAtLevel.Annotation<HtmlToWmlConverterCore.NumberedItemAnnotation>();
                                string thisLevelNumFmt = thisLevelNia.listStyleType;
                                if (thisLevelNumFmt == "lower-alpha" || thisLevelNumFmt == "lower-latin")
                                {
                                    numFmt[i] = "lowerLetter";
                                    //just[i] = "left";
                                }
                                if (thisLevelNumFmt == "upper-alpha" || thisLevelNumFmt == "upper-latin")
                                {
                                    numFmt[i] = "upperLetter";
                                    //just[i] = "left";
                                }
                                if (thisLevelNumFmt == "decimal-leading-zero")
                                {
                                    numFmt[i] = "decimalZero";
                                    //just[i] = "left";
                                }
                                if (thisLevelNumFmt == "lower-roman")
                                {
                                    numFmt[i] = "lowerRoman";
                                    just[i] = "right";
                                }
                                if (thisLevelNumFmt == "upper-roman")
                                {
                                    numFmt[i] = "upperRoman";
                                    just[i] = "right";
                                }
                            }
                        }

                        XElement simpleNumAbstract = XElement.Parse(String.Format(OrderedListAbstractXml, currentAbstractId++,
                            numFmt[0], just[0], numFmt[1], just[1], numFmt[2], just[2], numFmt[3], just[3], numFmt[4], just[4], numFmt[5], just[5], numFmt[6], just[6], numFmt[7], just[7], numFmt[8], just[8]));
                        numberingXDoc.Root.Add(simpleNumAbstract);
                    }
                }
            }

            foreach (var list in numToAbstractNum)
            {
                numberingXDoc.Root.Add(
                    new XElement(W.num, new XAttribute(W.numId, list.Key),
                    new XElement(W.abstractNumId, new XAttribute(W.val, list.Value))));
            }

            wDoc.MainDocumentPart.NumberingDefinitionsPart.PutXDocument();
#if false
  <w:num w:numId='1'>
    <w:abstractNumId w:val='0'/>
  </w:num>
#endif
        }
    }

    class StylesUpdater
    {
        public static void UpdateStylesPart(
            WordprocessingDocument wDoc,
            XElement html,
            HtmlToWmlConverterSettings settings,
            CssDocument defaultCssDoc,
            CssDocument authorCssDoc,
            CssDocument userCssDoc)
        {
            XDocument styleXDoc = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument();

            if (settings.DefaultSpacingElement != null)
            {
                XElement spacingElement = styleXDoc.Root.Elements(W.docDefaults).Elements(W.pPrDefault).Elements(W.pPr).Elements(W.spacing).FirstOrDefault();
                if (spacingElement != null)
                    spacingElement.ReplaceWith(settings.DefaultSpacingElement);
            }

            var classes = html
                .DescendantsAndSelf()
                .Where(d => d.Attribute(XhtmlNoNamespace._class) != null && ((string)d.Attribute(XhtmlNoNamespace._class)).Split().Length == 1)
                .Select(d => (string)d.Attribute(XhtmlNoNamespace._class));

            foreach (var item in classes)
            {
                //string item = "ms-rteStyle-Byline";
                foreach (var ruleSet in authorCssDoc.RuleSets)
                {
                    var selector = ruleSet.Selectors.Where(
                        sel =>
                        {
                            bool found = sel.SimpleSelectors.Count() == 1 &&
                                sel.SimpleSelectors.First().Class == item &&
                                (sel.SimpleSelectors.First().ElementName == "" ||
                                sel.SimpleSelectors.First().ElementName == null);
                            return found;
                        }).FirstOrDefault();
                    var color = ruleSet.Declarations.FirstOrDefault(d => d.Name == "color");
                    if (selector != null)
                    {
                        //Console.WriteLine("found ruleset and selector for {0}", item);
                        string styleName = item.ToLower();
                        XElement newStyle = new XElement(W.style,
                            new XAttribute(W.type, "paragraph"),
                            new XAttribute(W.customStyle, "1"),
                            new XAttribute(W.styleId, styleName),
                            new XElement(W.name, new XAttribute(W.val, styleName)),
                            new XElement(W.basedOn, new XAttribute(W.val, "Normal")),
                            new XElement(W.pPr,
                                new XElement(W.spacing, new XAttribute(W.before, "100"),
                                    new XAttribute(W.beforeAutospacing, "1"),
                                    new XAttribute(W.after, "100"),
                                    new XAttribute(W.afterAutospacing, "1"),
                                    new XAttribute(W.line, "240"),
                                    new XAttribute(W.lineRule, "auto"))),
                            new XElement(W.rPr,
                                new XElement(W.rFonts, new XAttribute(W.ascii, "Times New Roman"),
                                    new XAttribute(W.eastAsiaTheme, "minorEastAsia"),
                                    new XAttribute(W.hAnsi, "Times New Roman"),
                                    new XAttribute(W.cs, "Times New Roman")),
                                color != null ? new XElement(W.color, new XAttribute(W.val, "this should be a color")) : null,
                                new XElement(W.sz, new XAttribute(W.val, "24")),
                                new XElement(W.szCs, new XAttribute(W.val, "24"))));
                        if (styleXDoc
                            .Root
                            .Elements(W.style)
                            .Where(e => (string)e.Attribute(W.type) == "paragraph" && (string)e.Attribute(W.styleId) == styleName)
                            .FirstOrDefault() == null)
                            styleXDoc.Root.Add(newStyle);
                    }
                }
            }

            if (html.Descendants(XhtmlNoNamespace.h1).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                    styleXDoc
                        .Root
                        .Elements(W.style)
                        .Where(e => (string)e.Attribute(W.type) == "paragraph" && (string)e.Attribute(W.styleId) == "Heading1")
                        .FirstOrDefault(),
@"<w:style w:type='paragraph'
        w:styleId='Heading1'
        xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
    <w:name w:val='heading 1'/>
    <w:basedOn w:val='Normal'/>
    <w:next w:val='Normal'/>
    <w:link w:val='Heading1Char'/>
    <w:uiPriority w:val='9'/>
    <w:qFormat/>
    <w:pPr>
    <w:keepNext/>
    <w:keepLines/>
    <w:spacing w:before='480'
                w:after='0'/>
    <w:outlineLvl w:val='0'/>
    </w:pPr>
    <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
                w:eastAsiaTheme='majorEastAsia'
                w:hAnsiTheme='majorHAnsi'
                w:cstheme='majorBidi'/>
    <w:b/>
    <w:bCs/>
    <w:color w:val='365F91'
                w:themeColor='accent1'
                w:themeShade='BF'/>
    <w:sz w:val='28'/>
    <w:szCs w:val='28'/>
    </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.h2).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                    styleXDoc
                        .Root
                        .Elements(W.style)
                        .Where(e => (string)e.Attribute(W.type) == "paragraph" && (string)e.Attribute(W.styleId) == "Heading2")
                        .FirstOrDefault(),
@"<w:style w:type='paragraph'
         w:styleId='Heading2'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='heading 2'/>
  <w:basedOn w:val='Normal'/>
  <w:next w:val='Normal'/>
  <w:link w:val='Heading2Char'/>
  <w:uiPriority w:val='9'/>
  <w:unhideWhenUsed/>
  <w:qFormat/>
  <w:pPr>
    <w:keepNext/>
    <w:keepLines/>
    <w:spacing w:before='200'
               w:after='0'/>
    <w:outlineLvl w:val='1'/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
              w:eastAsiaTheme='majorEastAsia'
              w:hAnsiTheme='majorHAnsi'
              w:cstheme='majorBidi'/>
    <w:b/>
    <w:bCs/>
    <w:color w:val='4F81BD'
             w:themeColor='accent1'/>
    <w:sz w:val='26'/>
    <w:szCs w:val='26'/>
  </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.h3).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                styleXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "paragraph" && (string)e.Attribute(W.styleId) == "Heading3")
                    .FirstOrDefault(),
@"<w:style w:type='paragraph'
         w:styleId='Heading3'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='heading 3'/>
  <w:basedOn w:val='Normal'/>
  <w:next w:val='Normal'/>
  <w:link w:val='Heading3Char'/>
  <w:uiPriority w:val='9'/>
  <w:unhideWhenUsed/>
  <w:qFormat/>
  <w:pPr>
    <w:keepNext/>
    <w:keepLines/>
    <w:spacing w:before='200'
               w:after='0'/>
    <w:outlineLvl w:val='2'/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
              w:eastAsiaTheme='majorEastAsia'
              w:hAnsiTheme='majorHAnsi'
              w:cstheme='majorBidi'/>
    <w:b/>
    <w:bCs/>
    <w:color w:val='4F81BD'
             w:themeColor='accent1'/>
  </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.h4).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                styleXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "paragraph" && (string)e.Attribute(W.styleId) == "Heading4")
                    .FirstOrDefault(),
@"<w:style w:type='paragraph'
         w:styleId='Heading4'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='heading 4'/>
  <w:basedOn w:val='Normal'/>
  <w:next w:val='Normal'/>
  <w:link w:val='Heading4Char'/>
  <w:uiPriority w:val='9'/>
  <w:unhideWhenUsed/>
  <w:qFormat/>
  <w:pPr>
    <w:keepNext/>
    <w:keepLines/>
    <w:spacing w:before='200'
               w:after='0'/>
    <w:outlineLvl w:val='3'/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
              w:eastAsiaTheme='majorEastAsia'
              w:hAnsiTheme='majorHAnsi'
              w:cstheme='majorBidi'/>
    <w:b/>
    <w:bCs/>
    <w:i/>
    <w:iCs/>
    <w:color w:val='4F81BD'
             w:themeColor='accent1'/>
  </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.h5).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                styleXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "paragraph" && (string)e.Attribute(W.styleId) == "Heading5")
                    .FirstOrDefault(),
@"<w:style w:type='paragraph'
         w:styleId='Heading5'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='heading 5'/>
  <w:basedOn w:val='Normal'/>
  <w:next w:val='Normal'/>
  <w:link w:val='Heading5Char'/>
  <w:uiPriority w:val='9'/>
  <w:unhideWhenUsed/>
  <w:qFormat/>
  <w:pPr>
    <w:keepNext/>
    <w:keepLines/>
    <w:spacing w:before='200'
               w:after='0'/>
    <w:outlineLvl w:val='4'/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
              w:eastAsiaTheme='majorEastAsia'
              w:hAnsiTheme='majorHAnsi'
              w:cstheme='majorBidi'/>
    <w:color w:val='243F60'
             w:themeColor='accent1'
             w:themeShade='7F'/>
  </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.h6).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                styleXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "paragraph" && (string)e.Attribute(W.styleId) == "Heading6")
                    .FirstOrDefault(),
@"<w:style w:type='paragraph'
         w:styleId='Heading6'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='heading 6'/>
  <w:basedOn w:val='Normal'/>
  <w:next w:val='Normal'/>
  <w:link w:val='Heading6Char'/>
  <w:uiPriority w:val='9'/>
  <w:unhideWhenUsed/>
  <w:qFormat/>
  <w:pPr>
    <w:keepNext/>
    <w:keepLines/>
    <w:spacing w:before='200'
               w:after='0'/>
    <w:outlineLvl w:val='5'/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
              w:eastAsiaTheme='majorEastAsia'
              w:hAnsiTheme='majorHAnsi'
              w:cstheme='majorBidi'/>
    <w:i/>
    <w:iCs/>
    <w:color w:val='243F60'
             w:themeColor='accent1'
             w:themeShade='7F'/>
  </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.h7).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                styleXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "paragraph" && (string)e.Attribute(W.styleId) == "Heading7")
                    .FirstOrDefault(),
@"<w:style w:type='paragraph'
         w:styleId='Heading7'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='heading 7'/>
  <w:basedOn w:val='Normal'/>
  <w:next w:val='Normal'/>
  <w:link w:val='Heading7Char'/>
  <w:uiPriority w:val='9'/>
  <w:unhideWhenUsed/>
  <w:qFormat/>
  <w:pPr>
    <w:keepNext/>
    <w:keepLines/>
    <w:spacing w:before='200'
               w:after='0'/>
    <w:outlineLvl w:val='6'/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
              w:eastAsiaTheme='majorEastAsia'
              w:hAnsiTheme='majorHAnsi'
              w:cstheme='majorBidi'/>
    <w:i/>
    <w:iCs/>
    <w:color w:val='404040'
             w:themeColor='text1'
             w:themeTint='BF'/>
  </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.h8).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                styleXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "paragraph" && (string)e.Attribute(W.styleId) == "Heading8")
                    .FirstOrDefault(),
@"<w:style w:type='paragraph'
         w:styleId='Heading8'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='heading 8'/>
  <w:basedOn w:val='Normal'/>
  <w:next w:val='Normal'/>
  <w:link w:val='Heading8Char'/>
  <w:uiPriority w:val='9'/>
  <w:unhideWhenUsed/>
  <w:qFormat/>
  <w:pPr>
    <w:keepNext/>
    <w:keepLines/>
    <w:spacing w:before='200'
               w:after='0'/>
    <w:outlineLvl w:val='7'/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
              w:eastAsiaTheme='majorEastAsia'
              w:hAnsiTheme='majorHAnsi'
              w:cstheme='majorBidi'/>
    <w:color w:val='404040'
             w:themeColor='text1'
             w:themeTint='BF'/>
    <w:sz w:val='20'/>
    <w:szCs w:val='20'/>
  </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.h9).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                styleXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "paragraph" && (string)e.Attribute(W.styleId) == "Heading9")
                    .FirstOrDefault(),
@"<w:style w:type='paragraph'
         w:styleId='Heading9'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='heading 9'/>
  <w:basedOn w:val='Normal'/>
  <w:next w:val='Normal'/>
  <w:link w:val='Heading9Char'/>
  <w:uiPriority w:val='9'/>
  <w:unhideWhenUsed/>
  <w:qFormat/>
  <w:pPr>
    <w:keepNext/>
    <w:keepLines/>
    <w:spacing w:before='200'
               w:after='0'/>
    <w:outlineLvl w:val='8'/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
              w:eastAsiaTheme='majorEastAsia'
              w:hAnsiTheme='majorHAnsi'
              w:cstheme='majorBidi'/>
    <w:i/>
    <w:iCs/>
    <w:color w:val='404040'
             w:themeColor='text1'
             w:themeTint='BF'/>
    <w:sz w:val='20'/>
    <w:szCs w:val='20'/>
  </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.h1).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                styleXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "character" && (string)e.Attribute(W.styleId) == "Heading1Char")
                    .FirstOrDefault(),
@"<w:style w:type='character'
         w:customStyle='1'
         w:styleId='Heading1Char'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='Heading 1 Char'/>
  <w:basedOn w:val='DefaultParagraphFont'/>
  <w:link w:val='Heading1'/>
  <w:uiPriority w:val='9'/>
  <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
              w:eastAsiaTheme='majorEastAsia'
              w:hAnsiTheme='majorHAnsi'
              w:cstheme='majorBidi'/>
    <w:b/>
    <w:bCs/>
    <w:color w:val='365F91'
             w:themeColor='accent1'
             w:themeShade='BF'/>
    <w:sz w:val='28'/>
    <w:szCs w:val='28'/>
  </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.h2).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                styleXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "character" && (string)e.Attribute(W.styleId) == "Heading2Char")
                    .FirstOrDefault(),
@"<w:style w:type='character'
         w:customStyle='1'
         w:styleId='Heading2Char'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='Heading 2 Char'/>
  <w:basedOn w:val='DefaultParagraphFont'/>
  <w:link w:val='Heading2'/>
  <w:uiPriority w:val='9'/>
  <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
              w:eastAsiaTheme='majorEastAsia'
              w:hAnsiTheme='majorHAnsi'
              w:cstheme='majorBidi'/>
    <w:b/>
    <w:bCs/>
    <w:color w:val='4F81BD'
             w:themeColor='accent1'/>
    <w:sz w:val='26'/>
    <w:szCs w:val='26'/>
  </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.h3).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                styleXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "character" && (string)e.Attribute(W.styleId) == "Heading3Char")
                    .FirstOrDefault(),
@"<w:style w:type='character'
         w:customStyle='1'
         w:styleId='Heading3Char'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='Heading 3 Char'/>
  <w:basedOn w:val='DefaultParagraphFont'/>
  <w:link w:val='Heading3'/>
  <w:uiPriority w:val='9'/>
  <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
              w:eastAsiaTheme='majorEastAsia'
              w:hAnsiTheme='majorHAnsi'
              w:cstheme='majorBidi'/>
    <w:b/>
    <w:bCs/>
    <w:color w:val='4F81BD'
             w:themeColor='accent1'/>
  </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.h4).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                styleXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "character" && (string)e.Attribute(W.styleId) == "Heading4Char")
                    .FirstOrDefault(),
@"<w:style w:type='character'
         w:customStyle='1'
         w:styleId='Heading4Char'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='Heading 4 Char'/>
  <w:basedOn w:val='DefaultParagraphFont'/>
  <w:link w:val='Heading4'/>
  <w:uiPriority w:val='9'/>
  <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
              w:eastAsiaTheme='majorEastAsia'
              w:hAnsiTheme='majorHAnsi'
              w:cstheme='majorBidi'/>
    <w:b/>
    <w:bCs/>
    <w:i/>
    <w:iCs/>
    <w:color w:val='4F81BD'
             w:themeColor='accent1'/>
  </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.h5).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                styleXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "character" && (string)e.Attribute(W.styleId) == "Heading5Char")
                    .FirstOrDefault(),
@"<w:style w:type='character'
         w:customStyle='1'
         w:styleId='Heading5Char'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='Heading 5 Char'/>
  <w:basedOn w:val='DefaultParagraphFont'/>
  <w:link w:val='Heading5'/>
  <w:uiPriority w:val='9'/>
  <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
              w:eastAsiaTheme='majorEastAsia'
              w:hAnsiTheme='majorHAnsi'
              w:cstheme='majorBidi'/>
    <w:color w:val='243F60'
             w:themeColor='accent1'
             w:themeShade='7F'/>
  </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.h6).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                styleXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "character" && (string)e.Attribute(W.styleId) == "Heading6Char")
                    .FirstOrDefault(),
@"<w:style w:type='character'
         w:customStyle='1'
         w:styleId='Heading6Char'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='Heading 6 Char'/>
  <w:basedOn w:val='DefaultParagraphFont'/>
  <w:link w:val='Heading6'/>
  <w:uiPriority w:val='9'/>
  <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
              w:eastAsiaTheme='majorEastAsia'
              w:hAnsiTheme='majorHAnsi'
              w:cstheme='majorBidi'/>
    <w:i/>
    <w:iCs/>
    <w:color w:val='243F60'
             w:themeColor='accent1'
             w:themeShade='7F'/>
  </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.h7).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                styleXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "character" && (string)e.Attribute(W.styleId) == "Heading7Char")
                    .FirstOrDefault(),
@"<w:style w:type='character'
         w:customStyle='1'
         w:styleId='Heading7Char'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='Heading 7 Char'/>
  <w:basedOn w:val='DefaultParagraphFont'/>
  <w:link w:val='Heading7'/>
  <w:uiPriority w:val='9'/>
  <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
              w:eastAsiaTheme='majorEastAsia'
              w:hAnsiTheme='majorHAnsi'
              w:cstheme='majorBidi'/>
    <w:i/>
    <w:iCs/>
    <w:color w:val='404040'
             w:themeColor='text1'
             w:themeTint='BF'/>
  </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.h8).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                styleXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "character" && (string)e.Attribute(W.styleId) == "Heading8Char")
                    .FirstOrDefault(),
@"<w:style w:type='character'
         w:customStyle='1'
         w:styleId='Heading8Char'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='Heading 8 Char'/>
  <w:basedOn w:val='DefaultParagraphFont'/>
  <w:link w:val='Heading8'/>
  <w:uiPriority w:val='9'/>
  <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
              w:eastAsiaTheme='majorEastAsia'
              w:hAnsiTheme='majorHAnsi'
              w:cstheme='majorBidi'/>
    <w:color w:val='404040'
             w:themeColor='text1'
             w:themeTint='BF'/>
    <w:sz w:val='20'/>
    <w:szCs w:val='20'/>
  </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.h9).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                styleXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "character" && (string)e.Attribute(W.styleId) == "Heading9Char")
                    .FirstOrDefault(),
@"<w:style w:type='character'
         w:customStyle='1'
         w:styleId='Heading9Char'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='Heading 9 Char'/>
  <w:basedOn w:val='DefaultParagraphFont'/>
  <w:link w:val='Heading9'/>
  <w:uiPriority w:val='9'/>
  <w:rPr>
    <w:rFonts w:asciiTheme='majorHAnsi'
              w:eastAsiaTheme='majorEastAsia'
              w:hAnsiTheme='majorHAnsi'
              w:cstheme='majorBidi'/>
    <w:i/>
    <w:iCs/>
    <w:color w:val='404040'
             w:themeColor='text1'
             w:themeTint='BF'/>
    <w:sz w:val='20'/>
    <w:szCs w:val='20'/>
  </w:rPr>
</w:style>");

            if (html.Descendants(XhtmlNoNamespace.a).Any())
                PtUtils.AddElementIfMissing(styleXDoc,
                styleXDoc
                    .Root
                    .Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "character" && (string)e.Attribute(W.styleId) == "Hyperlink")
                    .FirstOrDefault(),
@"<w:style w:type='character'
         w:styleId='Hyperlink'
         xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:name w:val='Hyperlink' />
  <w:basedOn w:val='DefaultParagraphFont' />
  <w:uiPriority w:val='99' />
  <w:semiHidden />
  <w:unhideWhenUsed />
  <w:rPr>
    <w:color w:val='0000FF' />
    <w:u w:val='single' />
  </w:rPr>
</w:style>");

            wDoc.MainDocumentPart.StyleDefinitionsPart.PutXDocument();
        }
    }

    class ThemeUpdater
    {
        public static void UpdateThemePart(WordprocessingDocument wDoc, XElement html, HtmlToWmlConverterSettings settings)
        {
            XDocument themeXDoc = wDoc.MainDocumentPart.ThemePart.GetXDocument();

            CssExpression minorFont = html.Descendants(XhtmlNoNamespace.body).FirstOrDefault().GetProp("font-family");
            XElement majorFontElement = html.Descendants().Where(e =>
                e.Name == XhtmlNoNamespace.h1 ||
                e.Name == XhtmlNoNamespace.h2 ||
                e.Name == XhtmlNoNamespace.h3 ||
                e.Name == XhtmlNoNamespace.h4 ||
                e.Name == XhtmlNoNamespace.h5 ||
                e.Name == XhtmlNoNamespace.h6 ||
                e.Name == XhtmlNoNamespace.h7 ||
                e.Name == XhtmlNoNamespace.h8 ||
                e.Name == XhtmlNoNamespace.h9).FirstOrDefault();
            CssExpression majorFont = null;
            if (majorFontElement != null)
                majorFont = majorFontElement.GetProp("font-family");

            XAttribute majorTypeface = themeXDoc
                .Root
                .Elements(A.themeElements)
                .Elements(A.fontScheme)
                .Elements(A.majorFont)
                .Elements(A.latin)
                .Attributes(NoNamespace.typeface)
                .FirstOrDefault();
            if (majorTypeface != null && majorFont != null)
            {
                CssTerm term = majorFont.Terms.FirstOrDefault();
                if (term != null)
                    majorTypeface.Value = term.Value;
            }
            XAttribute minorTypeface = themeXDoc
                .Root
                .Elements(A.themeElements)
                .Elements(A.fontScheme)
                .Elements(A.minorFont)
                .Elements(A.latin)
                .Attributes(NoNamespace.typeface)
                .FirstOrDefault();
            if (minorTypeface != null && minorFont != null)
            {
                CssTerm term = minorFont.Terms.FirstOrDefault();
                if (term != null)
                    minorTypeface.Value = term.Value;
            }

            wDoc.MainDocumentPart.ThemePart.PutXDocument();
        }
    }
}
