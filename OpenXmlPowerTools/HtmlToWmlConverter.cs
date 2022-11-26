// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

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

namespace OpenXmlPowerTools
{
    public class HtmlToWmlConverterSettings
    {
        public string MajorLatinFont;
        public string MinorLatinFont;
        public double DefaultFontSize;
        public XElement DefaultSpacingElement;
        public XElement DefaultSpacingElementForParagraphsInTables;
        public XElement SectPr;
        public string DefaultBlockContentMargin;
        public string BaseUriForImages;

        public Twip PageWidthTwips { get { return (long)SectPr.Elements(W.pgSz).Attributes(W._w).FirstOrDefault(); } }
        public Twip PageMarginLeftTwips { get { return (long)SectPr.Elements(W.pgMar).Attributes(W.left).FirstOrDefault(); } }
        public Twip PageMarginRightTwips { get { return (long)SectPr.Elements(W.pgMar).Attributes(W.right).FirstOrDefault(); } }
        public Emu PageWidthEmus { get { return Emu.TwipsToEmus(PageWidthTwips); } }
        public Emu PageMarginLeftEmus { get { return Emu.TwipsToEmus(PageMarginLeftTwips); } }
        public Emu PageMarginRightEmus { get { return Emu.TwipsToEmus(PageMarginRightTwips); } }
    }

    public class HtmlToWmlConverter
    {
        public static WmlDocument ConvertHtmlToWml(
            string defaultCss,
            string authorCss,
            string userCss,
            XElement xhtml,
            HtmlToWmlConverterSettings settings)
        {
            return HtmlToWmlConverterCore.ConvertHtmlToWml(defaultCss, authorCss, userCss, xhtml, settings, null, null);
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
            return HtmlToWmlConverterCore.ConvertHtmlToWml(defaultCss, authorCss, userCss, xhtml, settings, emptyDocument, annotatedHtmlDumpFileName);
        }

        private static string s_Blank_wml_base64 = @"UEsDBBQABgAIAAAAIQAJJIeCgQEAAI4FAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAAC
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC0
lE1Pg0AQhu8m/geyVwPbejDGlPag9ahNrPG8LkPZyH5kZ/v17x1KS6qhpVq9kMAy7/vMCzOD0UqX
0QI8KmtS1k96LAIjbabMLGWv08f4lkUYhMlEaQ2kbA3IRsPLi8F07QAjqjaYsiIEd8c5ygK0wMQ6
MHSSW69FoFs/407IDzEDft3r3XBpTQAT4lBpsOHgAXIxL0M0XtHjmsRDiSy6r1+svFImnCuVFIFI
+cJk31zirUNClZt3sFAOrwiD8VaH6uSwwbbumaLxKoNoInx4Epow+NL6jGdWzjX1kByXaeG0ea4k
NPWVmvNWAiJlrsukOdFCmR3/QQ4M6xLw7ylq3RPt31QoxnkOkj52dx4a46rppLbYq+12gxAopFNM
vv6CcVfouFXuRFjC+8u/UeyJd4LkNBpT8V7CCYn/MIxGuhMi0LwD31z7Z3NsZI5Z0mRMvHVI+8P/
ou3dgqiqYxo5Bz4oaFZE24g1jrR7zu4Pqu2WQdbizTfbdPgJAAD//wMAUEsDBBQABgAIAAAAIQAe
kRq38wAAAE4CAAALAAgCX3JlbHMvLnJlbHMgogQCKKAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjJLbSgNBDIbvBd9hyH032woi0tneSKF3
IusDhJnsAXcOzKTavr2jILpQ217m9OfLT9abg5vUO6c8Bq9hWdWg2JtgR99reG23iwdQWchbmoJn
DUfOsGlub9YvPJGUoTyMMaui4rOGQSQ+ImYzsKNchci+VLqQHEkJU4+RzBv1jKu6vsf0VwOamaba
WQ1pZ+9AtcdYNl/WDl03Gn4KZu/Yy4kVyAdhb9kuYipsScZyjWop9SwabDDPJZ2RYqwKNuBpotX1
RP9fi46FLAmhCYnP83x1nANaXg902aJ5x687HyFZLBZ9e/tDg7MvaD4BAAD//wMAUEsDBBQABgAI
AAAAIQB8O5c5IgEAALkDAAAcAAgBd29yZC9fcmVscy9kb2N1bWVudC54bWwucmVscyCiBAEooAAB
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";

        private static WmlDocument s_EmptyDocument = null;

        public static WmlDocument EmptyDocument
        {
            get {
                if (s_EmptyDocument == null)
                {
                    s_EmptyDocument = new WmlDocument("EmptyDocument.docx", Convert.FromBase64String(s_Blank_wml_base64));
                }
                return s_EmptyDocument;
            }
        }

        public static HtmlToWmlConverterSettings GetDefaultSettings()
        {
            return GetDefaultSettings(EmptyDocument);
        }

        public static HtmlToWmlConverterSettings GetDefaultSettings(WmlDocument wmlDocument)
        {
            HtmlToWmlConverterSettings settings = new HtmlToWmlConverterSettings();
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(wmlDocument.DocumentByteArray, 0, wmlDocument.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, false))
                {
                    string majorLatinFont, minorLatinFont;
                    double defaultFontSize;
                    GetDefaultFontInfo(wDoc, out majorLatinFont, out minorLatinFont, out defaultFontSize);
                    settings.MajorLatinFont = majorLatinFont;
                    settings.MinorLatinFont = minorLatinFont;
                    settings.DefaultFontSize = defaultFontSize;

                    settings.MinorLatinFont = "Times New Roman";
                    settings.DefaultFontSize = 12d;
                    settings.DefaultBlockContentMargin = "auto";
                    settings.DefaultSpacingElement = new XElement(W.spacing,
                        new XAttribute(W.before, 100),
                        new XAttribute(W.beforeAutospacing, 1),
                        new XAttribute(W.after, 100),
                        new XAttribute(W.afterAutospacing, 1),
                        new XAttribute(W.line, 240),
                        new XAttribute(W.lineRule, "auto"));
                    settings.DefaultSpacingElementForParagraphsInTables = new XElement(W.spacing,
                        new XAttribute(W.before, 100),
                        new XAttribute(W.beforeAutospacing, 1),
                        new XAttribute(W.after, 100),
                        new XAttribute(W.afterAutospacing, 1),
                        new XAttribute(W.line, 240),
                        new XAttribute(W.lineRule, "auto"));

                    XDocument mXDoc = wDoc.MainDocumentPart.GetXDocument();
                    XElement existingSectPr = mXDoc.Root.Descendants(W.sectPr).FirstOrDefault();
                    settings.SectPr = new XElement(W.sectPr,
                        existingSectPr.Elements(W.pgSz),
                        existingSectPr.Elements(W.pgMar));
                }
            }
            return settings;
        }

        private static void GetDefaultFontInfo(WordprocessingDocument wDoc, out string majorLatinFont, out string minorLatinFont, out double defaultFontSize)
        {
            if (wDoc.MainDocumentPart.ThemePart != null)
            {
                XElement fontScheme = wDoc.MainDocumentPart.ThemePart.GetXDocument().Root.Elements(A.themeElements).Elements(A.fontScheme).FirstOrDefault();
                if (fontScheme != null)
                {
                    majorLatinFont = (string)fontScheme.Elements(A.majorFont).Elements(A.latin).Attributes(NoNamespace.typeface).FirstOrDefault();
                    minorLatinFont = (string)fontScheme.Elements(A.minorFont).Elements(A.latin).Attributes(NoNamespace.typeface).FirstOrDefault();
                    string defaultFontSizeString = (string)wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument().Root.Elements(W.docDefaults)
                        .Elements(W.rPrDefault).Elements(W.rPr).Elements(W.sz).Attributes(W.val).FirstOrDefault();
                    if (defaultFontSizeString != null)
                    {
                        double dfs;
                        if (double.TryParse(defaultFontSizeString, out dfs))
                        {
                            defaultFontSize = dfs / 2d;
                            return;
                        }
                        defaultFontSize = 12;
                        return;
                    }
                }
            }
            majorLatinFont = "";
            minorLatinFont = "";
            defaultFontSize = 12;
        }

        public static string CleanUpCss(string css)
        {
            if (css == null)
                return "";
            css = css.Trim();
            string cleanCss = Regex.Split(css, "\r\n|\r|\n")
                .Where(l =>
                {
                    string lTrim = l.Trim();
                    if (lTrim == "//")
                        return false;
                    if (lTrim == "////")
                        return false;
                    if (lTrim == "<!--" || lTrim == "&lt;!--")
                        return false;
                    if (lTrim == "-->" || lTrim == "--&gt;")
                        return false;
                    return true;
                })
                .Select(l => l + Environment.NewLine )
                .StringConcatenate();
            return cleanCss;
        }
    }

    public struct Emu
    {
        public long m_Value;
        public static int s_EmusPerInch = 914400;

        public static Emu TwipsToEmus(long twips)
        {
            float v1 = (float)twips / 20f;
            float v2 = v1 / 72f;
            float v3 = v2 * s_EmusPerInch;
            long emus = (long)v3;
            return new Emu(emus);
        }

        public static Emu PointsToEmus(double points)
        {
            double v1 = points / 72;
            double v2 = v1 * s_EmusPerInch;
            long emus = (long)v2;
            return new Emu(emus);
        }

        public Emu(long value)
        {
            m_Value = value;
        }

        public static implicit operator long(Emu e)
        {
            return e.m_Value;
        }

        public static implicit operator Emu(long l)
        {
            return new Emu(l);
        }

        public override string ToString()
        {
            throw new OpenXmlPowerToolsException("Can't convert directly to string, must cast to long");
        }
    }

    public struct TPoint
    {
        public double m_Value;

        public TPoint(double value)
        {
            m_Value = value;
        }

        public static implicit operator double(TPoint t)
        {
            return t.m_Value;
        }

        public static implicit operator TPoint(double d)
        {
            return new TPoint(d);
        }

        public override string ToString()
        {
            throw new OpenXmlPowerToolsException("Can't convert directly to string, must cast to double");
        }
    }

    public struct Twip
    {
        public long m_Value;

        public Twip(long value)
        {
            m_Value = value;
        }

        public static implicit operator long(Twip t)
        {
            return t.m_Value;
        }

        public static implicit operator Twip(long l)
        {
            return new Twip(l);
        }

        public static implicit operator Twip(double d)
        {
            return new Twip((long)d);
        }

        public override string ToString()
        {
            throw new OpenXmlPowerToolsException("Can't convert directly to string, must cast to long");
        }
    }

    public class SizeEmu
    {
        public Emu m_Height;
        public Emu m_Width;

        public SizeEmu(Emu width, Emu height)
        {
            m_Width = width;
            m_Height = height;
        }

        public SizeEmu(long width, long height)
        {
            m_Width = width;
            m_Height = height;
        }
    }
}

