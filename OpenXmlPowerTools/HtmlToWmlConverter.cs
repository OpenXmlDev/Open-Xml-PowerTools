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

