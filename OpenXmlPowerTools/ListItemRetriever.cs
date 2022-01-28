using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Codeuctivity
{
    public class ListItemRetrieverSettings
    {
        public static Dictionary<string, Func<int, string, string>> DefaultListItemTextImplementations =
            new Dictionary<string, Func<int, string, string>>()
            {
                {"fr-FR", ListItemTextGetter_fr_FR.GetListItemText},
                {"tr-TR", ListItemTextGetter_tr_TR.GetListItemText},
                {"ru-RU", ListItemTextGetter_ru_RU.GetListItemText},
                {"sv-SE", ListItemTextGetter_sv_SE.GetListItemText},
                {"zh-CN", ListItemTextGetter_zh_CN.GetListItemText},
            };

        public Dictionary<string, Func<int, string, string>> ListItemTextImplementations;

        public ListItemRetrieverSettings()
        {
            ListItemTextImplementations = DefaultListItemTextImplementations;
        }
    }

    public class ListItemRetriever
    {
        public class ListItemSourceSet
        {
            /// <summary>
            /// numId from the paragraph or style
            /// </summary>
            public int NumId { get; set; }

            /// <summary>
            /// num element from the numbering part
            /// </summary>
            public XElement Num { get; set; }

            /// <summary>
            /// abstract numId
            /// </summary>
            public int AbstractNumId { get; set; }

            /// <summary>
            /// abstractNum element
            /// </summary>
            public XElement AbstractNum { get; set; }

            public ListItemSourceSet(XDocument numXDoc, int numId)
            {
                NumId = numId;

                Num = numXDoc
                    .Root
                    .Elements(W.num)
                    .FirstOrDefault(n => (int)n.Attribute(W.numId) == numId);

                AbstractNumId = (int)Num
                    .Elements(W.abstractNumId)
                    .Attributes(W.val)
                    .FirstOrDefault();

                AbstractNum = numXDoc
                    .Root
                    .Elements(W.abstractNum)
.FirstOrDefault(e => (int)e.Attribute(W.abstractNumId) == AbstractNumId);
            }

            public int? StartOverride(int ilvl)
            {
                var lvlOverride = Num
                    .Elements(W.lvlOverride)
                    .FirstOrDefault(nlo => (int)nlo.Attribute(W.ilvl) == ilvl);
                if (lvlOverride != null)
                {
                    return (int?)lvlOverride
                        .Elements(W.startOverride)
                        .Attributes(W.val)
                        .FirstOrDefault();
                }

                return null;
            }

            public XElement OverrideLvl(int ilvl)
            {
                var lvlOverride = Num
                    .Elements(W.lvlOverride)
                    .FirstOrDefault(nlo => (int)nlo.Attribute(W.ilvl) == ilvl);
                if (lvlOverride != null)
                {
                    return lvlOverride.Element(W.lvl);
                }

                return null;
            }

            public XElement AbstractLvl(int ilvl)
            {
                return AbstractNum
                    .Elements(W.lvl)
                    .FirstOrDefault(al => (int)al.Attribute(W.ilvl) == ilvl);
            }

            public XElement Lvl(int ilvl)
            {
                var overrideLvl = OverrideLvl(ilvl);
                if (overrideLvl != null)
                {
                    return overrideLvl;
                }

                return AbstractLvl(ilvl);
            }
        }

        public class ListItemSource
        {
            public ListItemSourceSet Main { get; set; }
            public string NumStyleLinkName { get; set; }
            public ListItemSourceSet NumStyleLink { get; set; }
            public int Style_ilvl { get; set; }

            // for list item sources that use numStyleLink, there are two abstractId values.
            // The abstractId that is use is in num->abstractNum->numStyleLink->style->num->abstractNum

            public ListItemSource(XDocument numXDoc, XDocument stylesXDoc, int numId)
            {
                Main = new ListItemSourceSet(numXDoc, numId);

                NumStyleLinkName = (string)Main
                    .AbstractNum
                    .Elements(W.numStyleLink)
                    .Attributes(W.val)
                    .FirstOrDefault();

                if (NumStyleLinkName != null)
                {
                    var numStyleLinkNumId = (int?)stylesXDoc
                        .Root
                        .Elements(W.style)
                        .Where(s => (string)s.Attribute(W.styleId) == NumStyleLinkName)
                        .Elements(W.pPr)
                        .Elements(W.numPr)
                        .Elements(W.numId)
                        .Attributes(W.val)
                        .FirstOrDefault();

                    if (numStyleLinkNumId != null)
                    {
                        NumStyleLink = new ListItemSourceSet(numXDoc, (int)numStyleLinkNumId);
                    }
                }
            }

            public XElement Lvl(int ilvl)
            {
                var lvl2 = Main.Lvl(ilvl);
                if (lvl2 == null)
                {
                    for (var i = ilvl - 1; i >= 0; i--)
                    {
                        lvl2 = Main.Lvl(i);
                        if (lvl2 != null)
                            break;
                    }
                }
                if (lvl2 != null)
                    return lvl2;
                if (NumStyleLink != null)
                {
                    var lvl = NumStyleLink.Lvl(ilvl);
                    if (lvl == null)
                    {
                        for (var i = ilvl - 1; i >= 0; i--)
                        {
                            lvl = NumStyleLink.Lvl(i);
                            if (lvl != null)
                            {
                                break;
                            }
                        }
                    }
                    return lvl;
                }
                return null;
            }

            public int? StartOverride(int ilvl)
            {
                if (NumStyleLink != null)
                {
                    var startOverride = NumStyleLink.StartOverride(ilvl);
                    if (startOverride != null)
                    {
                        return startOverride;
                    }
                }
                return Main.StartOverride(ilvl);
            }

            public int Start(int ilvl)
            {
                var lvl = Lvl(ilvl);
                var start = (int?)lvl.Elements(W.start).Attributes(W.val).FirstOrDefault();
                if (start != null)
                {
                    return (int)start;
                }

                return 0;
            }

            public int AbstractNumId => Main.AbstractNumId;
        }

        public class ListItemInfo
        {
            public bool IsListItem { get; set; }
            public bool IsZeroNumId { get; set; }

            public ListItemSource FromStyle { get; set; }
            public ListItemSource FromParagraph { get; set; }

            private int? mAbstractNumId { get; set; } = null;

            public int? AbstractNumId
            {
                get
                {
                    // note: this property does not get NumStyleLinkAbstractNumId
                    // it presumes that we are only interested in AbstractNumId
                    // however, it is easy enough to change if necessary

                    if (mAbstractNumId != null)
                    {
                        return mAbstractNumId;
                    }

                    if (FromParagraph != null)
                    {
                        mAbstractNumId = FromParagraph.AbstractNumId;
                    }
                    else if (FromStyle != null)
                    {
                        mAbstractNumId = FromStyle.AbstractNumId;
                    }

                    return mAbstractNumId;
                }
            }

            public XElement Lvl(int ilvl)
            {
                if (FromParagraph != null)
                {
                    var lvl = FromParagraph.Lvl(ilvl);
                    if (lvl == null)
                    {
                        for (var i = ilvl - 1; i >= 0; i--)
                        {
                            lvl = FromParagraph.Lvl(i);
                            if (lvl != null)
                            {
                                break;
                            }
                        }
                    }
                    return lvl;
                }
                var lvl2 = FromStyle.Lvl(ilvl);
                if (lvl2 == null)
                {
                    for (var i = ilvl - 1; i >= 0; i--)
                    {
                        lvl2 = FromParagraph.Lvl(i);
                        if (lvl2 != null)
                        {
                            break;
                        }
                    }
                }
                return lvl2;
            }

            public int Start(int ilvl)
            {
                if (FromParagraph != null)
                {
                    return FromParagraph.Start(ilvl);
                }

                return FromStyle.Start(ilvl);
            }

            public int Start(int ilvl, bool takeOverride, out bool isOverride)
            {
                if (FromParagraph != null)
                {
                    if (takeOverride)
                    {
                        var startOverride = FromParagraph.StartOverride(ilvl);
                        if (startOverride != null)
                        {
                            isOverride = true;
                            return (int)startOverride;
                        }
                    }
                    isOverride = false;
                    return FromParagraph.Start(ilvl);
                }
                else if (FromStyle != null)
                {
                    if (takeOverride)
                    {
                        var startOverride = FromStyle.StartOverride(ilvl);
                        if (startOverride != null)
                        {
                            isOverride = true;
                            return (int)startOverride;
                        }
                    }
                    isOverride = false;
                    return FromStyle.Start(ilvl);
                }
                isOverride = false;
                return 0;
            }

            public int? StartOverride(int ilvl)
            {
                if (FromParagraph != null)
                {
                    var startOverride = FromParagraph.StartOverride(ilvl);
                    if (startOverride != null)
                    {
                        return (int)startOverride;
                    }

                    return null;
                }
                else if (FromStyle != null)
                {
                    var startOverride = FromStyle.StartOverride(ilvl);
                    if (startOverride != null)
                    {
                        return (int)startOverride;
                    }

                    return null;
                }
                return null;
            }

            private int? mNumId;

            public int NumId
            {
                get
                {
                    if (mNumId != null)
                    {
                        return (int)mNumId;
                    }

                    if (FromParagraph != null)
                    {
                        mNumId = FromParagraph.Main.NumId;
                    }
                    else if (FromStyle != null)
                    {
                        mNumId = FromStyle.Main.NumId;
                    }

                    return (int)mNumId;
                }
            }

            public ListItemInfo()
            {
            }

            public ListItemInfo(bool isListItem, bool isZeroNumId)
            {
                IsListItem = isListItem;
                IsZeroNumId = isZeroNumId;
            }
        }

        public static void SetParagraphLevel(XElement paragraph, int ilvl)
        {
            var pi = paragraph.Annotation<ParagraphInfo>();
            if (pi == null)
            {
                pi = new ParagraphInfo()
                {
                    Ilvl = ilvl,
                };
                paragraph.AddAnnotation(pi);
                return;
            }
            throw new OpenXmlPowerToolsException("Internal error - should never set ilvl more than once.");
        }

        public static int GetParagraphLevel(XElement paragraph)
        {
            var pi = paragraph.Annotation<ParagraphInfo>();
            if (pi != null)
            {
                return pi.Ilvl;
            }

            throw new OpenXmlPowerToolsException("Internal error - should never ask for ilvl without it first being set.");
        }

        public static ListItemInfo GetListItemInfo(XElement paragraph)
        {
            // The following is an optimization - only determine ListItemInfo once for a
            // paragraph.
            var listItemInfo = paragraph.Annotation<ListItemInfo>();
            if (listItemInfo != null)
            {
                return listItemInfo;
            }

            throw new OpenXmlPowerToolsException("Attempting to retrieve ListItemInfo before initialization");
        }

        private static readonly ListItemInfo NotAListItem = new ListItemInfo(false, true);

        public static void InitListItemInfo(XDocument numXDoc, XDocument stylesXDoc, XElement paragraph)
        {
            if (FirstRunIsEmptySectionBreak(paragraph))
            {
                paragraph.AddAnnotation(NotAListItem);
                return;
            }

            int? paragraphNumId = null;

            var paragraphNumberingProperties = paragraph
                .Elements(W.pPr)
                .Elements(W.numPr)
                .FirstOrDefault();

            if (paragraphNumberingProperties != null)
            {
                paragraphNumId = (int?)paragraphNumberingProperties
                    .Elements(W.numId)
                    .Attributes(W.val)
                    .FirstOrDefault();

                // if numPr of paragraph does not contain numId, then it is not a list item.
                // if numId of paragraph == 0, then this is not a list item, regardless of the markup in the style.
                if (paragraphNumId == null || paragraphNumId == 0)
                {
                    paragraph.AddAnnotation(NotAListItem);
                    return;
                }
            }

            var paragraphStyleName = GetParagraphStyleName(stylesXDoc, paragraph);

            var listItemInfo = GetListItemInfoFromCache(numXDoc, paragraphStyleName, paragraphNumId);
            if (listItemInfo != null)
            {
                paragraph.AddAnnotation(listItemInfo);

                if (listItemInfo.FromParagraph != null)
                {
                    var para_ilvl = (int?)paragraphNumberingProperties
                        .Elements(W.ilvl)
                        .Attributes(W.val)
                        .FirstOrDefault();

                    if (para_ilvl == null)
                    {
                        para_ilvl = 0;
                    }

                    var abstractNum = listItemInfo.FromParagraph.Main.AbstractNum;
                    var multiLevelType = (string)abstractNum.Elements(W.multiLevelType).Attributes(W.val).FirstOrDefault();
                    if (multiLevelType == "singleLevel")
                    {
                        para_ilvl = 0;
                    }

                    SetParagraphLevel(paragraph, (int)para_ilvl);
                }
                else if (listItemInfo.FromStyle != null)
                {
                    var this_ilvl = listItemInfo.FromStyle.Style_ilvl;
                    var abstractNum = listItemInfo.FromStyle.Main.AbstractNum;
                    var multiLevelType = (string)abstractNum.Elements(W.multiLevelType).Attributes(W.val).FirstOrDefault();
                    if (multiLevelType == "singleLevel")
                    {
                        this_ilvl = 0;
                    }

                    SetParagraphLevel(paragraph, this_ilvl);
                }
                return;
            }

            listItemInfo = new ListItemInfo();

            int? style_ilvl = null;
            bool? styleZeroNumId = null;

            if (paragraphStyleName != null)
            {
                listItemInfo.FromStyle = InitializeStyleListItemSource(numXDoc, stylesXDoc, paragraph, out style_ilvl, out styleZeroNumId);
            }

            int? paragraph_ilvl = null;
            bool? paragraphZeroNumId = null;

            if (paragraphNumberingProperties != null && paragraphNumberingProperties.Element(W.numId) != null)
            {
                listItemInfo.FromParagraph = InitializeParagraphListItemSource(numXDoc, stylesXDoc, paragraph, paragraphNumberingProperties, out paragraph_ilvl, out paragraphZeroNumId);
            }

            if (styleZeroNumId == true && paragraphZeroNumId == null || paragraphZeroNumId == true)
            {
                paragraph.AddAnnotation(NotAListItem);
                AddListItemInfoIntoCache(numXDoc, paragraphStyleName, paragraphNumId, NotAListItem);
                return;
            }

            var ilvlToSet = 0;
            if (paragraph_ilvl != null)
            {
                ilvlToSet = (int)paragraph_ilvl;
            }
            else if (style_ilvl != null)
            {
                ilvlToSet = (int)style_ilvl;
            }

            if (listItemInfo.FromParagraph != null)
            {
                var abstractNum = listItemInfo.FromParagraph.Main.AbstractNum;
                var multiLevelType = (string)abstractNum.Elements(W.multiLevelType).Attributes(W.val).FirstOrDefault();
                if (multiLevelType == "singleLevel")
                {
                    ilvlToSet = 0;
                }
            }
            else if (listItemInfo.FromStyle != null)
            {
                var abstractNum = listItemInfo.FromStyle.Main.AbstractNum;
                var multiLevelType = (string)abstractNum.Elements(W.multiLevelType).Attributes(W.val).FirstOrDefault();
                if (multiLevelType == "singleLevel")
                {
                    ilvlToSet = 0;
                }
            }

            SetParagraphLevel(paragraph, ilvlToSet);

            listItemInfo.IsListItem = listItemInfo.FromStyle != null || listItemInfo.FromParagraph != null;
            paragraph.AddAnnotation(listItemInfo);
            AddListItemInfoIntoCache(numXDoc, paragraphStyleName, paragraphNumId, listItemInfo);
        }

        private static string GetParagraphStyleName(XDocument stylesXDoc, XElement paragraph)
        {
            var paragraphStyleName = (string)paragraph
                 .Elements(W.pPr)
                 .Elements(W.pStyle)
                 .Attributes(W.val)
                 .FirstOrDefault();

            if (paragraphStyleName == null)
            {
                paragraphStyleName = GetDefaultParagraphStyleName(stylesXDoc);
            }

            return paragraphStyleName;
        }

        private static bool FirstRunIsEmptySectionBreak(XElement paragraph)
        {
            var firstRun = paragraph
                .DescendantsTrimmed(W.txbxContent)
.FirstOrDefault(d => d.Name == W.r);

            var hasTextElement = paragraph
                .DescendantsTrimmed(W.txbxContent)
                .Where(d => d.Name == W.r)
                .Elements(W.t)
                .Any();

            if (firstRun == null || !hasTextElement)
            {
                if (paragraph
                    .Elements(W.pPr)
                    .Elements(W.sectPr)
                    .Any())
                {
                    return true;
                }
            }
            return false;
        }

        private static ListItemSource InitializeParagraphListItemSource(XDocument numXDoc, XDocument stylesXDoc, XElement paragraph, XElement paragraphNumberingProperties, out int? ilvl, out bool? zeroNumId)
        {
            zeroNumId = null;

            // Paragraph numbering properties must contain a numId.
            var numId = (int?)paragraphNumberingProperties
                .Elements(W.numId)
                .Attributes(W.val)
                .FirstOrDefault();

            ilvl = (int?)paragraphNumberingProperties
                .Elements(W.ilvl)
                .Attributes(W.val)
                .FirstOrDefault();

            if (numId == null)
            {
                zeroNumId = true;
                return null;
            }

            var num = numXDoc
                .Root
                .Elements(W.num)
                .FirstOrDefault(n => (int)n.Attribute(W.numId) == numId);
            if (num == null)
            {
                zeroNumId = true;
                return null;
            }

            zeroNumId = false;

            if (ilvl == null)
            {
                ilvl = 0;
            }

            var listItemSource = new ListItemSource(numXDoc, stylesXDoc, (int)numId);

            return listItemSource;
        }

        private static ListItemSource InitializeStyleListItemSource(XDocument numXDoc, XDocument stylesXDoc, XElement paragraph, out int? ilvl, out bool? zeroNumId)
        {
            zeroNumId = null;
            var pPr = FormattingAssembler.ParagraphStyleRollup(paragraph, stylesXDoc, GetDefaultParagraphStyleName(stylesXDoc));
            if (pPr != null)
            {
                var styleNumberingProperties = pPr
                    .Elements(W.numPr)
                    .FirstOrDefault();

                if (styleNumberingProperties != null && styleNumberingProperties.Element(W.numId) != null)
                {
                    var numId = (int)styleNumberingProperties
                        .Elements(W.numId)
                        .Attributes(W.val)
                        .FirstOrDefault();

                    ilvl = (int?)styleNumberingProperties
                        .Elements(W.ilvl)
                        .Attributes(W.val)
                        .FirstOrDefault();

                    if (ilvl == null)
                    {
                        ilvl = 0;
                    }

                    if (numId == 0)
                    {
                        zeroNumId = true;
                        return null;
                    }

                    // make sure that the numId is valid
                    var num = numXDoc
                        .Root
                        .Elements(W.num)
.FirstOrDefault(e => (int)e.Attribute(W.numId) == numId);

                    if (num == null)
                    {
                        zeroNumId = true;
                        return null;
                    }

                    var listItemSource = new ListItemSource(numXDoc, stylesXDoc, numId)
                    {
                        Style_ilvl = (int)ilvl
                    };

                    zeroNumId = false;
                    return listItemSource;
                }
            }
            ilvl = null;
            return null;
        }

        private static string GetDefaultParagraphStyleName(XDocument stylesXDoc)
        {
            XElement defaultParagraphStyle;
            string defaultParagraphStyleName = null;

            var stylesInfo = stylesXDoc.Annotation<StylesInfo>();

            if (stylesInfo != null)
            {
                defaultParagraphStyleName = stylesInfo.DefaultParagraphStyleName;
            }
            else
            {
                defaultParagraphStyle = stylesXDoc
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(s =>
                    {
                        if ((string)s.Attribute(W.type) != "paragraph")
                        {
                            return false;
                        }

                        var defaultAttribute = s.Attribute(W._default);
                        var isDefault = false;
                        if (defaultAttribute != null &&
                            (bool)s.Attribute(W._default).ToBoolean())
                        {
                            isDefault = true;
                        }

                        return isDefault;
                    });
                defaultParagraphStyleName = null;
                if (defaultParagraphStyle != null)
                {
                    defaultParagraphStyleName = (string)defaultParagraphStyle.Attribute(W.styleId);
                }

                stylesInfo = new StylesInfo()
                {
                    DefaultParagraphStyleName = defaultParagraphStyleName,
                };
                stylesXDoc.AddAnnotation(stylesInfo);
            }
            return defaultParagraphStyleName;
        }

        private static ListItemInfo GetListItemInfoFromCache(XDocument numXDoc, string styleName, int? numId)
        {
            var key = (styleName ?? "") + "|" + (numId == null ? "" : numId.ToString());

            var numXDocRoot = numXDoc.Root;
            var listItemInfoCache =
                numXDocRoot.Annotation<Dictionary<string, ListItemInfo>>();
            if (listItemInfoCache == null)
            {
                listItemInfoCache = new Dictionary<string, ListItemInfo>();
                numXDocRoot.AddAnnotation(listItemInfoCache);
            }
            if (listItemInfoCache.ContainsKey(key))
            {
                return listItemInfoCache[key];
            }

            return null;
        }

        private static void AddListItemInfoIntoCache(XDocument numXDoc, string styleName, int? numId, ListItemInfo listItemInfo)
        {
            var key =
                (styleName == null ? "" : styleName) +
                "|" +
                (numId == null ? "" : numId.ToString());

            var numXDocRoot = numXDoc.Root;
            var listItemInfoCache =
                numXDocRoot.Annotation<Dictionary<string, ListItemInfo>>();
            if (listItemInfoCache == null)
            {
                listItemInfoCache = new Dictionary<string, ListItemInfo>();
                numXDocRoot.AddAnnotation(listItemInfoCache);
            }
            if (!listItemInfoCache.ContainsKey(key))
            {
                listItemInfoCache.Add(key, listItemInfo);
            }
        }

        public class LevelNumbers
        {
            public int[] LevelNumbersArray { get; set; }
        }

        private class StylesInfo
        {
            public string DefaultParagraphStyleName { get; set; }
        }

        private class ParagraphInfo
        {
            public int Ilvl { get; set; }
        }

        private class ReverseAxis
        {
            public XElement PreviousParagraph { get; set; }
        }

        public static string RetrieveListItem(WordprocessingDocument wordDoc, XElement paragraph)
        {
            return RetrieveListItem(wordDoc, paragraph, null);
        }

        public static string RetrieveListItem(WordprocessingDocument wordDoc, XElement paragraph, ListItemRetrieverSettings settings)
        {
            if (wordDoc.MainDocumentPart.NumberingDefinitionsPart == null)
            {
                return null;
            }

            var listItemInfo = paragraph.Annotation<ListItemInfo>();
            if (listItemInfo == null)
            {
                InitializeListItemRetriever(wordDoc);
            }

            listItemInfo = paragraph.Annotation<ListItemInfo>();
            if (!listItemInfo.IsListItem)
            {
                return null;
            }

            var numberingDefinitionsPart = wordDoc
                .MainDocumentPart
                .NumberingDefinitionsPart;

            if (numberingDefinitionsPart == null)
            {
                return null;
            }

            var styleDefinitionsPart = wordDoc
                .MainDocumentPart
                .StyleDefinitionsPart;

            if (styleDefinitionsPart == null)
            {
                return null;
            }

            var stylesXDoc = styleDefinitionsPart.GetXDocument();

            var paragraphLevel = GetParagraphLevel(paragraph);
            var lvl = listItemInfo.Lvl(paragraphLevel);

            var lvlText = (string)lvl.Elements(W.lvlText).Attributes(W.val).FirstOrDefault();
            if (lvlText == null)
            {
                return null;
            }

            var levelNumbersAnnotation = paragraph.Annotation<LevelNumbers>();
            if (levelNumbersAnnotation == null)
            {
                throw new OpenXmlPowerToolsException("Internal error");
            }

            var levelNumbers = levelNumbersAnnotation.LevelNumbersArray;
            var languageIdentifier = GetLanguageIdentifier(paragraph, stylesXDoc);
            var listItem = FormatListItem(listItemInfo, levelNumbers, GetParagraphLevel(paragraph),
                lvlText, languageIdentifier, settings);
            return listItem;
        }

        private static string GetLanguageIdentifier(XElement paragraph, XDocument stylesXDoc)
        {
            var languageType = (string)paragraph
                .DescendantsTrimmed(W.txbxContent)
                .Where(d => d.Name == W.r)
                .Attributes(PtOpenXml.LanguageType)
                .FirstOrDefault();

            string languageIdentifier = null;

            if (languageType == null || languageType == "western")
            {
                languageIdentifier = (string)paragraph
                    .Elements(W.r)
                    .Elements(W.rPr)
                    .Elements(W.lang)
                    .Attributes(W.val)
                    .FirstOrDefault();

                if (languageIdentifier == null)
                {
                    languageIdentifier = (string)stylesXDoc
                        .Root
                        .Elements(W.docDefaults)
                        .Elements(W.rPrDefault)
                        .Elements(W.rPr)
                        .Elements(W.lang)
                        .Attributes(W.val)
                        .FirstOrDefault();
                }
            }
            else if (languageType == "eastAsia")
            {
                languageIdentifier = (string)paragraph
                    .Elements(W.r)
                    .Elements(W.rPr)
                    .Elements(W.lang)
                    .Attributes(W.eastAsia)
                    .FirstOrDefault();

                if (languageIdentifier == null)
                {
                    languageIdentifier = (string)stylesXDoc
                        .Root
                        .Elements(W.docDefaults)
                        .Elements(W.rPrDefault)
                        .Elements(W.rPr)
                        .Elements(W.lang)
                        .Attributes(W.eastAsia)
                        .FirstOrDefault();
                }
            }
            else if (languageType == "bidi")
            {
                languageIdentifier = (string)paragraph
                    .Elements(W.r)
                    .Elements(W.rPr)
                    .Elements(W.lang)
                    .Attributes(W.bidi)
                    .FirstOrDefault();

                if (languageIdentifier == null)
                {
                    languageIdentifier = (string)stylesXDoc
                        .Root
                        .Elements(W.docDefaults)
                        .Elements(W.rPrDefault)
                        .Elements(W.rPr)
                        .Elements(W.lang)
                        .Attributes(W.bidi)
                        .FirstOrDefault();
                }
            }

            if (languageIdentifier == null)
            {
                languageIdentifier = "en-US";
            }

            return languageIdentifier;
        }

        private static void InitializeListItemRetriever(WordprocessingDocument wordDoc)
        {
            foreach (var part in wordDoc.ContentParts())
            {
                InitializeListItemRetrieverForPart(wordDoc, part);
            }
        }

        private static void InitializeListItemRetrieverForPart(WordprocessingDocument wordDoc, OpenXmlPart part)
        {
            var mainXDoc = part.GetXDocument();

            var numPart = wordDoc.MainDocumentPart.NumberingDefinitionsPart;
            if (numPart == null)
            {
                return;
            }

            var numXDoc = numPart.GetXDocument();

            var stylesPart = wordDoc.MainDocumentPart.StyleDefinitionsPart;
            if (stylesPart == null)
            {
                return;
            }

            var stylesXDoc = stylesPart.GetXDocument();

            var rootNode = mainXDoc.Root;

            InitializeListItemRetrieverForStory(numXDoc, stylesXDoc, rootNode);

            var textBoxes = mainXDoc
                .Root
                .Descendants(W.txbxContent);

            foreach (var textBox in textBoxes)
            {
                InitializeListItemRetrieverForStory(numXDoc, stylesXDoc, textBox);
            }
        }

        private static void InitializeListItemRetrieverForStory(XDocument numXDoc, XDocument stylesXDoc, XElement rootNode)
        {
            var paragraphs = rootNode
                .DescendantsTrimmed(W.txbxContent)
                .Where(p => p.Name == W.p);

            foreach (var paragraph in paragraphs)
            {
                InitListItemInfo(numXDoc, stylesXDoc, paragraph);
            }

            var abstractNumIds = paragraphs
                .Select(paragraph =>
                {
                    var listItemInfo = paragraph.Annotation<ListItemInfo>();
                    if (!listItemInfo.IsListItem)
                    {
                        return null;
                    }

                    return listItemInfo.AbstractNumId;
                })
                .Where(a => a != null)
                .Distinct()
                .ToList();

            // when debugging, it is sometimes useful to cause processing of a specific abstractNumId first.
            // the following code enables this.
            //int? abstractIdToProcessFirst = null;
            //if (abstractIdToProcessFirst != null)
            //{
            //    abstractNumIds = (new[] { abstractIdToProcessFirst })
            //        .Concat(abstractNumIds.Where(ani => ani != abstractIdToProcessFirst))
            //        .ToList();
            //}

            foreach (var abstractNumId in abstractNumIds)
            {
                var listItems = paragraphs
                    .Where(paragraph =>
                    {
                        var listItemInfo = paragraph.Annotation<ListItemInfo>();
                        if (!listItemInfo.IsListItem)
                        {
                            return false;
                        }

                        return listItemInfo.AbstractNumId == abstractNumId;
                    })
                    .ToList();

                // annotate paragraphs with previous paragraphs so that we can look backwards with good perf
                XElement prevParagraph = null;
                foreach (var paragraph in listItems)
                {
                    var reverse = new ReverseAxis()
                    {
                        PreviousParagraph = prevParagraph,
                    };
                    paragraph.AddAnnotation(reverse);
                    prevParagraph = paragraph;
                }

                var startOverrideAlreadyUsed = new List<int>();
                List<int> previous = null;
                var listItemInfoInEffectForStartOverride = new ListItemInfo[] {
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                };
                foreach (var paragraph in listItems)
                {
                    var listItemInfo = paragraph.Annotation<ListItemInfo>();
                    var ilvl = GetParagraphLevel(paragraph);
                    listItemInfoInEffectForStartOverride[ilvl] = listItemInfo;
                    ListItemInfo listItemInfoInEffect = null;
                    if (ilvl > 0)
                    {
                        listItemInfoInEffect = listItemInfoInEffectForStartOverride[ilvl - 1];
                    }

                    var levelNumbers = new List<int>();

                    for (var level = 0; level <= ilvl; level++)
                    {
                        var numId = listItemInfo.NumId;
                        var startOverride = listItemInfo.StartOverride(level);
                        int? inEffectStartOverride = null;
                        if (listItemInfoInEffect != null)
                        {
                            inEffectStartOverride = listItemInfoInEffect.StartOverride(level);
                        }

                        if (level == ilvl)
                        {
                            var lvl = listItemInfo.Lvl(ilvl);
                            var lvlRestart = (int?)lvl.Elements(W.lvlRestart).Attributes(W.val).FirstOrDefault();
                            if (lvlRestart != null)
                            {
                                var previousPara = PreviousParagraphsForLvlRestart(paragraph, (int)lvlRestart)
                                    .FirstOrDefault(p =>
                                    {
                                        var plvl = GetParagraphLevel(p);
                                        return plvl == ilvl;
                                    });
                                if (previousPara != null)
                                {
                                    previous = previousPara.Annotation<LevelNumbers>().LevelNumbersArray.ToList();
                                }
                            }
                        }

                        if (previous == null || level >= previous.Count || level == ilvl && startOverride != null && !startOverrideAlreadyUsed.Contains(numId))
                        {
                            if (previous == null || level >= previous.Count)
                            {
                                var start = listItemInfo.Start(level);
                                // only look at startOverride if the level that we're examining is same as the paragraph's level.
                                if (level == ilvl)
                                {
                                    if (startOverride != null && !startOverrideAlreadyUsed.Contains(numId))
                                    {
                                        startOverrideAlreadyUsed.Add(numId);
                                        start = (int)startOverride;
                                    }
                                    else
                                    {
                                        if (startOverride != null)
                                        {
                                            start = (int)startOverride;
                                        }

                                        if (inEffectStartOverride != null && inEffectStartOverride > start)
                                        {
                                            start = (int)inEffectStartOverride;
                                        }
                                    }
                                }
                                levelNumbers.Add(start);
                            }
                            else
                            {
                                var start = listItemInfo.Start(level);
                                // only look at startOverride if the level that we're examining is same as the paragraph's level.
                                if (level == ilvl)
                                {
                                    if (startOverride != null)
                                    {
                                        if (!startOverrideAlreadyUsed.Contains(numId))
                                        {
                                            startOverrideAlreadyUsed.Add(numId);
                                            start = (int)startOverride;
                                        }
                                    }
                                }
                                levelNumbers.Add(start);
                            }
                        }
                        else
                        {
                            int? thisNumber = null;
                            if (level == ilvl)
                            {
                                if (startOverride != null)
                                {
                                    if (!startOverrideAlreadyUsed.Contains(numId))
                                    {
                                        startOverrideAlreadyUsed.Add(numId);
                                        thisNumber = (int)startOverride;
                                    }
                                    thisNumber = previous.ElementAt(level) + 1;
                                }
                                else
                                {
                                    thisNumber = previous.ElementAt(level) + 1;
                                }
                            }
                            else
                            {
                                thisNumber = previous.ElementAt(level);
                            }
                            levelNumbers.Add((int)thisNumber);
                        }
                    }
                    var levelNumbersAnno = new LevelNumbers()
                    {
                        LevelNumbersArray = levelNumbers.ToArray()
                    };
                    paragraph.AddAnnotation(levelNumbersAnno);
                    previous = levelNumbers;
                }
            }
        }

        private static IEnumerable<XElement> PreviousParagraphsForLvlRestart(XElement paragraph, int ilvl)
        {
            var current = paragraph;
            while (true)
            {
                var ra = current.Annotation<ReverseAxis>();
                if (ra == null || ra.PreviousParagraph == null)
                {
                    yield break;
                }

                var raLvl = GetParagraphLevel(ra.PreviousParagraph);
                if (raLvl < ilvl)
                {
                    yield break;
                }

                yield return ra.PreviousParagraph;
                current = ra.PreviousParagraph;
            }
        }

        private static string FormatListItem(ListItemInfo lii, int[] levelNumbers, int ilvl, string lvlText, string languageCultureName, ListItemRetrieverSettings settings)
        {
            var formatTokens = GetFormatTokens(lvlText).ToArray();
            var lvl = lii.Lvl(ilvl);
            var isLgl = lvl.Elements(W.isLgl).Any();
            var listItem = formatTokens.Select((t, l) =>
            {
                if (t.Substring(0, 1) != "%")
                {
                    return t;
                }

                if (!int.TryParse(t.Substring(1), out var indentationLevel))
                {
                    return t;
                }

                indentationLevel -= 1;
                if (indentationLevel >= levelNumbers.Length)
                {
                    indentationLevel = levelNumbers.Length - 1;
                }

                var levelNumber = levelNumbers[indentationLevel];
                string levelText = null;
                var rlvl = lii.Lvl(indentationLevel);
                var numFmtForLevel = (string)rlvl.Elements(W.numFmt).Attributes(W.val).FirstOrDefault();
                if (numFmtForLevel == null)
                {
                    var numFmtElement = rlvl.Elements(MC.AlternateContent).Elements(MC.Choice).Elements(W.numFmt).FirstOrDefault();
                    if (numFmtElement != null && (string)numFmtElement.Attribute(W.val) == "custom")
                    {
                        numFmtForLevel = (string)numFmtElement.Attribute(W.format);
                    }
                }
                if (numFmtForLevel != "none")
                {
                    if (isLgl && numFmtForLevel != "decimalZero")
                    {
                        numFmtForLevel = "decimal";
                    }
                }
                if (languageCultureName != null && settings != null)
                {
                    if (settings.ListItemTextImplementations.ContainsKey(languageCultureName))
                    {
                        var impl = settings.ListItemTextImplementations[languageCultureName];
                        levelText = impl(levelNumber, numFmtForLevel);
                    }
                }
                if (levelText == null)
                {
                    levelText = ListItemTextGetter_Default.GetListItemText(levelNumber, numFmtForLevel);
                }

                return levelText;
            }).StringConcatenate();
            return listItem;
        }

        private static IEnumerable<string> GetFormatTokens(string lvlText)
        {
            var i = 0;
            while (true)
            {
                if (i >= lvlText.Length)
                {
                    yield break;
                }

                if (lvlText[i] == '%' && i <= lvlText.Length - 2)
                {
                    yield return lvlText.Substring(i, 2);
                    i += 2;
                    continue;
                }
                var percentIndex = lvlText.IndexOf('%', i);
                if (percentIndex == -1 || percentIndex > lvlText.Length - 2)
                {
                    yield return lvlText.Substring(i);
                    yield break;
                }
                yield return lvlText.Substring(i, percentIndex - i);
                yield return lvlText.Substring(percentIndex, 2);
                i = percentIndex + 2;
            }
        }
    }
}