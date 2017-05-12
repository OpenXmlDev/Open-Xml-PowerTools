/***************************************************************************

Copyright (c) Eric White 2016.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

Published at http://EricWhite.com
Resource Center and Documentation: http://ericwhite.com/

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

***************************************************************************/

// TODO we are not generating lang if English, but this needs revised after analysis

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using System.Drawing;

namespace OpenXmlPowerTools
{
    public class ContentTypeRule
    {
        public string ContentType;
        public string StyleName;
        public Regex StyleNameRegex;
        public Regex[] RegexArray;
        public Func<XElement, ContentTypeRule, WordprocessingDocument, WmlToXmlSettings, bool> MatchLambda;
        public bool ApplyRunContentTypes = true;
    }

    public class WmlToXmlProgressInfo
    {
        public int ContentCount;
        public int ContentTotal;
        public string InProgressMessage;
    }

    public class TransformInfo
    {
        public string DefaultLangFromStylesPart;
    }

    public class WmlToXmlSettings
    {
        public List<ContentTypeRule> GlobalContentTypeRules;
        public List<ContentTypeRule> DocumentTypeContentTypeRules;
        public List<ContentTypeRule> DocumentContentTypeRules;
        public List<ContentTypeRule> RunContentTypeRules;
        public ListItemRetrieverSettings ListItemRetrieverSettings;
        public bool? InjectCommentForContentTypes;
        public Func<XElement, WmlToXmlSettings, int?> ContentTypeHierarchyLambda;
        public Dictionary<string, Func<string, OpenXmlPart, XElement, WmlToXmlSettings, object>> XmlGenerationLambdas;
        public DirectoryInfo ImageBase;
        public bool WriteImageFiles = true;
        public Action<WmlToXmlProgressInfo> ProgressFunction;
        public XDocument ContentTypeRegexExtension;
        public string DefaultLang;
        public object UserData;

        public WmlToXmlSettings(
            List<ContentTypeRule> globalContentTypeRules,
            List<ContentTypeRule> documentTypeContentTypeRules,
            List<ContentTypeRule> documentContentTypeRules,
            List<ContentTypeRule> runContentTypeRules,
            Func<XElement, WmlToXmlSettings, int?> contentTypeHierarchyLambda,
            Dictionary<string, Func<string, OpenXmlPart, XElement, WmlToXmlSettings, object>> xmlGenerationLambdas,
            DirectoryInfo imageBase,
            XDocument contentTypeRegexExtension)
        {
            GlobalContentTypeRules = globalContentTypeRules;
            DocumentTypeContentTypeRules = documentTypeContentTypeRules;
            DocumentContentTypeRules = documentContentTypeRules;
            RunContentTypeRules = runContentTypeRules;
            ListItemRetrieverSettings = new ListItemRetrieverSettings();
            ContentTypeHierarchyLambda = contentTypeHierarchyLambda;
            XmlGenerationLambdas = xmlGenerationLambdas;
            ImageBase = imageBase;
            ContentTypeRegexExtension = contentTypeRegexExtension;
        }

        public WmlToXmlSettings(
            List<ContentTypeRule> globalContentTypeRules,
            List<ContentTypeRule> documentTypeContentTypeRules,
            List<ContentTypeRule> documentContentTypeRules,
            List<ContentTypeRule> runContentTypeRules,
            Func<XElement, WmlToXmlSettings, int?> contentTypeHierarchyLambda,
            Dictionary<string, Func<string, OpenXmlPart, XElement, WmlToXmlSettings, object>> xmlGenerationLambdas,
            ListItemRetrieverSettings listItemRetrieverSettings,
            DirectoryInfo imageBase,
            XDocument contentTypeRegexExtension)
        {
            GlobalContentTypeRules = globalContentTypeRules;
            DocumentTypeContentTypeRules = documentTypeContentTypeRules;
            DocumentContentTypeRules = documentContentTypeRules;
            RunContentTypeRules = runContentTypeRules;
            ListItemRetrieverSettings = listItemRetrieverSettings;
            ContentTypeHierarchyLambda = contentTypeHierarchyLambda;
            XmlGenerationLambdas = xmlGenerationLambdas;
            ImageBase = imageBase;
            ContentTypeRegexExtension = contentTypeRegexExtension;
        }
    }

    public static class WmlToXml
    {
        public static WmlDocument ApplyContentTypes(WmlDocument document, WmlToXmlSettings settings)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(document))
            {
                using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                {
                    ApplyContentTypes(doc, settings);
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }

        public static void ApplyContentTypes(WordprocessingDocument wDoc, WmlToXmlSettings settings)
        {
#if false
<Extensions>
  <Extension ContentType='Introduction'>
    <RegexExtension>
      <Regex>.*Infroduction.*</Regex>
      <Regex>.*Entroduction.*</Regex>
    </RegexExtension>
  </Extension>
</Extensions>
#endif
            if (settings.ContentTypeRegexExtension != null)
            {
                foreach (var ext in settings.ContentTypeRegexExtension.Root.Elements("Extension"))
                {
                    var ct = (string)ext.Attribute("ContentType");
                    var rules = settings.DocumentContentTypeRules.Concat(settings.DocumentTypeContentTypeRules).Concat(settings.GlobalContentTypeRules);
                    var ruleToUpdate = rules
                        .FirstOrDefault(r => r.ContentType == ct);
                    if (ruleToUpdate == null)
                        throw new OpenXmlPowerToolsException("ContentTypeRexexExtension refers to content type that does not exist");
                    var oldRegexRules = ruleToUpdate.RegexArray.ToList();
                    var newRegexRules = ext.Elements("RegexExtension").Elements("Regex").Select(z => new Regex(z.Value)).ToArray();
                    var regexArray = oldRegexRules.Concat(newRegexRules).ToArray();
                    ruleToUpdate.RegexArray = regexArray;
                }
            }

            if (settings.ProgressFunction != null)
            {
                WmlToXmlProgressInfo pi = new WmlToXmlProgressInfo()
                {
                    ContentCount = 0,
                    ContentTotal = 0,
                    InProgressMessage = "Simplify markup" + Environment.NewLine,
                };
                settings.ProgressFunction(pi);
            }

            SimplifyMarkupSettings markupSimplifierSettings = new SimplifyMarkupSettings()
            {
                AcceptRevisions = true,
                NormalizeXml = true,
                RemoveBookmarks = false,
                RemoveComments = true,
                RemoveContentControls = false,
                RemoveEndAndFootNotes = false,
                RemoveFieldCodes = false,
                RemoveGoBackBookmark = true,
                RemoveHyperlinks = false,
                RemoveLastRenderedPageBreak = true,
                RemoveMarkupForDocumentComparison = false,
                RemovePermissions = true,
                RemoveProof = true,
                RemoveRsidInfo = true,
                RemoveSmartTags = true,
                RemoveSoftHyphens = false,
                RemoveWebHidden = true,
                ReplaceTabsWithSpaces = false,
            };
            MarkupSimplifier.SimplifyMarkup(wDoc, markupSimplifierSettings);

            if (settings.ProgressFunction != null)
            {
                WmlToXmlProgressInfo pi = new WmlToXmlProgressInfo()
                {
                    ContentCount = 0,
                    ContentTotal = 0,
                    InProgressMessage = "Assemble formatting" + Environment.NewLine,
                };
                settings.ProgressFunction(pi);
            }

            FormattingAssemblerSettings formattingAssemblerSettings = new FormattingAssemblerSettings();
            formattingAssemblerSettings.RemoveStyleNamesFromParagraphAndRunProperties = false;
            formattingAssemblerSettings.RestrictToSupportedLanguages = false;
            formattingAssemblerSettings.RestrictToSupportedNumberingFormats = false;
            FormattingAssembler.AssembleFormatting(wDoc, formattingAssemblerSettings);

            ContentTypeApplierInfo ctai = new ContentTypeApplierInfo();

            XDocument sXDoc = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
            XElement defaultParagraphStyle = sXDoc
                .Root
                .Elements(W.style)
                .FirstOrDefault(st => st.Attribute(W._default).ToBoolean() == true &&
                    (string)st.Attribute(W.type) == "paragraph");
            if (defaultParagraphStyle != null)
                ctai.DefaultParagraphStyleName = (string)defaultParagraphStyle.Attribute(W.styleId);
            XElement defaultCharacterStyle = sXDoc
                .Root
                .Elements(W.style)
                .FirstOrDefault(st => st.Attribute(W._default).ToBoolean() == true &&
                    (string)st.Attribute(W.type) == "character");
            if (defaultCharacterStyle != null)
                ctai.DefaultCharacterStyleName = (string)defaultCharacterStyle.Attribute(W.styleId);
            XElement defaultTableStyle = sXDoc
                .Root
                .Elements(W.style)
                .FirstOrDefault(st => st.Attribute(W._default).ToBoolean() == true &&
                    (string)st.Attribute(W.type) == "table");
            if (defaultTableStyle != null)
                ctai.DefaultTableStyleName = (string)defaultTableStyle.Attribute(W.styleId);

            if (settings.ProgressFunction != null)
            {
                WmlToXmlProgressInfo pi = new WmlToXmlProgressInfo()
                {
                    ContentCount = 0,
                    ContentTotal = 0,
                    InProgressMessage = "Assemble list item information" + Environment.NewLine,
                };
                settings.ProgressFunction(pi);
            }

            ListItemRetrieverSettings listItemRetrieverSettings = new ListItemRetrieverSettings();
            AssembleListItemInformation(wDoc, settings.ListItemRetrieverSettings);







            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // TODO temporarily assign levels to headings.  This needs removed and replaced when I rework the hierarchical infrastructure.
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            TemporarilyAssignLevelsToHeadings(wDoc);








            ApplyContentTypesForRuleSet(settings, ctai, wDoc);
        }

        private static void TemporarilyAssignLevelsToHeadings(WordprocessingDocument wDoc)
        {
            XDocument xDoc = wDoc.MainDocumentPart.GetXDocument();
            var headings = xDoc.Descendants(W.p)
                .Where(p =>
                {
                    var styleId = (string)p
                        .Elements(W.pPr)
                        .Elements(W.pStyle)
                        .Attributes(W.val)
                        .FirstOrDefault();
                    if (styleId == "Heading1" ||
                        styleId == "Heading2" ||
                        styleId == "Heading3" ||
                        styleId == "Heading4" ||
                        styleId == "Heading5" ||
                        styleId == "Heading6" ||
                        styleId == "Heading9")
                        return true;
                    else
                        return false;
                });
            var seenHeading6 = false;
            foreach (var hdg in headings)
            {
                var styleId = (string)hdg
                    .Elements(W.pPr)
                    .Elements(W.pStyle)
                    .Attributes(W.val)
                    .FirstOrDefault();
                if (styleId == "Heading6")
                    seenHeading6 = true;
                int lvl = 0;
                if (seenHeading6)
                {
                    if (styleId == "Heading6")
                        lvl = 1;
                    else if (styleId == "Heading1")
                        lvl = 2;
                    else if (styleId == "Heading2")
                        lvl = 3;
                    else if (styleId == "Heading3")
                        lvl = 4;
                    else if (styleId == "Heading4")
                        lvl = 5;
                    else if (styleId == "Heading5")
                        lvl = 6;
                    else if (styleId == "Heading9")
                        lvl = 3;
                }
                else
                {
                    if (styleId == "Heading1")
                        lvl = 1;
                    else if (styleId == "Heading2")
                        lvl = 2;
                    else if (styleId == "Heading3")
                        lvl = 3;
                    else if (styleId == "Heading4")
                        lvl = 4;
                    else if (styleId == "Heading5")
                        lvl = 5;
                    else if (styleId == "Heading9")
                        lvl = 2;
                }
                hdg.Add(new XAttribute(PtOpenXml.Level, lvl));
            }
        }

        public static XElement ProduceContentTypeXml(WmlDocument document, WmlToXmlSettings settings)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(document))
            {
                using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                {
                    return ProduceContentTypeXml(doc, settings);
                }
            }
        }

        public static XElement ProduceContentTypeXml(WordprocessingDocument wDoc, WmlToXmlSettings settings)
        {
            var mainPart = wDoc.MainDocumentPart;
            var mainXDoc = mainPart.GetXDocument();

#if false
<w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" mc:Ignorable="w14 w15 w16se">
	<w:docDefaults>
		<w:rPrDefault>
			<w:rPr>
				<w:rFonts w:ascii="Georgia" w:eastAsiaTheme="minorHAnsi" w:hAnsi="Georgia" w:cs="Times New Roman"/>
				<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>
			</w:rPr>
		</w:rPrDefault>
		<w:pPrDefault/>
	</w:docDefaults>
#endif

            AssignLevelsToContent(mainXDoc, settings);

            // Call RetrieveListItem so that all paragraphs are initialized with ListItemInfo
            var firstParagraph = mainXDoc.Descendants(W.p).FirstOrDefault();

            // if there is no content, then return an empty document.
            if (firstParagraph == null)
                return new XElement("ContentTypeXml");

            var listItem = ListItemRetriever.RetrieveListItem(wDoc, firstParagraph);

            // Annotate runs associated with fields, so that can retrieve hyperlinks that are stored as fields.
            FieldRetriever.AnnotateWithFieldInfo(wDoc.MainDocumentPart);

            var body = mainXDoc.Root.Descendants(W.body).FirstOrDefault();
            if (body == null)
                throw new OpenXmlPowerToolsException("Internal error: invalid document");

            AnnotateRunsThatUseFieldsForNumbering(mainXDoc);

            var contentList = body.Elements()
                .Where(e => e.Attribute(PtOpenXml.Level) != null)
                .ToList();

            var rootLevelContentList = contentList
                .Where(h => (int)h.Attribute(PtOpenXml.Level) == 1)
                .ToList();

            var contentTypeXml = new XElement("ContentTypeXml",
                rootLevelContentList
                    .Select(h =>
                    {
                        var childrenHeadings = GetChildrenHeadings(mainPart, contentList, h, settings);
                        XElement xml = (XElement)ProduceXmlTransform(mainPart, h, settings);
                        if (xml != null)
                            xml.Add(childrenHeadings);
                        return xml;
                    }));
            return contentTypeXml;
        }

        private static void AnnotateRunsThatUseFieldsForNumbering(XDocument mainXDoc)
        {
            var cachedAnnotationInformation = mainXDoc.Root.Annotation<Dictionary<int, List<XElement>>>();
            if (cachedAnnotationInformation == null)
                return;

            StringBuilder sb = new StringBuilder();
            foreach (var item in cachedAnnotationInformation)
            {
                var instrText = FieldRetriever.InstrText(mainXDoc.Root, item.Key).TrimStart('{').TrimEnd('}');
                var fi = FieldRetriever.ParseField(instrText);
                if (fi.FieldType.ToUpper() == "SEQ")
                {
                    // have it

                    var runsForField = mainXDoc
                        .Root
                        .Descendants()
                        .Where(d =>
                        {
                            Stack<FieldRetriever.FieldElementTypeInfo> stack = d.Annotation<Stack<FieldRetriever.FieldElementTypeInfo>>();
                            if (stack == null)
                                return false;
                            if (stack.Any(stackItem => stackItem.Id == item.Key && stackItem.FieldElementType == FieldRetriever.FieldElementTypeEnum.Result))
                                return true;
                            return false;
                        })
                        .Select(d => d.AncestorsAndSelf(W.r).FirstOrDefault())
                        .Where(z9 => z9 != null)
                        .GroupAdjacent(o => o)
                        .Select(g => g.First())
                        .Where(r => r.Element(W.t) != null)
                        .ToList();

                    if (!runsForField.Any())
                        continue;

                    var lastRun = runsForField
                        .Last();

                    var lastRunTextElement = lastRun
                        .Element(W.t);

                    var lastRunText = lastRunTextElement.Value;
                    
                    var nextRun = lastRun
                        .ElementsAfterSelf(W.r)
                        .FirstOrDefault(r => r.Element(W.t) != null);

                    if (nextRun != null)
                    {
                        var nextRunTextElement = nextRun
                            .Element(W.t);

                        var nextRunText = nextRunTextElement.Value;
                        var sepChars = nextRunText
                            .TakeWhile(ch => ch == '.' || ch == ' ')
                            .ToList();

                        nextRunText = nextRunText.Substring(sepChars.Count());
                        nextRunTextElement.Value = nextRunText;

                        lastRunText = lastRunTextElement.Value + sepChars.Select(ch => ch.ToString()).StringConcatenate();
                        lastRunTextElement.Value = lastRunText;
                    }

                    lastRun.Add(new XAttribute(PtOpenXml.ListItemRun, lastRunText));

                    foreach (var runbefore in lastRun
                        .ElementsBeforeSelf(W.r)
                        .Where(rz => rz.Element(W.t) != null))
                    {
                        runbefore.Add(new XAttribute(PtOpenXml.ListItemRun, lastRunText));
                    }
                }
            }
        }

        // this method produces the XML for an endnote or footnote - the blockLevelContentContainer is the w:endnote or w:footnote element, and it produces the content type XML for the
        // contents of the endnote or footnote, to be inserted en situ in the ContentTypeXml.
        public static object ProduceContentTypeXmlForBlockLevelContentContainer(WordprocessingDocument wDoc, WmlToXmlSettings settings, OpenXmlPart part, XElement blockLevelContentContainer)
        {
            AssignLevelsToContent(blockLevelContentContainer, settings);

            // Call RetrieveListItem so that all paragraphs are initialized with ListItemInfo
            var firstParagraph = blockLevelContentContainer.Descendants(W.p).FirstOrDefault();
            var listItem = ListItemRetriever.RetrieveListItem(wDoc, firstParagraph);

            var contentList = blockLevelContentContainer.Elements()
                .Where(e => e.Attribute(PtOpenXml.Level) != null)
                .ToList();

            var rootLevelContentList = contentList
                .Where(h => (int)h.Attribute(PtOpenXml.Level) == 1)
                .ToList();

            var contentTypeXml = rootLevelContentList
                    .Select(h =>
                    {
                        var childrenHeadings = GetChildrenHeadings(part, contentList, h, settings);
                        XElement xml = (XElement)ProduceXmlTransform(part, h, settings);
                        if (xml != null)
                            xml.Add(childrenHeadings);
                        return xml;
                    });
            return contentTypeXml;
        }


        private static object GetChildrenHeadings(OpenXmlPart part, List<XElement> contentList, XElement parent, WmlToXmlSettings settings)
        {
            return contentList
                    .SkipWhile(h => h != parent)
                    .Skip(1)
                    .TakeWhile(h => (int)h.Attribute(PtOpenXml.Level) > (int)parent.Attribute(PtOpenXml.Level))
                    .Where(h => (int)h.Attribute(PtOpenXml.Level) == (int)parent.Attribute(PtOpenXml.Level) + 1)
                    .Select(h =>
                    {
                        var childrenHeadings = GetChildrenHeadings(part, contentList, h, settings);
                        XElement xml = (XElement)ProduceXmlTransform(part, h, settings);
                        if (xml != null)
                            xml.Add(childrenHeadings);
                        return xml;
                    }
                    );
        }

        public static object ProduceXmlTransform(OpenXmlPart part, XNode node, WmlToXmlSettings settings)
        {
            var element = node as XElement;
            if (element != null)
            {
                if (settings.XmlGenerationLambdas == null)
                    throw new ArgumentOutOfRangeException("Xml Generation Lambdas are required");

                var contentType = (string)element.Attribute(PtOpenXml.ContentType);

                if (element.Name == W.t)
                    return element.Nodes().Select(z => ProduceXmlTransform(part, z, settings));

                if (contentType == null && element.Name == W.r)
                {
                    if (settings.XmlGenerationLambdas.ContainsKey("Run"))
                    {
                        var lamda = settings.XmlGenerationLambdas["Run"];
                        var newElement = lamda(contentType, part, element, settings);
                        return newElement;
                    }
                    else
                    {
                        throw new OpenXmlPowerToolsException("Entry for Run content type in XML generation lambdas is required");
                    }
                }

                if (element.Name == W.hyperlink)
                {
                    if (settings.XmlGenerationLambdas.ContainsKey("Hyperlink"))
                    {
                        var lamda = settings.XmlGenerationLambdas["Hyperlink"];
                        var newElement = lamda(contentType, part, element, settings);
                        return newElement;
                    }
                    else
                    {
                        throw new OpenXmlPowerToolsException("Entry for Hyperlink content type in XML generation lambdas is required");
                    }
                }

                if (contentType != null)
                {

                    if (settings.XmlGenerationLambdas != null)
                    {
                        if (settings.XmlGenerationLambdas.ContainsKey(contentType))
                        {
                            var lamda = settings.XmlGenerationLambdas[contentType];
                            var newElement = lamda(contentType, part, element, settings);

                            string lang = (string)element.Elements(W.pPr).Elements(W.rPr).Elements(W.lang).Attributes(W.val).FirstOrDefault();
                            if (lang == null)
                                lang = settings.DefaultLang;
                            if (lang != null && ! lang.StartsWith("en"))  // TODO we are not generating lang if English, but this needs revised after analysis
                            {
                                var n = newElement as XElement;
                                if (n != null)
                                {
                                    n.Add(new XAttribute("Lang", lang));
                                    return n;
                                }
                            }

                            return newElement;
                        }

                    }

                    // if no generation rules are set, or if there is no rule for this content type, then
                    // generate the default, for now.

                    // todo this is not ideal in my mind.  Need to think about this more.  Maybe every content type
                    // must have a generation lambda.

                    return new XElement(contentType, new XElement("Content",
                        element.Elements().Select(rce => ProduceXmlTransform(part, rce, settings))));
                }

                // ignore any other elements
                return null;
            }

#if false
            // The following code inserts an XML comment for unicode characters above 256

            // This could be made more efficient - group characters together and create fewer XText nodes.
            // As it is, it is pretty slow, so should be used only for debugging.

            var xt = node as XText;
            if (xt != null)
            {
                var newContent = xt.Value.Select(c =>
                {
                    var ic = (int)c;
                    if (ic < 256)
                        return (object)new XText(c.ToString());

                    return new[] {
                        (object)new XText(c.ToString()),
                        new XComment(ic.ToString("X")),
                    };
                })
                .ToList();
                return newContent;
            }
#endif

            return node;
        }

        private static void AssignLevelsToContent(XDocument mainXDoc, WmlToXmlSettings settings)
        {
            var contentWithContentType = mainXDoc
                .Root
                .Descendants()
                .Where(d => d.Name == W.p || d.Name == W.tbl || d.Name == W.tr || d.Name == W.tc)
                .Where(d => d.Attribute(PtOpenXml.ContentType) != null)
                .ToList();

            int currentLevel = 1;
            foreach (var content in contentWithContentType)
            {
                var thisLevel = GetIndentLevel(content, settings);
                if (thisLevel == null)
                {
                    content.Add(new XAttribute(PtOpenXml.Level, currentLevel));
                }
                else
                {
                    if (content.Attribute(PtOpenXml.Level) == null)
                        content.Add(new XAttribute(PtOpenXml.Level, thisLevel));
                    currentLevel = (int)thisLevel + 1;
                }
            }
        }

        private static void AssignLevelsToContent(XElement blockLevelContentContainer, WmlToXmlSettings settings)
        {
            var contentWithContentType = blockLevelContentContainer
                .Descendants()
                .Where(d => d.Name == W.p || d.Name == W.tbl || d.Name == W.tr || d.Name == W.tc)
                .Where(d => d.Attribute(PtOpenXml.ContentType) != null)
                .ToList();

            int currentLevel = 1;
            foreach (var content in contentWithContentType)
            {
                var thisLevel = GetIndentLevel(content, settings);
                if (thisLevel == null)
                {
                    content.Add(new XAttribute(PtOpenXml.Level, currentLevel));
                }
                else
                {
                    content.Add(new XAttribute(PtOpenXml.Level, thisLevel));
                    currentLevel = (int)thisLevel + 1;
                }
            }
        }

        private static int? GetIndentLevel(XElement blockLevelContent, WmlToXmlSettings settings)
        {
            return settings.ContentTypeHierarchyLambda(blockLevelContent, settings);
        }

        // Apply the Document rules first, then apply the DocumentType rules, then apply the Global rules.  First one that matches, wins.
        private static void ApplyContentTypesForRuleSet(WmlToXmlSettings settings, ContentTypeApplierInfo ctai, WordprocessingDocument wDoc)
        {
            ApplyRulesToPart(settings, ctai, wDoc, wDoc.MainDocumentPart);
            if (wDoc.MainDocumentPart.EndnotesPart != null)
                ApplyRulesToPart(settings, ctai, wDoc, wDoc.MainDocumentPart.EndnotesPart);
            if (wDoc.MainDocumentPart.FootnotesPart != null)
                ApplyRulesToPart(settings, ctai, wDoc, wDoc.MainDocumentPart.FootnotesPart);
        }

        private static void ApplyRulesToPart(WmlToXmlSettings settings, ContentTypeApplierInfo ctai, WordprocessingDocument wDoc, OpenXmlPart part)
        {
            var partXDoc = part.GetXDocument();
            var styleXDoc = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
            var blockContent = partXDoc.Descendants()
                .Where(d => d.Name == W.p || d.Name == W.tbl || d.Name == W.tr || d.Name == W.tc);

            int totalCount = 0;
            if (settings.ProgressFunction != null)
            {
                totalCount = blockContent.Count();
                string message;
                if (part is MainDocumentPart)
                    message = "Apply rules to main document part";
                else if (part is EndnotesPart)
                    message = "Apply rules to endnotes part";
                else
                    message = "Apply rules to footnotes part";
                WmlToXmlProgressInfo pi = new WmlToXmlProgressInfo()
                {
                    ContentTotal = totalCount,
                    ContentCount = 0,
                    InProgressMessage = message + Environment.NewLine,
                };
                settings.ProgressFunction(pi);
            }

            var count = 0;
            foreach (var blc in blockContent)
            {
                if (settings.ProgressFunction != null)
                {
                    ++count;
                    if (count < 50 || (count) % 10 == 0 || count == totalCount)
                    {
                        var msg = string.Format("  {0} of {1}", count, totalCount);
                        msg += "".PadRight(msg.Length, '\b');
                        WmlToXmlProgressInfo pi2 = new WmlToXmlProgressInfo()
                        {
                            ContentTotal = totalCount,
                            ContentCount = count,
                            InProgressMessage = msg,
                        };
                        settings.ProgressFunction(pi2);
                    }
                }

                string styleOfBlc = null;
                string styleOfBlcUC = null;
                if (blc.Name == W.p)
                {
                    var styleIdOfBlc = (string)blc.Elements(W.pPr).Elements(W.pStyle).Attributes(W.val).FirstOrDefault();
                    if (styleIdOfBlc != null)
                    {
                        styleOfBlc = (string)styleXDoc
                            .Root
                            .Elements(W.style)
                            .Where(s => (string)s.Attribute(W.styleId) == styleIdOfBlc && (string)s.Attribute(W.type) == "paragraph")
                            .Elements(W.name)
                            .Attributes(W.val)
                            .FirstOrDefault();
                    }
                    if (styleOfBlc == null)
                        styleOfBlc = ctai.DefaultParagraphStyleName;
                    styleOfBlcUC = styleOfBlc.ToUpper();
                }
                else if (blc.Name == W.tbl)
                {
                    var styleIdOfBlc = (string)blc.Elements(W.tblPr).Elements(W.tblStyle).Attributes(W.val).FirstOrDefault();
                    if (styleIdOfBlc != null)
                    {
                        styleOfBlc = (string)styleXDoc
                            .Root
                            .Elements(W.style)
                            .Where(s => (string)s.Attribute(W.styleId) == styleIdOfBlc && (string)s.Attribute(W.type) == "table")
                            .Elements(W.name)
                            .Attributes(W.val)
                            .FirstOrDefault();
                    }
                    if (styleOfBlc == null)
                        styleOfBlc = ctai.DefaultTableStyleName;
                    styleOfBlcUC = styleOfBlc.ToUpper();
                }

                ///////////////////////////////////////////////////////////////////////////////////////////
                // The following is useful to get a list of all content types and the code gen list

                //var contentTypeList = settings
                //    .DocumentContentTypeRules
                //    .Concat(settings.DocumentTypeContentTypeRules)
                //    .Concat(settings.GlobalContentTypeRules)
                //    .Select(ct => ct.ContentType)
                //    .Distinct()
                //    .OrderBy(n => n)
                //    .ToList();

                //var contentTypeCodeGenList = settings
                //    .XmlGenerationLambdas
                //    .Select(xgl => xgl.Key)
                //    .OrderBy(n => n)
                //    .ToList();

                //var rulesWithoutGenCode = contentTypeList
                //    .Except(contentTypeCodeGenList)
                //    .ToList();

                //var codeGenWithoutRules = contentTypeCodeGenList
                //    .Except(contentTypeList)
                //    .ToList();

                //var s10 = codeGenWithoutRules.Select(m => m + Environment.NewLine).StringConcatenate();
                //Console.WriteLine(s10);

                //var s9 = contentTypeList.Select(m => m + Environment.NewLine).StringConcatenate();
                //Console.WriteLine(s9);

                // Apply the Document rules first, then apply the DocumentType rules, then apply the Global rules.  First one that matches, wins.
                foreach (var rule in settings.DocumentContentTypeRules.Concat(settings.DocumentTypeContentTypeRules).Concat(settings.GlobalContentTypeRules))
                {
                    bool stylePass = false;
                    bool styleRegexPass = false;
                    bool regexPass = false;
                    bool matchLambdaPass = false;

                    stylePass = rule.StyleName == null || rule.StyleName.ToUpper() == styleOfBlcUC;

                    if (stylePass)
                    {
                        styleRegexPass = rule.StyleNameRegex == null;
                        if (rule.StyleNameRegex != null)
                            styleRegexPass = rule.StyleNameRegex.IsMatch(styleOfBlc);
                    }

                    if (stylePass && styleRegexPass)
                    {
                        regexPass = rule.RegexArray == null;
                        if (rule.RegexArray != null)
                        {
                            for (int i = 0; i < rule.RegexArray.Length; i++)
                            {
                                // clone the blc because OpenXmlRegex.Match replaces content, mucks with the run, probably should not if it only is used to find content.
                                var clonedBlc = new XElement(blc);

                                // following removes the subtitle created by a soft break, so that the pattern matches appropriately.
                                clonedBlc = RemoveContentAfterBR(clonedBlc);

#if false
<p p1:FontName="Georgia" p1:LanguageType="western" p1:AbstractNumId="28" xmlns:p1="http://powertools.codeplex.com/2011" xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <r p1:ListItemRun="1.1" p1:FontName="Georgia" p1:LanguageType="western">
    <t xml:space="preserve">1.1</t>
  </r>
#endif
                                // remove list item runs so that they are not matched in the content
                                clonedBlc.Elements(W.r).Where(r => r.Attribute(PtOpenXml.ListItemRun) != null).Remove();

                                if (OpenXmlRegex.Match(new[] { clonedBlc }, rule.RegexArray[i]) != 0)
                                {
                                    regexPass = true;
                                    break;
                                }
                            }
                        }
                    }

                    if (stylePass && styleRegexPass && regexPass)
                    {
                        matchLambdaPass = rule.MatchLambda == null;
                        if (rule.MatchLambda != null)
                        {
                            if (rule.MatchLambda(blc, rule, wDoc, settings))
                                matchLambdaPass = true;
                        }
                    }

                    if (stylePass && styleRegexPass && regexPass && matchLambdaPass)
                    {
                        AddContentTypeToBlockContent(settings, part, blc, rule.ContentType);
                        if (rule.ApplyRunContentTypes)
                            ApplyRunContentTypes(settings, ctai, wDoc, blc, settings.RunContentTypeRules, part, partXDoc);
                        break;
                    }
                }
            }

            if (settings.ProgressFunction != null)
            {
                WmlToXmlProgressInfo pi = new WmlToXmlProgressInfo()
                {
                    ContentTotal = totalCount,
                    ContentCount = totalCount,
                    InProgressMessage = Environment.NewLine + "  Done" + Environment.NewLine,
                };
                settings.ProgressFunction(pi);
            }

            part.PutXDocument();
            var mainPart = part as MainDocumentPart;
            if (mainPart != null)
            {
                if (mainPart.WordprocessingCommentsPart != null)
                    mainPart.WordprocessingCommentsPart.PutXDocument();
            }
        }

        private static XElement RemoveContentAfterBR(XElement clonedBlc)
        {
            if (clonedBlc.Name != W.p)
                return clonedBlc;
            var cloned2 = new XElement(clonedBlc.Name,
                clonedBlc.Attributes(),
                clonedBlc.Elements().TakeWhile(r => r.Element(W.br) == null));
            return cloned2;
        }

        private static void ApplyRunContentTypes(WmlToXmlSettings settings, ContentTypeApplierInfo ctai, WordprocessingDocument wDoc,
            XElement blockLevelContent, List<ContentTypeRule> runContentTypeRuleList, OpenXmlPart part, XDocument mainXDoc)
        {
            var runContent = blockLevelContent.Descendants()
                .Where(d => d.Name == W.r || d.Name == W.hyperlink || d.Name == W.sdt || d.Name == W.bookmarkStart);
            foreach (var rlc in runContent)
            {
                if (rlc.Name == W.r || rlc.Name == W.sdt)
                {
                    var runStyle = (string)rlc.Elements(W.rPr).Elements(W.rStyle).Attributes(W.val).FirstOrDefault();
                    if (runStyle == null)
                        runStyle = ctai.DefaultCharacterStyleName;
                    foreach (var rule in runContentTypeRuleList)
                    {
                        if (rule.StyleName != null && rule.StyleName != runStyle)
                            continue;

                        if (rule.RegexArray != null)
                            throw new OpenXmlPowerToolsException("Invalid Run ContentType Rule - Regex not allowed");
                        if (rule.MatchLambda != null)
                        {
                            if (rule.MatchLambda(rlc, rule, wDoc, settings))
                            {
                                AddContentTypeToRunContent(settings, part, rlc, rule.ContentType);
                                break;
                            }
                            continue;
                        }
                        AddContentTypeToRunContent(settings, part, rlc, rule.ContentType);
                        break;
                    }
                }
                else if (rlc.Name == W.hyperlink)
                {
                    foreach (var run in rlc.Descendants(W.r))
                        AddContentTypeToRunContent(settings, part, run, "Hyperlink");
                }
                else if (rlc.Name == W.bookmarkStart)
                {
                    AddContentTypeToRunContent(settings, part, rlc, "Anchor");
                }
            }
        }

        private static XAttribute[] NamespaceAttributes =
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
            new XAttribute(XNamespace.Xmlns + "w15", W15.w15),
            new XAttribute(XNamespace.Xmlns + "w16se", W16SE.w16se),
            new XAttribute(XNamespace.Xmlns + "wpg", WPG.wpg),
            new XAttribute(XNamespace.Xmlns + "wpi", WPI.wpi),
            new XAttribute(XNamespace.Xmlns + "wne", WNE.wne),
            new XAttribute(XNamespace.Xmlns + "wps", WPS.wps),
            new XAttribute(XNamespace.Xmlns + "pt", PtOpenXml.pt),
            new XAttribute(MC.Ignorable, "w14 wp14 w15 w16se pt"),
        };

        private static void AddContentTypeToBlockContent(WmlToXmlSettings settings, OpenXmlPart part, XElement blc, string contentType)
        {
            // add the attribute to the block content
            blc.Add(new XAttribute(PtOpenXml.ContentType, contentType));

            var mainPart = part as MainDocumentPart;
            if (mainPart != null)
            {
                // add a comment, if appropriate
                int commentNumber = 1;
                XDocument newComments = null;
                if (settings.InjectCommentForContentTypes != null && (bool)settings.InjectCommentForContentTypes)
                {
                    if (mainPart.WordprocessingCommentsPart != null)
                    {
                        newComments = mainPart.WordprocessingCommentsPart.GetXDocument();
                        newComments.Declaration.Standalone = "yes";
                        newComments.Declaration.Encoding = "UTF-8";
                        var ids = newComments.Root.Elements(W.comment).Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            commentNumber = ids.Max() + 1;
                    }
                    else
                    {
                        part.AddNewPart<WordprocessingCommentsPart>();
                        newComments = mainPart.WordprocessingCommentsPart.GetXDocument();
                        newComments.Declaration.Standalone = "yes";
                        newComments.Declaration.Encoding = "UTF-8";
                        newComments.Add(new XElement(W.comments, NamespaceAttributes));
                        commentNumber = 1;
                    }
#if false
  <w:comment w:id="12"
             w:author="Eric White"
             w:date="2016-03-20T18:50:00Z"
             w:initials="EW">
    <w:p w14:paraId="7E227B98"
         w14:textId="6FA2BE6B"
         w:rsidR="00425889"
         w:rsidRDefault="00425889">
      <w:pPr>
        <w:pStyle w:val="CommentText"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="CommentReference"/>
        </w:rPr>
        <w:annotationRef/>
      </w:r>
      <w:r>
        <w:t>Nil</w:t>
      </w:r>
    </w:p>
  </w:comment>
#endif
                    XElement newElement = new XElement(W.comment,
                        new XAttribute(W.id, commentNumber),
                        new XElement(W.p,
                            new XElement(W.pPr,
                                new XElement(W.pStyle,
                                    new XAttribute(W.val, "CommentText"))),
                            new XElement(W.r,
                                new XElement(W.rPr,
                                    new XElement(W.rStyle,
                                        new XAttribute(W.val, "CommentReference"))),
                                        new XElement(W.annotationRef)),
                            new XElement(W.r,
                                new XElement(W.t,
                                    new XText(contentType)))));
                    newComments.Root.Add(newElement);

#if false
      <w:r>
        <w:rPr>
          <w:rStyle w:val="CommentReference"/>
        </w:rPr>
        <w:commentReference w:id="12"/>
      </w:r>
#endif

                    XElement commentRun = new XElement(W.r,
                        new XElement(W.rPr,
                            new XElement(W.rStyle, new XAttribute(W.val, "CommentReference"))),
                        new XElement(W.commentReference,
                            new XAttribute(W.id, commentNumber)));
                    var firstRunInParagraph = blc
                        .DescendantsTrimmed(W.txbxContent)
                        .Where(r => r.Name == W.r)
                        .FirstOrDefault();
                    if (firstRunInParagraph != null)
                    {
                        // for now, only do the work of inserting a comment if it is easy.  For content types for tables, rows and cells, not inserting a comment.
                        if (firstRunInParagraph.Parent.Name == W.p)
                            firstRunInParagraph.AddBeforeSelf(commentRun);
                    }
                    else
                    {
                        // for now, only do the work of inserting a comment if it is easy.  For content types for tables, rows and cells, not inserting a comment.
                        if (blc.Name == W.p)
                            blc.Add(commentRun);
                    }

                    if (mainPart.StyleDefinitionsPart == null)
                    {
                        throw new ContentApplierException("Document does not have styles definition part");
                    }
                    XDocument stylesXDoc = mainPart.StyleDefinitionsPart.GetXDocument();

                    var style =
@"<w:style w:type=""paragraph""
           w:styleId=""CommentText""
           xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:name w:val=""annotation text""/>
    <w:basedOn w:val=""Normal""/>
    <w:link w:val=""CommentTextChar""/>
    <w:semiHidden/>
    <w:rPr>
      <w:sz w:val=""20""/>
      <w:szCs w:val=""20""/>
    </w:rPr>
  </w:style>
";
                    AddIfMissing(stylesXDoc, style);
                    style =
@"<w:style w:type=""paragraph""
           w:styleId=""CommentSubject""
           xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:name w:val=""annotation subject""/>
    <w:basedOn w:val=""CommentText""/>
    <w:next w:val=""CommentText""/>
    <w:semiHidden/>
    <w:rPr>
      <w:b/>
      <w:bCs/>
    </w:rPr>
  </w:style>
";
                    AddIfMissing(stylesXDoc, style);
                    style =
@"<w:style w:type=""character""
           w:styleId=""CommentReference""
           xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:name w:val=""annotation reference""/>
    <w:basedOn w:val=""DefaultParagraphFont""/>
    <w:uiPriority w:val=""99""/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:rsid w:val=""00872729""/>
    <w:rPr>
      <w:sz w:val=""16""/>
      <w:szCs w:val=""16""/>
    </w:rPr>
  </w:style>
";
                    AddIfMissing(stylesXDoc, style);
                    style =
@"<w:style w:type=""character""
           w:customStyle=""1""
           w:styleId=""CommentTextChar""
           xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:name w:val=""Comment Text Char""/>
    <w:basedOn w:val=""DefaultParagraphFont""/>
    <w:link w:val=""CommentText""/>
    <w:semiHidden/>
    <w:rsid w:val=""00A43CEC""/>
    <w:rPr>
      <w:lang w:val=""en-GB""
              w:eastAsia=""zh-CN""/>
    </w:rPr>
  </w:style>
";
                    AddIfMissing(stylesXDoc, style);
                    mainPart.StyleDefinitionsPart.PutXDocument();
                }
            }

            var root = blc.Ancestors().LastOrDefault();
            if (root == null)
                throw new ContentApplierException("Internal error");
            var ptNamespace = root.Attribute(XNamespace.Xmlns + "pt");
            if (ptNamespace == null)
            {
                root.Add(new XAttribute(XNamespace.Xmlns + "pt", PtOpenXml.pt.NamespaceName));
            }
            var ignorable = (string)root.Attribute(MC.Ignorable);
            if (ignorable != null)
            {
                var list = ignorable.Split(' ');
                if (!list.Contains("pt"))
                {
                    ignorable += " pt";
                    root.Attribute(MC.Ignorable).Value = ignorable;
                }
            }
            else
            {
                root.Add(new XAttribute(MC.Ignorable, "pt"));
            }
        }

        private static void AddContentTypeToRunContent(WmlToXmlSettings settings, OpenXmlPart part, XElement rlc, string contentType)
        {
            // if there is already a content type for this run level content, then nothing to do.  First one wins.
            if (rlc.Attribute(PtOpenXml.ContentType) != null)
                return;

            // add the attribute to the block level content
            rlc.Add(new XAttribute(PtOpenXml.ContentType, contentType));

            var mainPart = part as MainDocumentPart;
            if (mainPart != null)
            {
                // add a comment, if appropriate
                int commentNumber = 1;
                XDocument newComments = null;
                if (settings.InjectCommentForContentTypes != null && (bool)settings.InjectCommentForContentTypes)
                {
                    if (mainPart.WordprocessingCommentsPart != null)
                    {
                        newComments = mainPart.WordprocessingCommentsPart.GetXDocument();
                        newComments.Declaration.Standalone = "yes";
                        newComments.Declaration.Encoding = "UTF-8";
                        var ids = newComments.Root.Elements(W.comment).Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            commentNumber = ids.Max() + 1;
                    }
                    else
                    {
                        mainPart.AddNewPart<WordprocessingCommentsPart>();
                        newComments = mainPart.WordprocessingCommentsPart.GetXDocument();
                        newComments.Declaration.Standalone = "yes";
                        newComments.Declaration.Encoding = "UTF-8";
                        newComments.Add(new XElement(W.comments, NamespaceAttributes));
                        commentNumber = 1;
                    }
                    XElement newElement = new XElement(W.comment,
                        new XAttribute(W.id, commentNumber),
                        new XElement(W.p,
                            new XElement(W.pPr,
                                new XElement(W.pStyle,
                                    new XAttribute(W.val, "CommentText"))),
                            new XElement(W.r,
                                new XElement(W.rPr,
                                    new XElement(W.rStyle,
                                        new XAttribute(W.val, "CommentReference"))),
                                        new XElement(W.annotationRef)),
                            new XElement(W.r,
                                new XElement(W.t,
                                    new XText(contentType)))));
                    newComments.Root.Add(newElement);
                    XElement commentRun = new XElement(W.r,
                        new XElement(W.rPr,
                            new XElement(W.rStyle, new XAttribute(W.val, "CommentReference"))),
                        new XElement(W.commentReference,
                            new XAttribute(W.id, commentNumber)));
                    var firstRunInParagraph = rlc
                        .DescendantsTrimmed(W.txbxContent)
                        .Where(r => r.Name == W.r)
                        .FirstOrDefault();

                    // for now, only do the work of inserting a comment if it is easy.  For content types for tables, rows and cells, not inserting a comment.
                    if (rlc.Parent.Name == W.p)
                        rlc.AddBeforeSelf(commentRun);
                    if (mainPart.StyleDefinitionsPart == null)
                    {
                        throw new ContentApplierException("Document does not have styles definition part");
                    }
                    XDocument stylesXDoc = mainPart.StyleDefinitionsPart.GetXDocument();

                    var style =
@"<w:style w:type=""paragraph""
           w:styleId=""CommentText""
           xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:name w:val=""annotation text""/>
    <w:basedOn w:val=""Normal""/>
    <w:link w:val=""CommentTextChar""/>
    <w:semiHidden/>
    <w:rPr>
      <w:sz w:val=""20""/>
      <w:szCs w:val=""20""/>
    </w:rPr>
  </w:style>
";
                    AddIfMissing(stylesXDoc, style);
                    style =
@"<w:style w:type=""paragraph""
           w:styleId=""CommentSubject""
           xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:name w:val=""annotation subject""/>
    <w:basedOn w:val=""CommentText""/>
    <w:next w:val=""CommentText""/>
    <w:semiHidden/>
    <w:rPr>
      <w:b/>
      <w:bCs/>
    </w:rPr>
  </w:style>
";
                    AddIfMissing(stylesXDoc, style);
                    style =
@"<w:style w:type=""character""
           w:styleId=""CommentReference""
           xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:name w:val=""annotation reference""/>
    <w:basedOn w:val=""DefaultParagraphFont""/>
    <w:uiPriority w:val=""99""/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:rsid w:val=""00872729""/>
    <w:rPr>
      <w:sz w:val=""16""/>
      <w:szCs w:val=""16""/>
    </w:rPr>
  </w:style>
";
                    AddIfMissing(stylesXDoc, style);
                    style =
@"<w:style w:type=""character""
           w:customStyle=""1""
           w:styleId=""CommentTextChar""
           xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:name w:val=""Comment Text Char""/>
    <w:basedOn w:val=""DefaultParagraphFont""/>
    <w:link w:val=""CommentText""/>
    <w:semiHidden/>
    <w:rsid w:val=""00A43CEC""/>
    <w:rPr>
      <w:lang w:val=""en-GB""
              w:eastAsia=""zh-CN""/>
    </w:rPr>
  </w:style>
";
                    AddIfMissing(stylesXDoc, style);
                    mainPart.StyleDefinitionsPart.PutXDocument();
                }
            }

            var root = rlc.Ancestors().LastOrDefault();
            if (root == null)
                throw new ContentApplierException("Internal error");
            var ptNamespace = root.Attribute(XNamespace.Xmlns + "pt");
            if (ptNamespace == null)
            {
                root.Add(new XAttribute(XNamespace.Xmlns + "pt", PtOpenXml.pt.NamespaceName));
            }
            var ignorable = (string)root.Attribute(MC.Ignorable);
            if (ignorable != null)
            {
                var list = ignorable.Split(' ');
                if (!list.Contains("pt"))
                {
                    ignorable += " pt";
                    root.Attribute(MC.Ignorable).Value = ignorable;
                }
            }
            else
            {
                root.Add(new XAttribute(MC.Ignorable, "pt"));
            }
        }

        private static void AddIfMissing(XDocument stylesXDoc, string commentStyle)
        {
            XElement e1 = XElement.Parse(commentStyle);
#if false
  <w:style w:type=""character""
           w:customStyle=""1""
           w:styleId=""CommentTextChar""
#endif
            var existingStyle = stylesXDoc
                .Root
                .Elements(W.style)
                .FirstOrDefault(e2 =>
                {
                    XName name = W.type;
                    string v1 = (string)e1.Attribute(name);
                    string v2 = (string)e2.Attribute(name);
                    if (v1 != v2)
                        return false;
                    name = W.customStyle;
                    v1 = (string)e1.Attribute(name);
                    v2 = (string)e2.Attribute(name);
                    if (v1 != v2)
                        return false;
                    name = W.styleId;
                    v1 = (string)e1.Attribute(name);
                    v2 = (string)e2.Attribute(name);
                    if (v1 != v2)
                        return false;
                    return true;
                });
            if (existingStyle != null)
                return;
            stylesXDoc.Root.Add(e1);
        }

        private static void AssembleListItemInformation(WordprocessingDocument wordDoc, ListItemRetrieverSettings settings)
        {
            XDocument xDoc = wordDoc.MainDocumentPart.GetXDocument();
            foreach (var para in xDoc.Descendants(W.p))
            {
                ListItemRetriever.RetrieveListItem(wordDoc, para, settings);
            }
        }

        private class ContentTypeApplierInfo
        {
            public string DefaultParagraphStyleName;
            public string DefaultCharacterStyleName;
            public string DefaultTableStyleName;
            public ContentTypeApplierInfo()
            {
            }
        }

        public class ContentApplierException : Exception
        {
            public ContentApplierException(string message) : base(message) { }
        }
    }
}
