using System;
using System.Collections.Generic;

namespace OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter
{
    /// <summary>
    /// WmlToHtmlConverterSettings
    /// </summary>
    public class WmlToHtmlConverterSettings
    {
        /// <summary>
        /// Page title of HTML
        /// </summary>
        public string PageTitle { get; }

        /// <summary>
        /// CSS class name prefix
        /// </summary>
        public string CssClassPrefix { get; }

        /// <summary>
        /// If FabricateCssClasses is true, CSS Classes will be generated instead of using inline styles
        /// </summary>
        public bool FabricateCssClasses { get; }

        public string GeneralCss { get; }
        public string AdditionalCss { get; }
        public bool RestrictToSupportedLanguages { get; }
        public bool RestrictToSupportedNumberingFormats { get; }

        public Dictionary<string, Func<int, string, string>> ListItemImplementations { get; set; } = ListItemRetrieverSettings.DefaultListItemTextImplementations;

        /// <summary>
        /// Image handler
        /// </summary>
        public IImageHandler ImageHandler { get; }

        /// <summary>
        /// Break handler
        /// </summary>
        public IBreakHandler BreakHandler { get; }

        /// <summary>
        /// Handler that get applied to w:t
        /// </summary>
        public ITextHandler TextHandler { get; }

        /// <summary>
        /// Symbol handler
        /// </summary>
        public ISymbolHandler SymbolHandler { get; }

        /// <summary>
        /// Default ctor WmlToHtmlConverterSettings
        /// </summary>
        /// <param name="pageTitle">Page title</param>
        public WmlToHtmlConverterSettings(string pageTitle)
        {
            AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }";
            GeneralCss = "span { white-space: pre-wrap; }";
            PageTitle = pageTitle;
            FabricateCssClasses = true;
            CssClassPrefix = "pt-";
            ImageHandler = new ImageHandler();
            TextHandler = new TextDummyHandler();
            SymbolHandler = new SymbolHandler();
            BreakHandler = new BreakHandler();
        }

        /// <summary>
        /// Ctor WmlToHtmlConverterSettings
        /// </summary>
        /// <param name="pageTitle">Page title</param>
        /// <param name="customImageHandler">Handler used to convert open XML images to HTML images</param>
        /// <param name="textHandler">Handler used to convert open XML text to HTML compatible text</param>
        /// <param name="symbolHandler">Handler used to convert open XML symbols to HTML compatible text</param>
        /// <param name="breakHandler">Handler used to convert open XML breaks to HTML equivalent</param>
        /// <param name="fabricateCssClasses">Set to true, if CSS style should be stored in classes instead of an inline attribute on each node</param>
        public WmlToHtmlConverterSettings(string pageTitle, IImageHandler customImageHandler, ITextHandler textHandler, ISymbolHandler symbolHandler, IBreakHandler breakHandler, bool fabricateCssClasses)
        {
            AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }";
            GeneralCss = "span { white-space: pre-wrap; }";
            PageTitle = pageTitle;
            FabricateCssClasses = fabricateCssClasses;
            CssClassPrefix = "pt-";
            ImageHandler = customImageHandler;
            TextHandler = textHandler;
            SymbolHandler = symbolHandler;
            BreakHandler = breakHandler;
        }
    }
}