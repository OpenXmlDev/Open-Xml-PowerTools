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
        public string PageTitle { get; set; } = default!;

        /// <summary>
        /// Css class name prefix
        /// </summary>
        public string CssClassPrefix { get; set; } = "pt-";

        public bool FabricateCssClasses { get; set; } = true;
        public string GeneralCss { get; set; } = "span { white-space: pre-wrap; }";
        public string AdditionalCss { get; set; } = default!;
        public bool RestrictToSupportedLanguages { get; set; }
        public bool RestrictToSupportedNumberingFormats { get; set; }

        public Dictionary<string, Func<int, string, string>> ListItemImplementations { get; set; } = ListItemRetrieverSettings.DefaultListItemTextImplementations;

        /// <summary>
        /// Image handler
        /// </summary>
        public IImageHandler ImageHandler { get; }

        /// <summary>
        /// Handler that get applied to w:t
        /// </summary>
        public IWordprocessingTextHandler WordprocessingTextHandler { get; }

        public IWordprocessingSymbolHandler WordprocessingSymbolHandler { get; }

        /// <summary>
        /// Default ctor WmlToHtmlConverterSettings
        /// </summary>
        public WmlToHtmlConverterSettings()
        {
            ImageHandler = new DefaultImageHandler();
            WordprocessingTextHandler = new WordprocessingTextDummyHandler();
            WordprocessingSymbolHandler = new DefaultSymbolHandler();
        }

        /// <summary>
        /// Ctor WmlToHtmlConverterSettings
        /// </summary>
        /// <param name="customImageHandler">Handler used to convert open XML images to HTML images</param>
        /// <param name="wordprocessingTextHandler">Handler used to convert open XML text to HTML compatible text</param>
        /// <param name="wordprocessingSymbolHandler">Handler used to convert open XML symbols to HTML compatible text</param>
        public WmlToHtmlConverterSettings(IImageHandler customImageHandler, IWordprocessingTextHandler wordprocessingTextHandler, IWordprocessingSymbolHandler wordprocessingSymbolHandler)
        {
            ImageHandler = customImageHandler;
            WordprocessingTextHandler = wordprocessingTextHandler;
            WordprocessingSymbolHandler = wordprocessingSymbolHandler;
        }
    }
}