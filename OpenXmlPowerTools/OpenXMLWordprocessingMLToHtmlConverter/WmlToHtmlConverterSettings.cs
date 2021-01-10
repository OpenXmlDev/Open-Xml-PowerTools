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

        /// <summary>
        /// Default ctor WmlToHtmlConverterSettings
        /// </summary>
        public WmlToHtmlConverterSettings()
        {
            ImageHandler = new DefaultImageHandler();
            WordprocessingTextHandler = new WordprocessingTextDummyHandler();
        }

        /// <summary>
        /// Ctor WmlToHtmlConverterSettings
        /// </summary>
        /// <param name="customImageHandler"></param>
        /// <param name="wordprocessingTextHandler"></param>
        public WmlToHtmlConverterSettings(IImageHandler customImageHandler, IWordprocessingTextHandler wordprocessingTextHandler)
        {
            ImageHandler = customImageHandler;
            WordprocessingTextHandler = wordprocessingTextHandler;
        }
    }
}