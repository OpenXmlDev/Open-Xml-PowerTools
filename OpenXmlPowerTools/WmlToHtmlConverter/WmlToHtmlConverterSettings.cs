using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public class WmlToHtmlConverterSettings
    {
        /// <summary>
        /// Pate title of Html
        /// </summary>
        public string PageTitle { get; set; }

        /// <summary>
        /// Css class name prefix
        /// </summary>
        public string CssClassPrefix { get; set; }

        public bool FabricateCssClasses { get; set; }
        public string GeneralCss { get; set; }
        public string AdditionalCss { get; set; }
        public bool RestrictToSupportedLanguages { get; set; }
        public bool RestrictToSupportedNumberingFormats { get; set; }
        public Dictionary<string, Func<int, string, string>> ListItemImplementations { get; set; }

        /// <summary>
        /// Image handler
        /// </summary>
        public Func<ImageInfo, XElement> ImageHandler { get; set; }

        /// <summary>
        /// Handler that get applied to w:t
        /// </summary>
        public IWordprocessingTextHandler WordprocessingTextHandler { get; }

        public WmlToHtmlConverterSettings()
        {
            PageTitle = "";
            CssClassPrefix = "pt-";
            FabricateCssClasses = true;
            GeneralCss = "span { white-space: pre-wrap; }";
            AdditionalCss = "";
            RestrictToSupportedLanguages = false;
            RestrictToSupportedNumberingFormats = false;
            ListItemImplementations = ListItemRetrieverSettings.DefaultListItemTextImplementations;
            WordprocessingTextHandler = new WordprocessingTextSymbolToUnicodeHandler();
        }
    }
}