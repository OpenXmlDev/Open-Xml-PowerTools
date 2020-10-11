

using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public class HtmlConverterSettings
    {
        public string PageTitle { get; set; }
        public string CssClassPrefix { get; set; }
        public bool FabricateCssClasses { get; set; }
        public string GeneralCss { get; set; }
        public string AdditionalCss { get; set; }
        public bool RestrictToSupportedLanguages { get; set; }
        public bool RestrictToSupportedNumberingFormats { get; set; }
        public Dictionary<string, Func<int, string, string>> ListItemImplementations { get; set; }
        public Func<ImageInfo, XElement> ImageHandler { get; set; }

        public HtmlConverterSettings()
        {
            PageTitle = "";
            CssClassPrefix = "pt-";
            FabricateCssClasses = true;
            GeneralCss = "span { white-space: pre-wrap; }";
            AdditionalCss = "";
            RestrictToSupportedLanguages = false;
            RestrictToSupportedNumberingFormats = false;
            ListItemImplementations = ListItemRetrieverSettings.DefaultListItemTextImplementations;
        }
    }
}