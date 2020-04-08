// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public class WmlToHtmlConverterSettings
    {
        public string PageTitle { get; set; }
        public string CssClassPrefix { get; set; }
        public bool FabricateCssClasses { get; set; }
        public string GeneralCss { get; set; }
        public string AdditionalCss { get; set; }
        public bool RestrictToSupportedLanguages { get; set; }
        public bool RestrictToSupportedNumberingFormats { get; set; }
        public Dictionary<string, Func<string, int, string, string>> ListItemImplementations { get; set; }
        public Func<ImageInfo, XElement> ImageHandler { get; set; }

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
        }

        public WmlToHtmlConverterSettings(HtmlConverterSettings htmlConverterSettings)
        {
            PageTitle = htmlConverterSettings.PageTitle;
            CssClassPrefix = htmlConverterSettings.CssClassPrefix;
            FabricateCssClasses = htmlConverterSettings.FabricateCssClasses;
            GeneralCss = htmlConverterSettings.GeneralCss;
            AdditionalCss = htmlConverterSettings.AdditionalCss;
            RestrictToSupportedLanguages = htmlConverterSettings.RestrictToSupportedLanguages;
            RestrictToSupportedNumberingFormats = htmlConverterSettings.RestrictToSupportedNumberingFormats;
            ListItemImplementations = htmlConverterSettings.ListItemImplementations;
            ImageHandler = htmlConverterSettings.ImageHandler;
        }
    }
}