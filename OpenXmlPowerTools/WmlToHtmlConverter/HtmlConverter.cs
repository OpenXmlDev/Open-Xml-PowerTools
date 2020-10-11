using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class HtmlConverter
    {
        public static XElement ConvertToHtml(WmlDocument wmlDoc, HtmlConverterSettings htmlConverterSettings)
        {
            var settings = new WmlToHtmlConverterSettings(htmlConverterSettings);
            return WmlToHtmlConverter.ConvertToHtml(wmlDoc, settings);
        }

        public static XElement ConvertToHtml(WordprocessingDocument wDoc, HtmlConverterSettings htmlConverterSettings)
        {
            var settings = new WmlToHtmlConverterSettings(htmlConverterSettings);
            return WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
        }
    }
}