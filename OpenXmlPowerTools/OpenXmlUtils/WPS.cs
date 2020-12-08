using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class WPS
    {
        public static readonly XNamespace wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

        public static readonly XName altTxbx = wps + "altTxbx";
        public static readonly XName bodyPr = wps + "bodyPr";
        public static readonly XName cNvSpPr = wps + "cNvSpPr";
        public static readonly XName spPr = wps + "spPr";
        public static readonly XName style = wps + "style";
        public static readonly XName textbox = wps + "textbox";
        public static readonly XName txbx = wps + "txbx";
        public static readonly XName wsp = wps + "wsp";
    }
}