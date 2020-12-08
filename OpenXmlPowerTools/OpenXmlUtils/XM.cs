using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class XM
    {
        public static readonly XNamespace xm = "http://schemas.microsoft.com/office/excel/2006/main";

        public static readonly XName f = xm + "f";
        public static readonly XName _ref = xm + "ref";
        public static readonly XName sqref = xm + "sqref";
    }
}