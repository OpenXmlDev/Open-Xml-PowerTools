using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class ACTIVEX
    {
        public static readonly XNamespace activex = "http://schemas.microsoft.com/office/2006/activeX";

        public static readonly XName classid = activex + "classid";
        public static readonly XName font = activex + "font";
        public static readonly XName license = activex + "license";
        public static readonly XName name = activex + "name";
        public static readonly XName ocx = activex + "ocx";
        public static readonly XName ocxPr = activex + "ocxPr";
        public static readonly XName persistence = activex + "persistence";
        public static readonly XName value = activex + "value";
    }
}