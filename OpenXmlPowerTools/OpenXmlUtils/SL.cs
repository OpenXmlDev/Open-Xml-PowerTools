using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class SL
    {
        public static readonly XNamespace sl = "http://schemas.openxmlformats.org/schemaLibrary/2006/main";

        public static readonly XName manifestLocation = sl + "manifestLocation";
        public static readonly XName schema = sl + "schema";
        public static readonly XName schemaLibrary = sl + "schemaLibrary";
        public static readonly XName uri = sl + "uri";
    }
}