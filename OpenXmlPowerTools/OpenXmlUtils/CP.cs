using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class CP
    {
        public static readonly XNamespace cp =
            "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

        public static readonly XName category = cp + "category";
        public static readonly XName contentStatus = cp + "contentStatus";
        public static readonly XName contentType = cp + "contentType";
        public static readonly XName coreProperties = cp + "coreProperties";
        public static readonly XName keywords = cp + "keywords";
        public static readonly XName lastModifiedBy = cp + "lastModifiedBy";
        public static readonly XName lastPrinted = cp + "lastPrinted";
        public static readonly XName revision = cp + "revision";
    }
}