using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class DCTERMS
    {
        public static readonly XNamespace dcterms = "http://purl.org/dc/terms/";

        public static readonly XName created = dcterms + "created";
        public static readonly XName modified = dcterms + "modified";
    }
}