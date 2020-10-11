using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class DC
    {
        public static readonly XNamespace dc = "http://purl.org/dc/elements/1.1/";

        public static readonly XName creator = dc + "creator";
        public static readonly XName description = dc + "description";
        public static readonly XName subject = dc + "subject";
        public static readonly XName title = dc + "title";
    }
}