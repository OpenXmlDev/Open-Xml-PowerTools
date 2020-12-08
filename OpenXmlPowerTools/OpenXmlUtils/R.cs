using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class R
    {
        public static readonly XNamespace r =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        public static readonly XName blip = r + "blip";
        public static readonly XName cs = r + "cs";
        public static readonly XName dm = r + "dm";
        public static readonly XName embed = r + "embed";
        public static readonly XName href = r + "href";
        public static readonly XName id = r + "id";
        public static readonly XName link = r + "link";
        public static readonly XName lo = r + "lo";
        public static readonly XName pict = r + "pict";
        public static readonly XName qs = r + "qs";
        public static readonly XName verticalDpi = r + "verticalDpi";
    }
}