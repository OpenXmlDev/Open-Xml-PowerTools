using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class W10
    {
        public static readonly XNamespace w10 = "urn:schemas-microsoft-com:office:word";

        public static readonly XName anchorlock = w10 + "anchorlock";
        public static readonly XName borderbottom = w10 + "borderbottom";
        public static readonly XName borderleft = w10 + "borderleft";
        public static readonly XName borderright = w10 + "borderright";
        public static readonly XName bordertop = w10 + "bordertop";
        public static readonly XName wrap = w10 + "wrap";
    }
}