using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class INK
    {
        public static readonly XNamespace ink = "http://schemas.microsoft.com/ink/2010/main";

        public static readonly XName context = ink + "context";
        public static readonly XName sourceLink = ink + "sourceLink";
    }
}