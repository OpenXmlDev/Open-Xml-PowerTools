using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class MC
    {
        public static readonly XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        public static readonly XName AlternateContent = mc + "AlternateContent";
        public static readonly XName Choice = mc + "Choice";
        public static readonly XName Fallback = mc + "Fallback";
        public static readonly XName Ignorable = mc + "Ignorable";
        public static readonly XName PreserveAttributes = mc + "PreserveAttributes";
    }
}