using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class Plegacy
    {
        public static readonly XNamespace plegacy = "urn:schemas-microsoft-com:office:powerpoint";
        public static readonly XName textdata = plegacy + "textdata";
    }
}