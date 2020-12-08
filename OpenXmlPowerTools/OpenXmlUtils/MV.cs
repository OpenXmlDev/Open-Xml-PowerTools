using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class MV
    {
        public static readonly XNamespace mv = "urn:schemas-microsoft-com:mac:vml";

        public static readonly XName blur = mv + "blur";
        public static readonly XName complextextbox = mv + "complextextbox";
    }
}