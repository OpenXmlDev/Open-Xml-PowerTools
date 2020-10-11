using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class PAV
    {
        public static readonly XNamespace pav = "http://schemas.microsoft.com/office/2007/6/19/audiovideo";
        public static readonly XName media = pav + "media";
        public static readonly XName srcMedia = pav + "srcMedia";
        public static readonly XName bmkLst = pav + "bmkLst";
    }
}