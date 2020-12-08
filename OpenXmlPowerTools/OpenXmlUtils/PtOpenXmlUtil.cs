using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class PtOpenXmlUtil
    {
        public static readonly XNamespace mp = "http://schemas.microsoft.com/office/mac/powerpoint/2008/main";

        public static readonly XName cube = mp + "cube";
        public static readonly XName flip = mp + "flip";
        public static readonly XName transition = mp + "transition";
    }
}