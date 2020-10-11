using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class Pic
    {
        public static readonly XNamespace pic =
            "http://schemas.openxmlformats.org/drawingml/2006/picture";

        public static readonly XName blipFill = pic + "blipFill";
        public static readonly XName cNvPicPr = pic + "cNvPicPr";
        public static readonly XName cNvPr = pic + "cNvPr";
        public static readonly XName nvPicPr = pic + "nvPicPr";
        public static readonly XName _pic = pic + "pic";
        public static readonly XName spPr = pic + "spPr";
    }
}