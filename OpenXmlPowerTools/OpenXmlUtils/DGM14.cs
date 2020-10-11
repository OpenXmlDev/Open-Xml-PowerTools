using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class DGM14
    {
        public static readonly XNamespace dgm14 =
            "http://schemas.microsoft.com/office/drawing/2010/diagram";

        public static readonly XName cNvPr = dgm14 + "cNvPr";
        public static readonly XName recolorImg = dgm14 + "recolorImg";
    }
}