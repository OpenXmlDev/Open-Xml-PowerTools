using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class LC
    {
        public static readonly XNamespace lc = "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas";

        public static readonly XName lockedCanvas = lc + "lockedCanvas";
    }
}