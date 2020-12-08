using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class COM
    {
        public static readonly XNamespace com =
            "http://schemas.openxmlformats.org/drawingml/2006/compatibility";

        public static readonly XName legacyDrawing = com + "legacyDrawing";
    }
}