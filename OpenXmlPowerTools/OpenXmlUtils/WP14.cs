using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class WP14
    {
        public static readonly XNamespace wp14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";

        public static readonly XName anchorId = wp14 + "anchorId";
        public static readonly XName editId = wp14 + "editId";
        public static readonly XName pctHeight = wp14 + "pctHeight";
        public static readonly XName pctPosVOffset = wp14 + "pctPosVOffset";
        public static readonly XName pctWidth = wp14 + "pctWidth";
        public static readonly XName sizeRelH = wp14 + "sizeRelH";
        public static readonly XName sizeRelV = wp14 + "sizeRelV";
    }
}