using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class WP
    {
        public static readonly XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

        public static readonly XName align = wp + "align";
        public static readonly XName anchor = wp + "anchor";
        public static readonly XName cNvGraphicFramePr = wp + "cNvGraphicFramePr";
        public static readonly XName docPr = wp + "docPr";
        public static readonly XName effectExtent = wp + "effectExtent";
        public static readonly XName extent = wp + "extent";
        public static readonly XName inline = wp + "inline";
        public static readonly XName lineTo = wp + "lineTo";
        public static readonly XName positionH = wp + "positionH";
        public static readonly XName positionV = wp + "positionV";
        public static readonly XName posOffset = wp + "posOffset";
        public static readonly XName simplePos = wp + "simplePos";
        public static readonly XName start = wp + "start";
        public static readonly XName wrapNone = wp + "wrapNone";
        public static readonly XName wrapPolygon = wp + "wrapPolygon";
        public static readonly XName wrapSquare = wp + "wrapSquare";
        public static readonly XName wrapThrough = wp + "wrapThrough";
        public static readonly XName wrapTight = wp + "wrapTight";
        public static readonly XName wrapTopAndBottom = wp + "wrapTopAndBottom";
    }
}