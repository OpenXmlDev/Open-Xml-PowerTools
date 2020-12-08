using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class X
    {
        public static readonly XNamespace x =
            "urn:schemas-microsoft-com:office:excel";

        public static readonly XName Anchor = x + "Anchor";
        public static readonly XName AutoFill = x + "AutoFill";
        public static readonly XName ClientData = x + "ClientData";
        public static readonly XName Column = x + "Column";
        public static readonly XName MoveWithCells = x + "MoveWithCells";
        public static readonly XName Row = x + "Row";
        public static readonly XName SizeWithCells = x + "SizeWithCells";
    }
}