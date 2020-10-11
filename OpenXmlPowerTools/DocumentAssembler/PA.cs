

using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public partial class DocumentAssembler
    {
        private static class PA
        {
            public static readonly XName Content = "Content";
            public static readonly XName Table = "Table";
            public static readonly XName Repeat = "Repeat";
            public static readonly XName EndRepeat = "EndRepeat";
            public static readonly XName Conditional = "Conditional";
            public static readonly XName EndConditional = "EndConditional";

            public static readonly XName Select = "Select";
            public static readonly XName Optional = "Optional";
            public static readonly XName Match = "Match";
            public static readonly XName NotMatch = "NotMatch";
            public static readonly XName Depth = "Depth";
        }
    }
}