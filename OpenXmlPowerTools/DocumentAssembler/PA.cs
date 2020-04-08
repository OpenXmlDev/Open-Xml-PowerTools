// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public partial class DocumentAssembler
    {
        private class PA
        {
            public static XName Content = "Content";
            public static XName Table = "Table";
            public static XName Repeat = "Repeat";
            public static XName EndRepeat = "EndRepeat";
            public static XName Conditional = "Conditional";
            public static XName EndConditional = "EndConditional";

            public static XName Select = "Select";
            public static XName Optional = "Optional";
            public static XName Match = "Match";
            public static XName NotMatch = "NotMatch";
            public static XName Depth = "Depth";
        }
    }
}