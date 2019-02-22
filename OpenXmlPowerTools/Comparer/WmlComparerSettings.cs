// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using System.IO;

namespace OpenXmlPowerTools
{
    public class WmlComparerSettings
    {
        public char[] WordSeparators;
        public string AuthorForRevisions = "Open-Xml-PowerTools";
        public string DateTimeForRevisions = DateTime.Now.ToString("o");
        public double DetailThreshold = 0.15;
        public bool CaseInsensitive = false;
        public CultureInfo CultureInfo = null;
        public Action<string> LogCallback = null;
        public int StartingIdForFootnotesEndnotes = 1;

        public DirectoryInfo DebugTempFileDi;

        public WmlComparerSettings()
        {
            // note that , and . are processed explicitly to handle cases where they are in a number or word
            WordSeparators = new[] { ' ', '-', ')', '(', ';', ',' }; // todo need to fix this for complete list
        }
    }
}
