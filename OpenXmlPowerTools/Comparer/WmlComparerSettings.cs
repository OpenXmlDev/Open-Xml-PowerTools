// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using System.IO;

namespace OpenXmlPowerTools
{
    public class WmlComparerSettings
    {
        public char[] WordSeparators { get; set; }
        public string AuthorForRevisions { get; set; } = "Open-Xml-PowerTools";
        public string DateTimeForRevisions { get; set; } = DateTime.Now.ToString("o");
        public double DetailThreshold { get; set; } = 0.15;
        public bool CaseInsensitive { get; set; } = false;
        public CultureInfo CultureInfo { get; set; } = null;
        public Action<string> LogCallback { get; set; } = null;
        public int StartingIdForFootnotesEndnotes { get; set; } = 1;

        public DirectoryInfo DebugTempFileDi { get; set; }

        public WmlComparerSettings()
        {
            // note that , and . are processed explicitly to handle cases where they are in a number or word
            WordSeparators = new[] { ' ', '-', ')', '(', ';', ',' };
        }
    }
}