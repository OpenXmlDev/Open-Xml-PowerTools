// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;

namespace OpenXmlPowerTools
{
    internal class MetricsGetter01
    {
        private static void Main(string[] args)
        {
            MetricsGetterSettings settings = null;
            FileInfo fi = null;

            fi = new FileInfo("../../ContentControls.docx");
            settings = new MetricsGetterSettings
            {
                IncludeTextInContentControls = false
            };
            Console.WriteLine("============== No text from content controls ==============");
            Console.WriteLine(fi.FullName);
            Console.WriteLine(MetricsGetter.GetMetrics(fi.FullName, settings));
            Console.WriteLine();

            fi = new FileInfo("../../ContentControls.docx");
            settings = new MetricsGetterSettings
            {
                IncludeTextInContentControls = true
            };
            Console.WriteLine("============== With text from content controls ==============");
            Console.WriteLine(fi.FullName);
            Console.WriteLine(MetricsGetter.GetMetrics(fi.FullName, settings));
            Console.WriteLine();

            fi = new FileInfo("../../TrackedRevisions.docx");
            settings = new MetricsGetterSettings
            {
                IncludeTextInContentControls = true
            };
            Console.WriteLine("============== Tracked Revisions ==============");
            Console.WriteLine(fi.FullName);
            Console.WriteLine(MetricsGetter.GetMetrics(fi.FullName, settings));
            Console.WriteLine();

            fi = new FileInfo("../../Styles.docx");
            settings = new MetricsGetterSettings
            {
                IncludeTextInContentControls = false
            };
            Console.WriteLine("============== Style Hierarchy ==============");
            Console.WriteLine(fi.FullName);
            Console.WriteLine(MetricsGetter.GetMetrics(fi.FullName, settings));
            Console.WriteLine();

            fi = new FileInfo("../../Tables.xlsx");
            settings = new MetricsGetterSettings
            {
                IncludeTextInContentControls = false,
                IncludeXlsxTableCellData = true
            };
            Console.WriteLine("============== Spreadsheet Tables ==============");
            Console.WriteLine(fi.FullName);
            Console.WriteLine(MetricsGetter.GetMetrics(fi.FullName, settings));
            Console.WriteLine();
        }
    }
}