using Codeuctivity.OpenXmlPowerTools;
using System;
using System.IO;

namespace MetricsGetter01
{
    internal class MetricsGetter01
    {
        private static void Main()
        {
            var fi = new FileInfo("../../ContentControls.docx");
            var settings = new MetricsGetterSettings
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