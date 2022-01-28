using Codeuctivity;
using DocumentFormat.OpenXml.Packaging;
using MarkupSimplifierApp.Properties;
using System;

namespace MarkupSimplifierApp
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Example output files are in a DateTime stamped directory in ./bin/debug.  The directory name is ExampleOutput-yy-mm-dd-hhmmss.");
                Console.WriteLine("If you are building in release mode, they will, of course, be in ./bin/release.");
                Console.WriteLine("MarkupSimplifierApp.exe 1.docx 2.docx");
            }

            foreach (var item in args)
            {
                using var doc = WordprocessingDocument.Open(item, true);
                var settings = new SimplifyMarkupSettings
                {
                    RemoveContentControls = Settings.Default.RemoveContentControls,
                    RemoveSmartTags = Settings.Default.RemoveSmartTags,
                    RemoveRsidInfo = Settings.Default.RemoveRsidInfo,
                    RemoveComments = Settings.Default.RemoveComments,
                    RemoveEndAndFootNotes = Settings.Default.RemoveEndAndFootNotes,
                    ReplaceTabsWithSpaces = Settings.Default.ReplaceTabsWithSpaces,
                    RemoveFieldCodes = Settings.Default.RemoveFieldCodes,
                    RemovePermissions = Settings.Default.RemovePermissions,
                    RemoveProof = Settings.Default.RemoveProof,
                    RemoveSoftHyphens = Settings.Default.RemoveSoftHyphens,
                    RemoveLastRenderedPageBreak = Settings.Default.RemoveLastRenderedPageBreak,
                    RemoveBookmarks = Settings.Default.RemoveBookmarks,
                    RemoveWebHidden = Settings.Default.RemoveWebHidden,
                    NormalizeXml = Settings.Default.NormalizeXml,
                };
                MarkupSimplifier.SimplifyMarkup(doc, settings);
            }
        }
    }
}