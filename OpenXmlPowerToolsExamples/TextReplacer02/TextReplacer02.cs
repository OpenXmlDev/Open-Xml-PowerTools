// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    internal static class Program
    {
        private static void Main()
        {
            var inputDirectory = new DirectoryInfo("../../../");
            DirectoryInfo outputDirectory = CreateOutputDirectory();

            foreach (FileInfo file in inputDirectory.GetFiles("*.docx"))
            {
                file.CopyTo(Path.Combine(outputDirectory.FullName, file.Name));
            }

            using (WordprocessingDocument doc =
                   WordprocessingDocument.Open(Path.Combine(outputDirectory.FullName, "Test01.docx"), true))
            {
                TextReplacer.SearchAndReplace(doc, "the", "this", false);
            }

            try
            {
                using WordprocessingDocument doc =
                    WordprocessingDocument.Open(Path.Combine(outputDirectory.FullName, "Test02.docx"), true);

                TextReplacer.SearchAndReplace(doc, "the", "this", false);
            }
            catch (Exception)
            {
                // ignored
            }

            try
            {
                using WordprocessingDocument doc =
                    WordprocessingDocument.Open(Path.Combine(outputDirectory.FullName, "Test03.docx"), true);

                TextReplacer.SearchAndReplace(doc, "the", "this", false);
            }
            catch (Exception)
            {
                // ignored
            }

            using (WordprocessingDocument doc =
                   WordprocessingDocument.Open(Path.Combine(outputDirectory.FullName, "Test04.docx"), true))
            {
                TextReplacer.SearchAndReplace(doc, "the", "this", true);
            }

            using (WordprocessingDocument doc =
                   WordprocessingDocument.Open(Path.Combine(outputDirectory.FullName, "Test05.docx"), true))
            {
                TextReplacer.SearchAndReplace(doc, "is on", "is above", true);
            }

            using (WordprocessingDocument doc =
                   WordprocessingDocument.Open(Path.Combine(outputDirectory.FullName, "Test06.docx"), true))
            {
                TextReplacer.SearchAndReplace(doc, "the", "this", false);
            }

            using (WordprocessingDocument doc =
                   WordprocessingDocument.Open(Path.Combine(outputDirectory.FullName, "Test07.docx"), true))
            {
                TextReplacer.SearchAndReplace(doc, "the", "this", true);
            }

            using (WordprocessingDocument doc =
                   WordprocessingDocument.Open(Path.Combine(outputDirectory.FullName, "Test08.docx"), true))
            {
                TextReplacer.SearchAndReplace(doc, "the", "this", true);
            }

            using (WordprocessingDocument doc =
                   WordprocessingDocument.Open(Path.Combine(outputDirectory.FullName, "Test09.docx"), true))
            {
                TextReplacer.SearchAndReplace(doc, "===== Replace this text =====", "***zzz***", true);
            }
        }

        private static DirectoryInfo CreateOutputDirectory()
        {
            DateTime n = DateTime.Now;

            var outputDirectory = new DirectoryInfo(
                $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");

            outputDirectory.Create();

            return outputDirectory;
        }
    }
}
