// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

internal static class TestTocAdder
{
    private static void Main()
    {
        DateTime n = DateTime.Now;

        var outputDirectory = new DirectoryInfo(
            $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");

        outputDirectory.Create();

        var di2 = new DirectoryInfo("../../");

        foreach (FileInfo file in di2.GetFiles("*.docx"))
        {
            file.CopyTo(Path.Combine(outputDirectory.FullName, file.Name));
        }

        // Inserts a basic TOC before the first paragraph of the document
        using (WordprocessingDocument wdoc =
               WordprocessingDocument.Open(Path.Combine(outputDirectory.FullName, "Test01.docx"), true))
        {
            ReferenceAdder.AddToc(wdoc, "/w:document/w:body/w:p[1]",
                @"TOC \o '1-3' \h \z \u", null, null);
        }

        // Inserts a TOC after the title of the document
        using (WordprocessingDocument wdoc =
               WordprocessingDocument.Open(Path.Combine(outputDirectory.FullName, "Test02.docx"), true))
        {
            ReferenceAdder.AddToc(wdoc, "/w:document/w:body/w:p[2]",
                @"TOC \o '1-3' \h \z \u", null, null);
        }

        // Inserts a TOC with a different title
        using (WordprocessingDocument wdoc =
               WordprocessingDocument.Open(Path.Combine(outputDirectory.FullName, "Test03.docx"), true))
        {
            ReferenceAdder.AddToc(wdoc, "/w:document/w:body/w:p[1]",
                @"TOC \o '1-3' \h \z \u", "Table of Contents", null);
        }

        // Inserts a TOC that includes headings through level 4
        using (WordprocessingDocument wdoc =
               WordprocessingDocument.Open(Path.Combine(outputDirectory.FullName, "Test04.docx"), true))
        {
            ReferenceAdder.AddToc(wdoc, "/w:document/w:body/w:p[1]",
                @"TOC \o '1-4' \h \z \u", null, null);
        }

        // Inserts a table of figures
        using (WordprocessingDocument wdoc =
               WordprocessingDocument.Open(Path.Combine(outputDirectory.FullName, "Test05.docx"), true))
        {
            ReferenceAdder.AddTof(wdoc, "/w:document/w:body/w:p[2]",
                @"TOC \h \z \c ""Figure""", null);
        }

        // Inserts a basic TOC before the first paragraph of the document.
        // Test06.docx does not include a StylesWithEffects part.
        using (WordprocessingDocument wdoc =
               WordprocessingDocument.Open(Path.Combine(outputDirectory.FullName, "Test06.docx"), true))
        {
            ReferenceAdder.AddToc(wdoc, "/w:document/w:body/w:p[1]",
                @"TOC \o '1-3' \h \z \u", null, null);
        }

        // Inserts a TOA before the first paragraph of the document.
        using (WordprocessingDocument wdoc =
               WordprocessingDocument.Open(Path.Combine(outputDirectory.FullName, "Test07.docx"), true))
        {
            ReferenceAdder.AddToa(wdoc, "/w:document/w:body/w:p[2]",
                @"TOA \h \c ""1"" \p", null);
        }
    }
}
