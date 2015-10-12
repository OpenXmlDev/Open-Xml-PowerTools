using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

class TestTocAdder
{
    static void Main(string[] args)
    {
        var n = DateTime.Now;
        var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
        tempDi.Create();

        DirectoryInfo di2 = new DirectoryInfo("../../");
        foreach (var file in di2.GetFiles("*.docx"))
            file.CopyTo(Path.Combine(tempDi.FullName, file.Name));

        List<string> filesToProcess = new List<string>();

        // Inserts a basic TOC before the first paragraph of the document
        using (WordprocessingDocument wdoc =
            WordprocessingDocument.Open(Path.Combine(tempDi.FullName, "Test01.docx"), true))
        {
            ReferenceAdder.AddToc(wdoc, "/w:document/w:body/w:p[1]",
                @"TOC \o '1-3' \h \z \u", null, null);
        }

        // Inserts a TOC after the title of the document
        using (WordprocessingDocument wdoc =
            WordprocessingDocument.Open(Path.Combine(tempDi.FullName, "Test02.docx"), true))
        {
            ReferenceAdder.AddToc(wdoc, "/w:document/w:body/w:p[2]",
                @"TOC \o '1-3' \h \z \u", null, null);
        }

        // Inserts a TOC with a different title
        using (WordprocessingDocument wdoc =
            WordprocessingDocument.Open(Path.Combine(tempDi.FullName, "Test03.docx"), true))
        {
            ReferenceAdder.AddToc(wdoc, "/w:document/w:body/w:p[1]",
                @"TOC \o '1-3' \h \z \u", "Table of Contents", null);
        }

        // Inserts a TOC that includes headings through level 4
        using (WordprocessingDocument wdoc =
            WordprocessingDocument.Open(Path.Combine(tempDi.FullName, "Test04.docx"), true))
        {
            ReferenceAdder.AddToc(wdoc, "/w:document/w:body/w:p[1]",
                @"TOC \o '1-4' \h \z \u", null, null);
        }

        // Inserts a table of figures
        using (WordprocessingDocument wdoc =
            WordprocessingDocument.Open(Path.Combine(tempDi.FullName, "Test05.docx"), true))
        {
            ReferenceAdder.AddTof(wdoc, "/w:document/w:body/w:p[2]",
                @"TOC \h \z \c ""Figure""", null);
        }

        // Inserts a basic TOC before the first paragraph of the document.
        // Test06.docx does not include a StylesWithEffects part.
        using (WordprocessingDocument wdoc =
            WordprocessingDocument.Open(Path.Combine(tempDi.FullName, "Test06.docx"), true))
        {
            ReferenceAdder.AddToc(wdoc, "/w:document/w:body/w:p[1]",
                @"TOC \o '1-3' \h \z \u", null, null);
        }

        // Inserts a TOA before the first paragraph of the document.
        using (WordprocessingDocument wdoc =
            WordprocessingDocument.Open(Path.Combine(tempDi.FullName, "Test07.docx"), true))
        {
            ReferenceAdder.AddToa(wdoc, "/w:document/w:body/w:p[2]",
                @"TOA \h \c ""1"" \p", null);
        }
    }
}
