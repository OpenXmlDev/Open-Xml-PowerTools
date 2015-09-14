using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenXmlPowerTools;

namespace DocumentBuilderExample
{
    class DocumentBuilderExample
    {
        static void Main(string[] args)
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            string source1 = "../../Source1.docx";
            string source2 = "../../Source2.docx";
            string source3 = "../../Source3.docx";
            List<Source> sources = null;

            // Create new document from 10 paragraphs starting at paragraph 5 of Source1.docx
            sources = new List<Source>()
            {
                new Source(new WmlDocument(source1), 5, 10, true),
            };
            DocumentBuilder.BuildDocument(sources, Path.Combine(tempDi.FullName, "Out1.docx"));

            // Create new document from paragraph 1, and paragraphs 5 through end of Source3.docx.
            // This effectively 'deletes' paragraphs 2-4
            sources = new List<Source>()
            {
                new Source(new WmlDocument(source3), 0, 1, false),
                new Source(new WmlDocument(source3), 4, false),
            };
            DocumentBuilder.BuildDocument(sources, Path.Combine(tempDi.FullName, "Out2.docx"));

            // Create a new document that consists of the entirety of Source1.docx and Source2.docx.  Use
            // the section information (headings and footers) from source1.
            sources = new List<Source>()
            {
                new Source(new WmlDocument(source1), true),
                new Source(new WmlDocument(source2), false),
            };
            DocumentBuilder.BuildDocument(sources, Path.Combine(tempDi.FullName, "Out3.docx"));

            // Create a new document that consists of the entirety of Source1.docx and Source2.docx.  Use
            // the section information (headings and footers) from source2.
            sources = new List<Source>()
            {
                new Source(new WmlDocument(source1), false),
                new Source(new WmlDocument(source2), true),
            };
            DocumentBuilder.BuildDocument(sources, Path.Combine(tempDi.FullName, "Out4.docx"));

            // Create a new document that consists of the first 5 paragraphs of Source1.docx and the first
            // five paragraphs of Source2.docx.  This example returns a new WmlDocument, when you then can
            // serialize to a SharePoint document library, or use in some other interesting scenario.
            sources = new List<Source>()
            {
                new Source(new WmlDocument(source1), 0, 5, false),
                new Source(new WmlDocument(source2), 0, 5, true),
            };
            WmlDocument out5 = DocumentBuilder.BuildDocument(sources);
            out5.SaveAs(Path.Combine(tempDi.FullName, "Out5.docx"));  // save it to the file system, but we could just as easily done something
                                                                      // else with it.
        }
    }
}
