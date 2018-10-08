// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using OpenXmlPowerTools;

namespace DocumentBuilderExample
{
    internal class DocumentBuilderExample
    {
        private static void Main()
        {
            DateTime n = DateTime.Now;
            var tempDi = new DirectoryInfo(
                $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");

            tempDi.Create();

            const string source1 = "../../../Source1.docx";
            const string source2 = "../../../Source2.docx";
            const string source3 = "../../../Source3.docx";

            // Create new document from 10 paragraphs starting at paragraph 5 of Source1.docx
            var sources = new List<Source>
            {
                new(new WmlDocument(source1), 5, 10, true)
            };

            DocumentBuilder.BuildDocument(sources, Path.Combine(tempDi.FullName, "Out1.docx"));

            // Create new document from paragraph 1, and paragraphs 5 through end of Source3.docx.
            // This effectively 'deletes' paragraphs 2-4
            sources = new List<Source>
            {
                new(new WmlDocument(source3), 0, 1, false),
                new(new WmlDocument(source3), 4, false)
            };

            DocumentBuilder.BuildDocument(sources, Path.Combine(tempDi.FullName, "Out2.docx"));

            // Create a new document that consists of the entirety of Source1.docx and Source2.docx.  Use
            // the section information (headings and footers) from source1.
            sources = new List<Source>
            {
                new(new WmlDocument(source1), true),
                new(new WmlDocument(source2), false)
            };

            DocumentBuilder.BuildDocument(sources, Path.Combine(tempDi.FullName, "Out3.docx"));

            // Create a new document that consists of the entirety of Source1.docx and Source2.docx.  Use
            // the section information (headings and footers) from source2.
            sources = new List<Source>
            {
                new(new WmlDocument(source1), false),
                new(new WmlDocument(source2), true)
            };

            DocumentBuilder.BuildDocument(sources, Path.Combine(tempDi.FullName, "Out4.docx"));

            // Create a new document that consists of the first 5 paragraphs of Source1.docx and the first
            // five paragraphs of Source2.docx.  This example returns a new WmlDocument, when you then can
            // serialize to a SharePoint document library, or use in some other interesting scenario.
            sources = new List<Source>
            {
                new(new WmlDocument(source1), 0, 5, false),
                new(new WmlDocument(source2), 0, 5, true)
            };

            WmlDocument out5 = DocumentBuilder.BuildDocument(sources);
            out5.SaveAs(Path.Combine(tempDi.FullName, "Out5.docx"));
        }
    }
}
