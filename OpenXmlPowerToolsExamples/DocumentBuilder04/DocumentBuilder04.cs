// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace ExampleDocumentBuilder04
{
    internal class ContentControlsExample
    {
        private static void Main()
        {
            DateTime n = DateTime.Now;
            var tempDi = new DirectoryInfo(
                $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");

            tempDi.Create();

            var solarSystemDoc = new WmlDocument("../../../solar-system.docx");
            using var streamDoc = new OpenXmlMemoryStreamDocument(solarSystemDoc);
            using WordprocessingDocument solarSystem = streamDoc.GetWordprocessingDocument();
            XElement root = solarSystem.MainDocumentPart.GetXElement();

            // Get children elements of the <w:body> element, ignoring the w:sectPr element.
            IEnumerable<XElement> q1 = root
                .Elements(W.body)
                .Elements()
                .Where(e => e.Name != W.sectPr);

            // Project collection of tuples containing element and type.
            var q2 = q1.Select(e => new
            {
                Element = e,
                KeyForGroupAdjacent = e.Name.LocalName switch
                {
                    nameof(W.sdt) => e.Element(W.sdtPr)?.Element(W.tag)?.Attribute(W.val)?.Value,
                    _ => ".NonContentControl"
                }
            });

            // Group by key.
            var q3 = q2.GroupAdjacent(e => e.KeyForGroupAdjacent).ToList();

            // Validate existence of files referenced in content controls.
            foreach (var tagValue in q3.Where(g => g.Key != ".NonContentControl"))
            {
                string filename = "../../../" + tagValue.Key + ".docx";
                var fi = new FileInfo(filename);
                if (fi.Exists) continue;

                Console.WriteLine($"{filename} doesn't exist.");
                Environment.Exit(0);
            }

            // Project collection with opened WordProcessingDocument.
            var q4 = q3.Select(g => new
            {
                Group = g,
                Document = g.Key != ".NonContentControl" ? new WmlDocument("../../../" + g.Key + ".docx") : solarSystemDoc
            });

            // Project collection of OpenXml.PowerTools.Source.
            List<Source> sources = q4
                .Select(g => g.Group.Key == ".NonContentControl"
                    ? new Source(g.Document, g.Group.First().Element.ElementsBeforeSelf().Count(), g.Group.Count(), false)
                    : new Source(g.Document, false))
                .ToList();

            DocumentBuilder.BuildDocument(sources, Path.Combine(tempDi.FullName, "solar-system-new.docx"));
        }
    }
}
