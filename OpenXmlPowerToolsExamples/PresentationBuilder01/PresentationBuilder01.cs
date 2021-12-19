// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using OpenXmlPowerTools;

namespace ExamplePresentatonBuilder01
{
    internal static class ExamplePresentationBuilder01
    {
        private static void Main()
        {
            DateTime n = DateTime.Now;

            var outputDirectory = new DirectoryInfo(
                $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");

            outputDirectory.Create();

            const string source1 = "../../../Contoso.pptx";
            const string source2 = "../../../Companies.pptx";
            const string source3 = "../../../Customer Content.pptx";
            const string source4 = "../../../Presentation One.pptx";
            const string source5 = "../../../Presentation Two.pptx";
            const string source6 = "../../../Presentation Three.pptx";
            const string contoso1 = "../../../Contoso One.pptx";
            const string contoso2 = "../../../Contoso Two.pptx";
            const string contoso3 = "../../../Contoso Three.pptx";

            var sourceDoc = new PmlDocument(source1);

            var sources = new List<SlideSource>
            {
                new(sourceDoc, 0, 1, false), // Title
                new(sourceDoc, 1, 1, false), // First intro (of 3)
                new(sourceDoc, 4, 2, false), // Sales bios
                new(sourceDoc, 9, 3, false), // Content slides
                new(sourceDoc, 13, 1, false), // Closing summary
            };

            PresentationBuilder.BuildPresentation(sources, Path.Combine(outputDirectory.FullName, "Out1.pptx"));

            sources = new List<SlideSource>
            {
                new(new PmlDocument(source2), 2, 1, true), // Choose company
                new(new PmlDocument(source3), false), // Content
            };

            PresentationBuilder.BuildPresentation(sources, Path.Combine(outputDirectory.FullName, "Out2.pptx"));

            sources = new List<SlideSource>
            {
                new(new PmlDocument(source4), true),
                new(new PmlDocument(source5), true),
                new(new PmlDocument(source6), true),
            };

            PresentationBuilder.BuildPresentation(sources, Path.Combine(outputDirectory.FullName, "Out3.pptx"));

            sources = new List<SlideSource>
            {
                new(new PmlDocument(contoso1), true),
                new(new PmlDocument(contoso2), true),
                new(new PmlDocument(contoso3), true),
            };

            PresentationBuilder.BuildPresentation(sources, Path.Combine(outputDirectory.FullName, "Out4.pptx"));
        }
    }
}
