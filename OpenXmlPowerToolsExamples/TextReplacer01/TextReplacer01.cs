// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

internal static class TestPmlTextReplacer
{
    private static void Main()
    {
        DirectoryInfo outputDirectory = CreateOutputDirectory();

        File.Copy("../../../Test01.pptx", Path.Combine(outputDirectory.FullName, "Test01out.pptx"));

        using (PresentationDocument pDoc =
               PresentationDocument.Open(Path.Combine(outputDirectory.FullName, "Test01out.pptx"), true))
        {
            TextReplacer.SearchAndReplace(pDoc, "Hello", "Goodbye", true);
        }

        File.Copy("../../../Test02.pptx", Path.Combine(outputDirectory.FullName, "Test02out.pptx"));

        using (PresentationDocument pDoc =
               PresentationDocument.Open(Path.Combine(outputDirectory.FullName, "Test02out.pptx"), true))
        {
            TextReplacer.SearchAndReplace(pDoc, "Hello", "Goodbye", true);
        }

        File.Copy("../../../Test03.pptx", Path.Combine(outputDirectory.FullName, "Test03out.pptx"));

        using (PresentationDocument pDoc =
               PresentationDocument.Open(Path.Combine(outputDirectory.FullName, "Test03out.pptx"), true))
        {
            TextReplacer.SearchAndReplace(pDoc, "Hello", "Goodbye", false);
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
