// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using OpenXmlPowerTools;

namespace RevisionAccepterExample
{
    internal static class RevisionAccepterExample
    {
        private static void Main()
        {
            DateTime n = DateTime.Now;

            var outputDirectory = new DirectoryInfo(
                $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");

            outputDirectory.Create();

            // Accept all revisions, save result as a new document
            WmlDocument result = RevisionAccepter.AcceptRevisions(new WmlDocument("../../../Source1.docx"));
            result.SaveAs(Path.Combine(outputDirectory.FullName, "Out1.docx"));
        }
    }
}
