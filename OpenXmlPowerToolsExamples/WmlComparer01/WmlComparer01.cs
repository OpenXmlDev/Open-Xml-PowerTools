// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;

namespace OpenXmlPowerTools
{
    internal static class WmlComparer01
    {
        private static void Main()
        {
            var settings = new WmlComparerSettings();

            WmlDocument result = WmlComparer.Compare(new WmlDocument("../../../Source1.docx"),
                new WmlDocument("../../../Source2.docx"), settings);

            result.SaveAs(Path.Combine(CreateOutputDirectory().FullName, "Compared.docx"));

            List<WmlComparer.WmlComparerRevision> revisions = WmlComparer.GetRevisions(result, settings);

            foreach (WmlComparer.WmlComparerRevision rev in revisions)
            {
                Console.WriteLine("Author: " + rev.Author);
                Console.WriteLine("Revision type: " + rev.RevisionType);
                Console.WriteLine("Revision text: " + rev.Text);
                Console.WriteLine();
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
