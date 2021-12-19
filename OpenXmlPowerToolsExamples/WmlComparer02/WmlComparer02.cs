// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace OpenXmlPowerTools
{
    internal static class WmlComparer02
    {
        private static void Main()
        {
            var originalWml = new WmlDocument("../../../Original.docx");

            var revisedDocumentInfoList = new List<WmlRevisedDocumentInfo>
            {
                new()
                {
                    RevisedDocument = new WmlDocument("../../../RevisedByBob.docx"),
                    Revisor = "Bob",
                    Color = Color.LightBlue,
                },
                new()
                {
                    RevisedDocument = new WmlDocument("../../../RevisedByMary.docx"),
                    Revisor = "Mary",
                    Color = Color.LightYellow,
                },
            };

            var settings = new WmlComparerSettings();
            WmlDocument consolidatedWml = WmlComparer.Consolidate(originalWml, revisedDocumentInfoList, settings);
            consolidatedWml.SaveAs(Path.Combine(CreateOutputDirectory().FullName, "Consolidated.docx"));
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
