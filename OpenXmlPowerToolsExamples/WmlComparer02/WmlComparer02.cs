// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace OpenXmlPowerTools
{
    internal class WmlComparer02
    {
        private static void Main(string[] args)
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            var originalWml = new WmlDocument("../../Original.docx");
            var revisedDocumentInfoList = new List<WmlRevisedDocumentInfo>()
            {
                new WmlRevisedDocumentInfo()
                {
                    RevisedDocument = new WmlDocument("../../RevisedByBob.docx"),
                    Revisor = "Bob",
                    Color = Color.LightBlue,
                },
                new WmlRevisedDocumentInfo()
                {
                    RevisedDocument = new WmlDocument("../../RevisedByMary.docx"),
                    Revisor = "Mary",
                    Color = Color.LightYellow,
                },
            };
            var settings = new WmlComparerSettings();
            var consolidatedWml = WmlComparer.Consolidate(
                originalWml,
                revisedDocumentInfoList,
                settings);
            consolidatedWml.SaveAs(Path.Combine(tempDi.FullName, "Consolidated.docx"));
        }
    }
}