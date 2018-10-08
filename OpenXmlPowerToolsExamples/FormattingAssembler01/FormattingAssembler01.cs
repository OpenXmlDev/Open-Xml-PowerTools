// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace FormattingAssembler01
{
    internal class FormattingAssembler01
    {
        private static void Main()
        {
            DateTime n = DateTime.Now;
            var tempDi = new DirectoryInfo(
                $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");

            tempDi.Create();

            var di = new DirectoryInfo("../../../");
            foreach (FileInfo file in di.GetFiles("*.docx"))
            {
                Console.WriteLine(file.Name);
                var newFile = new FileInfo(Path.Combine(tempDi.FullName, file.Name.Replace(".docx", "out.docx")));
                File.Copy(file.FullName, newFile.FullName);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(newFile.FullName, true))
                {
                    var settings = new FormattingAssemblerSettings
                    {
                        ClearStyles = true,
                        RemoveStyleNamesFromParagraphAndRunProperties = true,
                        CreateHtmlConverterAnnotationAttributes = true,
                        OrderElementsPerStandard = true,
                        RestrictToSupportedLanguages = true,
                        RestrictToSupportedNumberingFormats = true
                    };
                    FormattingAssembler.AssembleFormatting(wDoc, settings);
                }
            }
        }
    }
}
