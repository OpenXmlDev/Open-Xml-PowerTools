// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System;
using System.IO;

namespace FormattingAssembler01
{
    internal class FormattingAssembler01
    {
        private static void Main(string[] args)
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            var di = new DirectoryInfo("../../");
            foreach (var file in di.GetFiles("*.docx"))
            {
                Console.WriteLine(file.Name);
                var newFile = new FileInfo(Path.Combine(tempDi.FullName, file.Name.Replace(".docx", "out.docx")));
                File.Copy(file.FullName, newFile.FullName);
                using (var wDoc = WordprocessingDocument.Open(newFile.FullName, true))
                {
                    var settings = new FormattingAssemblerSettings()
                    {
                        ClearStyles = true,
                        RemoveStyleNamesFromParagraphAndRunProperties = true,
                        CreateHtmlConverterAnnotationAttributes = true,
                        OrderElementsPerStandard = true,
                        RestrictToSupportedLanguages = true,
                        RestrictToSupportedNumberingFormats = true,
                    };
                    FormattingAssembler.AssembleFormatting(wDoc, settings);
                }
            }
        }
    }
}