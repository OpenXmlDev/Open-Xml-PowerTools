// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    internal class Program
    {
        private static void Main()
        {
            DateTime n = DateTime.Now;
            var tempDi = new DirectoryInfo(
                $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");

            tempDi.Create();

            var templateDoc = new FileInfo("../../../TemplateDocument.docx");
            var dataFile = new FileInfo("../../../Data.xml");

            var wmlDoc = new WmlDocument(templateDoc.FullName);
            XElement data = XElement.Load(dataFile.FullName);
            WmlDocument wmlAssembledDoc = DocumentAssembler.AssembleDocument(wmlDoc, data, out bool templateError);
            if (templateError)
            {
                Console.WriteLine("Errors in template.");
                Console.WriteLine("See AssembledDoc.docx to determine the errors in the template.");
            }

            var assembledDoc = new FileInfo(Path.Combine(tempDi.FullName, "AssembledDoc.docx"));
            wmlAssembledDoc.SaveAs(assembledDoc.FullName);
        }
    }
}
