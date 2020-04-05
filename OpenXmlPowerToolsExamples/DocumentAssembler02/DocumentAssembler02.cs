// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            var templateDoc = new FileInfo("../../TemplateDocument.docx");
            var dataFile = new FileInfo("../../Data.xml");

            var wmlDoc = new WmlDocument(templateDoc.FullName);
            var data = XElement.Load(dataFile.FullName);
            var wmlAssembledDoc = DocumentAssembler.AssembleDocument(wmlDoc, data, out var templateError);
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