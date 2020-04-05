// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

internal class Program
{
    private static void Main(string[] args)
    {
        var n = DateTime.Now;
        var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
        tempDi.Create();

        var doc1 = new WmlDocument(@"..\..\Template.docx");
        using (var mem = new MemoryStream())
        {
            mem.Write(doc1.DocumentByteArray, 0, doc1.DocumentByteArray.Length);
            using (var doc = WordprocessingDocument.Open(mem, true))
            {
                var xDoc = doc.MainDocumentPart.GetXDocument();
                var frontMatterPara = xDoc.Root.Descendants(W.txbxContent).Elements(W.p).FirstOrDefault();
                frontMatterPara.ReplaceWith(
                    new XElement(PtOpenXml.Insert,
                        new XAttribute("Id", "Front")));
                var tbl = xDoc.Root.Element(W.body).Elements(W.tbl).FirstOrDefault();
                var firstCell = tbl.Descendants(W.tr).First().Descendants(W.p).First();
                firstCell.ReplaceWith(
                    new XElement(PtOpenXml.Insert,
                        new XAttribute("Id", "Liz")));
                var secondCell = tbl.Descendants(W.tr).Skip(1).First().Descendants(W.p).First();
                secondCell.ReplaceWith(
                    new XElement(PtOpenXml.Insert,
                        new XAttribute("Id", "Eric")));
                doc.MainDocumentPart.PutXDocument();
            }
            doc1.DocumentByteArray = mem.ToArray();
        }

        var outFileName = Path.Combine(tempDi.FullName, "Out.docx");
        var sources = new List<Source>()
            {
                new Source(doc1, true),
                new Source(new WmlDocument(@"..\..\Insert-01.docx"), "Liz"),
                new Source(new WmlDocument(@"..\..\Insert-02.docx"), "Eric"),
                new Source(new WmlDocument(@"..\..\FrontMatter.docx"), "Front"),
            };
        DocumentBuilder.BuildDocument(sources, outFileName);
    }
}