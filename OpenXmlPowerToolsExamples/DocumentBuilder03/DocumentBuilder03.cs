// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

internal class Program
{
    private static void Main()
    {
        DateTime n = DateTime.Now;
        var tempDi = new DirectoryInfo(
            $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");

        tempDi.Create();

        var doc1 = new WmlDocument(@"..\..\..\Template.docx");

        using (var mem = new MemoryStream())
        {
            mem.Write(doc1.DocumentByteArray, 0, doc1.DocumentByteArray.Length);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(mem, true))
            {
                XElement root = doc.MainDocumentPart.GetXElement() ?? throw new ArgumentException();

                XElement frontMatterPara = root.Descendants(W.txbxContent).Elements(W.p).First();
                frontMatterPara.ReplaceWith(
                    new XElement(PtOpenXml.Insert,
                        new XAttribute("Id", "Front")));

                XElement tbl = root.Elements(W.body).Elements(W.tbl).First();

                XElement firstCell = tbl.Descendants(W.tr).First().Descendants(W.p).First();
                firstCell.ReplaceWith(
                    new XElement(PtOpenXml.Insert,
                        new XAttribute("Id", "Liz")));

                XElement secondCell = tbl.Descendants(W.tr).Skip(1).First().Descendants(W.p).First();
                secondCell.ReplaceWith(
                    new XElement(PtOpenXml.Insert,
                        new XAttribute("Id", "Eric")));

                doc.MainDocumentPart.PutXElement();
            }

            doc1.DocumentByteArray = mem.ToArray();
        }

        string outFileName = Path.Combine(tempDi.FullName, "Out.docx");
        var sources = new List<Source>
        {
            new(doc1, true),
            new(new WmlDocument(@"..\..\..\Insert-01.docx"), "Liz"),
            new(new WmlDocument(@"..\..\..\Insert-02.docx"), "Eric"),
            new(new WmlDocument(@"..\..\..\FrontMatter.docx"), "Front")
        };

        DocumentBuilder.BuildDocument(sources, outFileName);
    }
}
