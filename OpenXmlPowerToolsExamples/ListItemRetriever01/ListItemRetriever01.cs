// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

/***************************************************************************

Copyright (c) Microsoft Corporation 2014.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license
can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

***************************************************************************/

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

internal class ListItemRetriever01
{
    private static void Main()
    {
        DateTime n = DateTime.Now;
        var tempDi = new DirectoryInfo(
            $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");

        tempDi.Create();

        using WordprocessingDocument wDoc = WordprocessingDocument.Open("../../../NumberedListTest.docx", false);
        const int abstractNumId = 0;
        XElement xml = ConvertDocToXml(wDoc, abstractNumId);
        Console.WriteLine(xml);
        xml.Save(Path.Combine(tempDi.FullName, "Out.xml"));
    }

    private static XElement ConvertDocToXml(WordprocessingDocument wDoc, int abstractNumId)
    {
        XDocument xd = wDoc.MainDocumentPart.GetXDocument();

        // First, call RetrieveListItem so that all paragraphs are initialized with ListItemInfo
        XElement firstParagraph = xd.Descendants(W.p).First();
        string listItem = ListItemRetriever.RetrieveListItem(wDoc, firstParagraph);

        var xml = new XElement("Root");
        var current = new Stack<XmlStackItem>();
        current.Push(
            new XmlStackItem
            {
                Element = xml,
                LevelNumbers = new int[] { }
            });
        foreach (XElement paragraph in xd.Descendants(W.p))
        {
            // The following does not take into account documents that have tracked revisions.
            // As necessary, call RevisionAccepter.AcceptRevisions before converting to XML.
            string text = paragraph.Descendants(W.t).Select(t => (string)t).StringConcatenate();
            var lii =
                paragraph.Annotation<ListItemRetriever.ListItemInfo>();
            if (lii.IsListItem && lii.AbstractNumId == abstractNumId)
            {
                var levelNums =
                    paragraph.Annotation<ListItemRetriever.LevelNumbers>();
                if (levelNums.LevelNumbersArray.Length == current.Peek().LevelNumbers.Length)
                {
                    current.Pop();
                    int[] levelNumsForThisIndent = levelNums.LevelNumbersArray;
                    string levelText = levelNums
                        .LevelNumbersArray
                        .Select(l => l + ".")
                        .StringConcatenate()
                        .TrimEnd('.');
                    var newCurrentElement = new XElement("Indent",
                        new XAttribute("Level", levelText));
                    current.Peek().Element.Add(newCurrentElement);
                    current.Push(
                        new XmlStackItem
                        {
                            Element = newCurrentElement,
                            LevelNumbers = levelNumsForThisIndent
                        });
                    current.Peek().Element.Add(new XElement("Heading", text));
                }
                else if (levelNums.LevelNumbersArray.Length > current.Peek().LevelNumbers.Length)
                {
                    for (int i = current.Peek().LevelNumbers.Length;
                         i < levelNums.LevelNumbersArray.Length;
                         i++)
                    {
                        int[] levelNumsForThisIndent = levelNums
                            .LevelNumbersArray
                            .Take(i + 1)
                            .ToArray();
                        string levelText = levelNums
                            .LevelNumbersArray
                            .Select(l => l + ".")
                            .StringConcatenate()
                            .TrimEnd('.');
                        var newCurrentElement = new XElement("Indent",
                            new XAttribute("Level", levelText));
                        current.Peek().Element.Add(newCurrentElement);
                        current.Push(
                            new XmlStackItem
                            {
                                Element = newCurrentElement,
                                LevelNumbers = levelNumsForThisIndent
                            });
                        current.Peek().Element.Add(new XElement("Heading", text));
                    }
                }
                else if (levelNums.LevelNumbersArray.Length < current.Peek().LevelNumbers.Length)
                {
                    for (int i = current.Peek().LevelNumbers.Length;
                         i > levelNums.LevelNumbersArray.Length;
                         i--)
                        current.Pop();
                    current.Pop();
                    int[] levelNumsForThisIndent = levelNums.LevelNumbersArray;
                    string levelText = levelNums
                        .LevelNumbersArray
                        .Select(l => l + ".")
                        .StringConcatenate()
                        .TrimEnd('.');
                    var newCurrentElement = new XElement("Indent",
                        new XAttribute("Level", levelText));
                    current.Peek().Element.Add(newCurrentElement);
                    current.Push(
                        new XmlStackItem
                        {
                            Element = newCurrentElement,
                            LevelNumbers = levelNumsForThisIndent
                        });
                    current.Peek().Element.Add(new XElement("Heading", text));
                }
            }
            else
            {
                current.Peek().Element.Add(new XElement("Paragraph", text));
            }
        }

        return xml;
    }

    private class XmlStackItem
    {
        public XElement Element;
        public int[] LevelNumbers;
    }
}
