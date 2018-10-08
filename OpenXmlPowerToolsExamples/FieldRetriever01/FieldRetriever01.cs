// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

internal class FieldRetriever01
{
    private static void Main()
    {
        DateTime n = DateTime.Now;
        var tempDi = new DirectoryInfo(
            $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");

        tempDi.Create();

        var docWithFooter = new FileInfo("../../../DocWithFooter1.docx");
        var scrubbedDocument = new FileInfo(Path.Combine(tempDi.FullName, "DocWithFooterScrubbed1.docx"));
        File.Copy(docWithFooter.FullName, scrubbedDocument.FullName);
        using (WordprocessingDocument wDoc = WordprocessingDocument.Open(scrubbedDocument.FullName, true))
        {
            ScrubFooter(wDoc, new[] { "PAGE" });
        }

        docWithFooter = new FileInfo("../../../DocWithFooter2.docx");
        scrubbedDocument = new FileInfo(Path.Combine(tempDi.FullName, "DocWithFooterScrubbed2.docx"));
        File.Copy(docWithFooter.FullName, scrubbedDocument.FullName);
        using (WordprocessingDocument wDoc = WordprocessingDocument.Open(scrubbedDocument.FullName, true))
        {
            ScrubFooter(wDoc, new[] { "PAGE", "DATE" });
        }
    }

    private static void ScrubFooter(WordprocessingDocument wDoc, string[] fieldTypesToKeep)
    {
        foreach (FooterPart footer in wDoc.MainDocumentPart!.FooterParts)
        {
            FieldRetriever.AnnotateWithFieldInfo(footer);
            XElement root = footer.GetXDocument().Root;
            RemoveAllButSpecificFields(root, fieldTypesToKeep);
            footer.PutXDocument();
        }
    }

    private static void RemoveAllButSpecificFields(XElement root, string[] fieldTypesToRetain)
    {
        var cachedAnnotationInformation = root.Annotation<Dictionary<int, List<XElement>>>();
        var runsToKeep = new List<XElement>();
        foreach (KeyValuePair<int, List<XElement>> item in cachedAnnotationInformation)
        {
            List<XElement> runsForField = root
                .Descendants()
                .Where(d =>
                {
                    var stack = d.Annotation<Stack<FieldRetriever.FieldElementTypeInfo>>();
                    if (stack == null)
                        return false;
                    if (stack.Any(stackItem => stackItem.Id == item.Key))
                        return true;

                    return false;
                })
                .Select(d => d.AncestorsAndSelf(W.r).FirstOrDefault())
                .GroupAdjacent(o => o)
                .Select(g => g.First())
                .ToList();
            foreach (XElement r in runsForField)
                runsToKeep.Add(r);
        }

        foreach (XElement paragraph in root.Descendants(W.p).ToList())
        {
            if (paragraph.Elements(W.r).Any(r => runsToKeep.Contains(r)))
            {
                paragraph.Elements(W.r)
                    .Where(r => !runsToKeep.Contains(r) &&
                                !r.Elements(W.tab).Any())
                    .Remove();
                paragraph.Elements(W.r)
                    .Where(r => !runsToKeep.Contains(r))
                    .Elements()
                    .Where(rc => rc.Name != W.rPr &&
                                 rc.Name != W.tab)
                    .Remove();
            }
            else
            {
                paragraph.Remove();
            }
        }

        root.Descendants(W.tbl).Remove();
    }
}
