// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace PresentationBuilder02
{
    internal class PresentationBuilder02
    {
        private static void Main()
        {
            DateTime n = DateTime.Now;
            var tempDi = new DirectoryInfo(
                $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");

            tempDi.Create();

            const string presentation = "../../../Presentation1.pptx";
            const string hiddenPresentation = "../../../HiddenPresentation.pptx";

            // First, load both presentations into byte arrays, simulating retrieving presentations from some source
            // such as a SharePoint server
            byte[] baPresentation = File.ReadAllBytes(presentation);
            byte[] baHiddenPresentation = File.ReadAllBytes(hiddenPresentation);

            // Next, replace "thee" with "the" in the main presentation
            var pmlMainPresentation = new PmlDocument("Main.pptx", baPresentation);
            PmlDocument modifiedMainPresentation;

            using (var streamDoc = new OpenXmlMemoryStreamDocument(pmlMainPresentation))
            {
                using (PresentationDocument document = streamDoc.GetPresentationDocument())
                {
                    XDocument pXDoc = document.PresentationPart!.GetXDocument();
                    foreach (XElement slideId in pXDoc.Root!.Elements(P.sldIdLst).Elements(P.sldId))
                    {
                        var slideRelId = (string)slideId.Attribute(R.id);
                        OpenXmlPart slidePart = document.PresentationPart.GetPartById(slideRelId);
                        XDocument slideXDoc = slidePart.GetXDocument();
                        List<XElement> paragraphs = slideXDoc.Descendants(A.p).ToList();
                        OpenXmlRegex.Replace(paragraphs, new Regex("thee"), "the", null);
                        slidePart.SaveXDocument();
                    }
                }

                modifiedMainPresentation = streamDoc.GetModifiedPmlDocument();
            }

            // Combine the two presentations into a single presentation
            var slideSources = new List<SlideSource>
            {
                new(modifiedMainPresentation, 0, 1, true),
                new(new PmlDocument("Hidden.pptx", baHiddenPresentation), true),
                new(modifiedMainPresentation, 1, true)
            };

            PmlDocument combinedPresentation = PresentationBuilder.BuildPresentation(slideSources);

            // Replace <# TRADEMARK #> with AdventureWorks (c)
            PmlDocument modifiedCombinedPresentation;

            using (var streamDoc = new OpenXmlMemoryStreamDocument(combinedPresentation))
            {
                using (PresentationDocument document = streamDoc.GetPresentationDocument())
                {
                    XDocument pXDoc = document.PresentationPart!.GetXDocument();
                    foreach (XElement slideId in pXDoc.Root!.Elements(P.sldIdLst).Elements(P.sldId).Skip(1).Take(1))
                    {
                        var slideRelId = (string)slideId.Attribute(R.id);
                        OpenXmlPart slidePart = document.PresentationPart.GetPartById(slideRelId);
                        XDocument slideXDoc = slidePart.GetXDocument();
                        List<XElement> paragraphs = slideXDoc.Descendants(A.p).ToList();
                        OpenXmlRegex.Replace(paragraphs, new Regex("<# TRADEMARK #>"), "AdventureWorks (c)", null);
                        slidePart.SaveXDocument();
                    }
                }

                modifiedCombinedPresentation = streamDoc.GetModifiedPmlDocument();
            }

            // we now have a PmlDocument (which is essentially a byte array) that can be saved as necessary.
            modifiedCombinedPresentation.SaveAs(Path.Combine(tempDi.FullName, "Modified.pptx"));
        }
    }
}
