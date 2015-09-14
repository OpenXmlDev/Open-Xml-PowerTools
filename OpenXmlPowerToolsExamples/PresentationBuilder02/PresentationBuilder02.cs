using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace PresentationBuilder02
{
    class PresentationBuilder02
    {
        static void Main(string[] args)
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            string presentation = "../../Presentation1.pptx";
            string hiddenPresentation = "../../HiddenPresentation.pptx";

            // First, load both presentations into byte arrays, simulating retrieving presentations from some source
            // such as a SharePoint server
            var baPresentation = File.ReadAllBytes(presentation);
            var baHiddenPresentation = File.ReadAllBytes(hiddenPresentation);

            // Next, replace "thee" with "the" in the main presentation
            var pmlMainPresentation = new PmlDocument("Main.pptx", baPresentation);
            PmlDocument modifiedMainPresentation = null;
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(pmlMainPresentation))
            {
                using (PresentationDocument document = streamDoc.GetPresentationDocument())
                {
                    var pXDoc = document.PresentationPart.GetXDocument();
                    foreach (var slideId in pXDoc.Root.Elements(P.sldIdLst).Elements(P.sldId))
                    {
                        var slideRelId = (string)slideId.Attribute(R.id);
                        var slidePart = document.PresentationPart.GetPartById(slideRelId);
                        var slideXDoc = slidePart.GetXDocument();
                        var paragraphs = slideXDoc.Descendants(A.p).ToList();
                        OpenXmlRegex.Replace(paragraphs, new Regex("thee"), "the", null);
                        slidePart.PutXDocument();
                    }
                }
                modifiedMainPresentation = streamDoc.GetModifiedPmlDocument();
            }

            // Combine the two presentations into a single presentation
            var slideSources = new List<SlideSource>() {
                new SlideSource(modifiedMainPresentation, 0, 1, true),
                new SlideSource(new PmlDocument("Hidden.pptx", baHiddenPresentation), true),
                new SlideSource(modifiedMainPresentation, 1, true),
            };
            PmlDocument combinedPresentation = PresentationBuilder.BuildPresentation(slideSources);

            // Replace <# TRADEMARK #> with AdventureWorks (c)
            PmlDocument modifiedCombinedPresentation = null;
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(combinedPresentation))
            {
                using (PresentationDocument document = streamDoc.GetPresentationDocument())
                {
                    var pXDoc = document.PresentationPart.GetXDocument();
                    foreach (var slideId in pXDoc.Root.Elements(P.sldIdLst).Elements(P.sldId).Skip(1).Take(1))
                    {
                        var slideRelId = (string)slideId.Attribute(R.id);
                        var slidePart = document.PresentationPart.GetPartById(slideRelId);
                        var slideXDoc = slidePart.GetXDocument();
                        var paragraphs = slideXDoc.Descendants(A.p).ToList();
                        OpenXmlRegex.Replace(paragraphs, new Regex("<# TRADEMARK #>"), "AdventureWorks (c)", null);
                        slidePart.PutXDocument();
                    }
                }
                modifiedCombinedPresentation = streamDoc.GetModifiedPmlDocument();
            }

            // we now have a PmlDocument (which is essentially a byte array) that can be saved as necessary.
            modifiedCombinedPresentation.SaveAs(Path.Combine(tempDi.FullName, "Modified.pptx"));
        }
    }
}
