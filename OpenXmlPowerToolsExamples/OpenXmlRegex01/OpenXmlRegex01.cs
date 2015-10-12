using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

class OpenXmlRegexExample
{
    static void Main(string[] args)
    {
        var n = DateTime.Now;
        var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
        tempDi.Create();

        var sourceDoc = new FileInfo("../../TestDocument.docx");
        var newDoc = new FileInfo(Path.Combine(tempDi.FullName, "Modified.docx"));
        File.Copy(sourceDoc.FullName, newDoc.FullName);
        using (WordprocessingDocument wDoc = WordprocessingDocument.Open(newDoc.FullName, true))
        {
            int count;
            var xDoc = wDoc.MainDocumentPart.GetXDocument();
            Regex regex;
            IEnumerable<XElement> content;

            // Match content (paragraph 1)
            content = xDoc.Descendants(W.p).Take(1);
            regex = new Regex("Video");
            count = OpenXmlRegex.Match(content, regex);
            Console.WriteLine("Example #1 Count: {0}", count);

            // Match content, case insensitive (paragraph 1)
            content = xDoc.Descendants(W.p).Take(1);
            regex = new Regex("video", RegexOptions.IgnoreCase);
            count = OpenXmlRegex.Match(content, regex);
            Console.WriteLine("Example #2 Count: {0}", count);

            // Match content, with callback (paragraph 1)
            content = xDoc.Descendants(W.p).Take(1);
            regex = new Regex("video", RegexOptions.IgnoreCase);
            count = OpenXmlRegex.Match(content, regex, (element, match) =>
                Console.WriteLine("Example #3 Found value: >{0}<", match.Value));

            // Replace content, beginning of paragraph (paragraph 2)
            content = xDoc.Descendants(W.p).Skip(1).Take(1);
            regex = new Regex("^Video provides");
            count = OpenXmlRegex.Replace(content, regex, "Audio gives", null);
            Console.WriteLine("Example #4 Replaced: {0}", count);

            // Replace content, middle of paragraph (paragraph 3)
            content = xDoc.Descendants(W.p).Skip(2).Take(1);
            regex = new Regex("powerful");
            count = OpenXmlRegex.Replace(content, regex, "good", null);
            Console.WriteLine("Example #5 Replaced: {0}", count);

            // Replace content, end of paragraph (paragraph 4)
            content = xDoc.Descendants(W.p).Skip(3).Take(1);
            regex = new Regex(" [a-z.]*$");
            count = OpenXmlRegex.Replace(content, regex, " super good point!", null);
            Console.WriteLine("Example #6 Replaced: {0}", count);

            // Delete content, beginning of paragraph (paragraph 5)
            content = xDoc.Descendants(W.p).Skip(4).Take(1);
            regex = new Regex("^Video provides");
            count = OpenXmlRegex.Replace(content, regex, "", null);
            Console.WriteLine("Example #7 Deleted: {0}", count);

            // Delete content, middle of paragraph (paragraph 6)
            content = xDoc.Descendants(W.p).Skip(5).Take(1);
            regex = new Regex("powerful ");
            count = OpenXmlRegex.Replace(content, regex, "", null);
            Console.WriteLine("Example #8 Deleted: {0}", count);

            // Delete content, end of paragraph (paragraph 7)
            content = xDoc.Descendants(W.p).Skip(6).Take(1);
            regex = new Regex("[.]$");
            count = OpenXmlRegex.Replace(content, regex, "", null);
            Console.WriteLine("Example #9 Deleted: {0}", count);

            // Replace content in inserted text, same author (paragraph 8)
            content = xDoc.Descendants(W.p).Skip(7).Take(1);
            regex = new Regex("Video");
            count = OpenXmlRegex.Replace(content, regex, "Audio", null, true, "Eric White");
            Console.WriteLine("Example #10 Deleted: {0}", count);

            // Delete content in inserted text, same author (paragraph 9)
            content = xDoc.Descendants(W.p).Skip(8).Take(1);
            regex = new Regex("powerful ");
            count = OpenXmlRegex.Replace(content, regex, "", null, true, "Eric White");
            Console.WriteLine("Example #11 Deleted: {0}", count);

            // Replace content partially in inserted text, same author (paragraph 10)
            content = xDoc.Descendants(W.p).Skip(9).Take(1);
            regex = new Regex("Video provides ");
            count = OpenXmlRegex.Replace(content, regex, "Audio gives ", null, true, "Eric White");
            Console.WriteLine("Example #12 Replaced: {0}", count);

            // Delete content partially in inserted text, same author (paragraph 11)
            content = xDoc.Descendants(W.p).Skip(10).Take(1);
            regex = new Regex(" to help you prove your point");
            count = OpenXmlRegex.Replace(content, regex, "", null, true, "Eric White");
            Console.WriteLine("Example #13 Deleted: {0}", count);

            // Replace content in inserted text, different author (paragraph 12)
            content = xDoc.Descendants(W.p).Skip(11).Take(1);
            regex = new Regex("Video");
            count = OpenXmlRegex.Replace(content, regex, "Audio", null, true, "John Doe");
            Console.WriteLine("Example #14 Deleted: {0}", count);

            // Delete content in inserted text, different author (paragraph 13)
            content = xDoc.Descendants(W.p).Skip(12).Take(1);
            regex = new Regex("powerful ");
            count = OpenXmlRegex.Replace(content, regex, "", null, true, "John Doe");
            Console.WriteLine("Example #15 Deleted: {0}", count);

            // Replace content partially in inserted text, different author (paragraph 14)
            content = xDoc.Descendants(W.p).Skip(13).Take(1);
            regex = new Regex("Video provides ");
            count = OpenXmlRegex.Replace(content, regex, "Audio gives ", null, true, "John Doe");
            Console.WriteLine("Example #16 Replaced: {0}", count);

            // Delete content partially in inserted text, different author (paragraph 15)
            content = xDoc.Descendants(W.p).Skip(14).Take(1);
            regex = new Regex(" to help you prove your point");
            count = OpenXmlRegex.Replace(content, regex, "", null, true, "John Doe");
            Console.WriteLine("Example #17 Deleted: {0}", count);

            wDoc.MainDocumentPart.PutXDocument();
        }

        var sourcePres = new FileInfo("../../TestPresentation.pptx");
        var newPres = new FileInfo(Path.Combine(tempDi.FullName, "Modified.pptx"));
        File.Copy(sourcePres.FullName, newPres.FullName);
        using (PresentationDocument pDoc = PresentationDocument.Open(newPres.FullName, true))
        {
            foreach (var slidePart in pDoc.PresentationPart.SlideParts)
            {
                int count;
                var xDoc = slidePart.GetXDocument();
                Regex regex;
                IEnumerable<XElement> content;

                // Replace content
                content = xDoc.Descendants(A.p);
                regex = new Regex("Hello");
                count = OpenXmlRegex.Replace(content, regex, "H e l l o", null);
                Console.WriteLine("Example #18 Replaced: {0}", count);

                // If you absolutely want to preserve compatibility with PowerPoint 2007, then you will need to strip the xml:space="preserve" attribute throughout.
                // This is an issue for PowerPoint only, not Word, and for 2007 only.
                // The side-effect of this is that if a run has space at the beginning or end of it, the space will be stripped upon loading, and content/layout will be affected.
                xDoc.Descendants().Attributes(XNamespace.Xml + "space").Remove();

                slidePart.PutXDocument();
            }
        }
    }
}
