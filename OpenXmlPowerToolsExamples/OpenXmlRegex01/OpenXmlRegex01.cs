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

namespace OpenXmlRegex01
{
    public class OpenXmlRegexExample
    {
        public static void Main(string[] args)
        {
            DateTime n = DateTime.Now;
            var tempDi = new DirectoryInfo(
                $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");
            tempDi.Create();

            var sourceDoc = new FileInfo("../../TestDocument.docx");
            var newDoc = new FileInfo(Path.Combine(tempDi.FullName, "Modified.docx"));
            File.Copy(sourceDoc.FullName, newDoc.FullName);
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(newDoc.FullName, true))
            {
                XDocument xDoc = wDoc.MainDocumentPart.GetXDocument();

                // Match content (paragraph 1)
                IEnumerable<XElement> content = xDoc.Descendants(W.p).Take(1);
                var regex = new Regex("Video");
                int count = OpenXmlRegex.Match(content, regex);
                Console.WriteLine("Example #1 Count: {0}", count);

                // Match content, case insensitive (paragraph 1)
                content = xDoc.Descendants(W.p).Take(1);
                regex = new Regex("video", RegexOptions.IgnoreCase);
                count = OpenXmlRegex.Match(content, regex);
                Console.WriteLine("Example #2 Count: {0}", count);

                // Match content, with callback (paragraph 1)
                content = xDoc.Descendants(W.p).Take(1);
                regex = new Regex("video", RegexOptions.IgnoreCase);
                OpenXmlRegex.Match(content, regex, (element, match) =>
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

                const string leftDoubleQuotationMarks = @"[\u0022“„«»”]";
                const string words = @"[\w\-&/]+(?:\s[\w\-&/]+)*";
                const string rightDoubleQuotationMarks = @"[\u0022”‟»«“]";

                // Replace content using replacement pattern (paragraph 16)
                content = xDoc.Descendants(W.p).Skip(15).Take(1);
                regex = new Regex($"{leftDoubleQuotationMarks}(?<words>{words}){rightDoubleQuotationMarks}");
                count = OpenXmlRegex.Replace(content, regex, "‘${words}’", null);
                Console.WriteLine("Example #18 Replaced: {0}", count);

                // Replace content using replacement pattern in partially inserted text (paragraph 17)
                content = xDoc.Descendants(W.p).Skip(16).Take(1);
                regex = new Regex($"{leftDoubleQuotationMarks}(?<words>{words}){rightDoubleQuotationMarks}");
                count = OpenXmlRegex.Replace(content, regex, "‘${words}’", null, true, "John Doe");
                Console.WriteLine("Example #19 Replaced: {0}", count);

                // Replace content using replacement pattern (paragraph 18)
                content = xDoc.Descendants(W.p).Skip(17).Take(1);
                regex = new Regex($"({leftDoubleQuotationMarks})(video)({rightDoubleQuotationMarks})");
                count = OpenXmlRegex.Replace(content, regex, "$1audio$3", null, true, "John Doe");
                Console.WriteLine("Example #20 Replaced: {0}", count);

                // Recognize tabs (paragraph 19)
                content = xDoc.Descendants(W.p).Skip(18).Take(1);
                regex = new Regex(@"([1-9])\.\t");
                count = OpenXmlRegex.Replace(content, regex, "($1)\t", null);
                Console.WriteLine("Example #21 Replaced: {0}", count);

                // The next two examples deal with line breaks, i.e., the <w:br/> elements.
                // Note that you should use the U+000D (Carriage Return) character (i.e., '\r')
                // to match a <w:br/> (or <w:cr/>) and replace content with a <w:br/> element.
                // Depending on your platform, the end of line character(s) provided by
                // Environment.NewLine might be "\n" (Unix), "\r\n" (Windows), or "\r" (Mac).

                // Recognize tabs and insert line breaks (paragraph 20).
                content = xDoc.Descendants(W.p).Skip(19).Take(1);
                regex = new Regex($@"([1-9])\.{UnicodeMapper.HorizontalTabulation}");
                count = OpenXmlRegex.Replace(content, regex, $"Article $1{UnicodeMapper.CarriageReturn}", null);
                Console.WriteLine("Example #22 Replaced: {0}", count);

                // Recognize and remove line breaks (paragraph 21)
                content = xDoc.Descendants(W.p).Skip(20).Take(1);
                regex = new Regex($"{UnicodeMapper.CarriageReturn}");
                count = OpenXmlRegex.Replace(content, regex, " ", null);
                Console.WriteLine("Example #23 Replaced: {0}", count);

                // Remove soft hyphens (paragraph 22)
                List<XElement> paras = xDoc.Descendants(W.p).Skip(21).Take(1).ToList();
                count = OpenXmlRegex.Replace(paras, new Regex($"{UnicodeMapper.SoftHyphen}"), "", null);
                count += OpenXmlRegex.Replace(paras, new Regex("use"), "no longer use", null);
                Console.WriteLine("Example #24 Replaced: {0}", count);

                // The next example deals with symbols (i.e., w:sym elements).
                // To work with symbols, you should acquire the Unicode values for the
                // symbols you wish to match or use in replacement patterns. The reason
                // is that UnicodeMapper will (a) mimic Microsoft Word in shifting the
                // Unicode values into the Unicode private use area (by adding U+F000)
                // and (b) use replacements for Unicode values that have been used in
                // conjunction with different fonts already (by adding U+E000).
                //
                // The replacement Únicode values will depend on the order in which
                // symbols are retrieved. Therefore, you should not rely on any fixed
                // assignment.
                //
                // In the example below, pencil will be represented by U+F021, whereas
                // spider (same value with different font) will be represented by U+E001.
                // If spider had been assigned first, spider would be U+F021 and pencil
                // would be U+E001.
                char oldPhone = UnicodeMapper.SymToChar("Wingdings", 40);
                char newPhone = UnicodeMapper.SymToChar("Wingdings", 41);
                char pencil = UnicodeMapper.SymToChar("Wingdings", 0x21);
                char spider = UnicodeMapper.SymToChar("Webdings", 0x21);

                // Replace or comment on symbols (paragraph 23)
                paras = xDoc.Descendants(W.p).Skip(22).Take(1).ToList();
                count = OpenXmlRegex.Replace(paras, new Regex($"{oldPhone}"), $"{newPhone} (replaced with new phone)", null);
                count += OpenXmlRegex.Replace(paras, new Regex($"({pencil})"), "$1 (same pencil)", null);
                count += OpenXmlRegex.Replace(paras, new Regex($"({spider})"), "$1 (same spider)", null);
                Console.WriteLine("Example #25 Replaced: {0}", count);

                wDoc.MainDocumentPart.PutXDocument();
            }

            var sourcePres = new FileInfo("../../TestPresentation.pptx");
            var newPres = new FileInfo(Path.Combine(tempDi.FullName, "Modified.pptx"));
            File.Copy(sourcePres.FullName, newPres.FullName);
            using (PresentationDocument pDoc = PresentationDocument.Open(newPres.FullName, true))
            {
                foreach (SlidePart slidePart in pDoc.PresentationPart.SlideParts)
                {
                    XDocument xDoc = slidePart.GetXDocument();

                    // Replace content
                    IEnumerable<XElement> content = xDoc.Descendants(A.p);
                    var regex = new Regex("Hello");
                    int count = OpenXmlRegex.Replace(content, regex, "H e l l o", null);
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
}
