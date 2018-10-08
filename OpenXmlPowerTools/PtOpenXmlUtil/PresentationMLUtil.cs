using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public static class PresentationMLUtil
    {
        public static void FixUpPresentationDocument(PresentationDocument pDoc)
        {
            foreach (OpenXmlPart part in pDoc.GetAllParts())
            {
                if (part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slide+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.theme+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.chart+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml" ||
                    part.ContentType == "application/vnd.ms-office.drawingml.diagramDrawing+xml")
                {
                    XDocument xd = part.GetXDocument();
                    xd.Descendants().Attributes("smtClean").Remove();
                    xd.Descendants().Attributes("smtId").Remove();
                    part.PutXDocument();
                }

                if (part.ContentType == "application/vnd.openxmlformats-officedocument.vmlDrawing")
                {
                    string fixedContent = null;

                    using (Stream stream = part.GetStream(FileMode.Open, FileAccess.ReadWrite))
                    {
                        using (var sr = new StreamReader(stream))
                        {
                            //string input = @"    <![if gte mso 9]><v:imagedata o:relid=""rId15""";
                            string input = sr.ReadToEnd();
                            var pattern = @"<!\[(?<test>.*)\]>";

                            //string replacement = "<![CDATA[${test}]]>";
                            //fixedContent = Regex.Replace(input, pattern, replacement, RegexOptions.Multiline);
                            fixedContent = Regex.Replace(input, pattern, m =>
                                {
                                    string g = m.Groups[1].Value;
                                    if (g.StartsWith("CDATA["))
                                        return "<![" + g + "]>";

                                    return "<![CDATA[" + g + "]]>";
                                },
                                RegexOptions.Multiline);

                            //var input = @"xxxxx o:relid=""rId1"" o:relid=""rId1"" xxxxx";
                            pattern = @"o:relid=[""'](?<id1>.*)[""'] o:relid=[""'](?<id2>.*)[""']";
                            fixedContent = Regex.Replace(fixedContent, pattern, m =>
                                {
                                    string g = m.Groups[1].Value;
                                    return @"o:relid=""" + g + @"""";
                                },
                                RegexOptions.Multiline);

                            fixedContent = fixedContent.Replace("</xml>ml>", "</xml>");

                            stream.SetLength(fixedContent.Length);
                        }
                    }

                    using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(fixedContent)))
                        part.FeedData(ms);
                }
            }
        }
    }
}
