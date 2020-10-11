using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class PresentationMLUtil
    {
        public static void FixUpPresentationDocument(PresentationDocument pDoc)
        {
            foreach (var part in pDoc.GetAllParts())
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
                    var xd = part.GetXDocument();
                    xd.Descendants().Attributes("smtClean").Remove();
                    xd.Descendants().Attributes("smtId").Remove();
                    part.PutXDocument();
                }
                if (part.ContentType == "application/vnd.openxmlformats-officedocument.vmlDrawing")
                {
                    string fixedContent = null;

                    using (var stream = part.GetStream(FileMode.Open, FileAccess.ReadWrite))
                    {
                        using var sr = new StreamReader(stream);
                        var input = sr.ReadToEnd();
                        var pattern = @"<!\[(?<test>.*)\]>";
                        fixedContent = Regex.Replace(input, pattern, m =>
                        {
                            var g = m.Groups[1].Value;
                            if (g.StartsWith("CDATA["))
                            {
                                return "<![" + g + "]>";
                            }
                            else
                            {
                                return "<![CDATA[" + g + "]]>";
                            }
                        },
                        RegexOptions.Multiline);

                        pattern = @"o:relid=[""'](?<id1>.*)[""'] o:relid=[""'](?<id2>.*)[""']";
                        fixedContent = Regex.Replace(fixedContent, pattern, m =>
                        {
                            var g = m.Groups[1].Value;
                            return @"o:relid=""" + g + @"""";
                        },
                        RegexOptions.Multiline);

                        fixedContent = fixedContent.Replace("</xml>ml>", "</xml>");

                        stream.SetLength(fixedContent.Length);
                    }
                    using var ms = new MemoryStream(Encoding.UTF8.GetBytes(fixedContent));
                    part.FeedData(ms);
                }
            }
        }
    }
}