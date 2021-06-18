using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter
{
    public class FontHandler : IFontHandler
    {
        public string TranslateParagraphStyleFont(XElement paragraph)
        {
            return (string)paragraph.Attributes(PtOpenXml.FontName).FirstOrDefault();
        }

        public string TranslateRunStyleFont(XElement run)
        {
            var sym = run.Element(W.sym);
            return sym != null ? (string)sym.Attributes(W.font).FirstOrDefault() : (string)run.Attributes(PtOpenXml.FontName).FirstOrDefault();
        }
    }
}