using System.Xml.Linq;

namespace Codeuctivity.OpenXMLWordprocessingMLToHtmlConverter
{
    public interface IFontHandler
    {
        public string TranslateRunStyleFont(XElement run);
        public string TranslateParagraphStyleFont(XElement paragraph);
    }
}