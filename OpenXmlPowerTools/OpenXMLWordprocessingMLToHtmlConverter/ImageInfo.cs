using System.IO;
using System.Xml.Linq;

namespace Codeuctivity.OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter
{
    public class ImageInfo
    {
        public Stream Image { get; set; }
        public XAttribute ImgStyleAttribute { get; set; }
        public string ContentType { get; set; }
        public XElement DrawingElement { get; set; }
        public string AltText { get; set; }
        public static int EmusPerCm => 360000;
        public static int EmusPerInch => 914400;
    }
}