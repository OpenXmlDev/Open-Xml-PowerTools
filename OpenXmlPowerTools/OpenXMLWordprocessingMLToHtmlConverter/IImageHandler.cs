using System.Xml.Linq;

namespace OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter
{
    /// <summary>
    /// Implement an imageHandler to get image support in HTML
    /// </summary>
    public interface IImageHandler
    {
        /// <summary>
        /// Transforms OpenXml Images to HTML embeddable images
        /// </summary>
        /// <param name="imageInfo"></param>
        /// <returns></returns>
        public XElement TransformImage(ImageInfo imageInfo);
    }
}