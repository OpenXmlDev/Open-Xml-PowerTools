using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats;
using System;
using System.IO;
using System.Xml.Linq;

namespace OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter
{
    /// <summary>
    /// Default image handler
    /// </summary>
    public class ImageHandler : IImageHandler
    {

        /// <summary>
        /// Transforms OpenXml Images to HTML embeddable images
        /// </summary>
        /// <param name="imageInfo"></param>
        /// <returns></returns>
        public XElement TransformImage(ImageInfo imageInfo)
        {
            IImageFormat format;
            using var imageStream = new MemoryStream();
            imageInfo.Image.CopyTo(imageStream);
            imageStream.Position = 0;
            using var image = Image.Load(imageStream, out format);
            var base64 = Convert.ToBase64String(imageStream.ToArray());
            var mimeType = format.DefaultMimeType;

            var imageSource = $"data:{mimeType};base64,{base64}";

            return new XElement(Xhtml.img, new XAttribute(NoNamespace.src, imageSource), imageInfo.ImgStyleAttribute, imageInfo.AltText != null ? new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
        }
    }
}