using System;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
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
            using var memoryStream = new MemoryStream();
            imageInfo.Bitmap.Save(memoryStream, imageInfo.Bitmap.RawFormat);
            var base64 = Convert.ToBase64String(memoryStream.ToArray());
            var format = imageInfo.Bitmap.RawFormat;
            var codec = ImageCodecInfo.GetImageDecoders().First(imageCodecInfo => imageCodecInfo.FormatID == format.Guid);
            var mimeType = codec.MimeType;

            var imageSource = $"data:{mimeType};base64,{base64}";

            return new XElement(Xhtml.img, new XAttribute(NoNamespace.src, imageSource), imageInfo.ImgStyleAttribute, imageInfo.AltText != null ? new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
        }
    }
}