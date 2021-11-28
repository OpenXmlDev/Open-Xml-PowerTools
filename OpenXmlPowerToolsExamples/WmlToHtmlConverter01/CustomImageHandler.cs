using OpenXmlPowerTools;
using OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter;
using System.IO;
using System.Xml.Linq;

internal class CustomImageHandler : IImageHandler
{
    public CustomImageHandler(string imageDirectoryName)
    {
        ImageDirectoryName = imageDirectoryName;
        ImageCounter = 0;
    }

    public string ImageDirectoryName { get; }
    public int ImageCounter { get; private set; }

    public XElement TransformImage(ImageInfo imageInfo)
    {
        var localDirInfo = new DirectoryInfo(ImageDirectoryName);
        if (!localDirInfo.Exists)
        {
            localDirInfo.Create();
        }

        ++ImageCounter;
        var extension = imageInfo.ContentType.Split('/')[1].ToLower();

        var imageFileName = ImageDirectoryName + "/image" + ImageCounter.ToString() + "." + extension;
        try
        {
            using var fileStream = new FileStream(imageFileName, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.ReadWrite);
            imageInfo.Image.CopyTo(fileStream);
        }
        catch (System.Runtime.InteropServices.ExternalException)
        {
            return null;
        }
        var imageSource = localDirInfo.Name + "/image" + ImageCounter.ToString() + "." + extension;

        var img = new XElement(Xhtml.img, new XAttribute(NoNamespace.src, imageSource), imageInfo.ImgStyleAttribute, imageInfo.AltText != null ? new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
        return img;
    }
}