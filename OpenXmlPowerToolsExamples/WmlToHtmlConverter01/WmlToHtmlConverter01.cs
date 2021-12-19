// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

/// <summary>
/// Converts Word Opem XML markup into HTML.
/// Image files are written to a separate image directory and referenced in the HTML document.
/// </summary>
internal static class WmlToHtmlConverter01
{
    private static void Main()
    {
        DirectoryInfo outputDirectory = CreateOutputDirectory();

        foreach (FileInfo docxFile in Directory.GetFiles("../../../", "*.docx").Select(path => new FileInfo(path)))
        {
            Console.WriteLine(docxFile.Name);
            ConvertToHtml(docxFile, outputDirectory);
        }
    }

    private static DirectoryInfo CreateOutputDirectory()
    {
        DateTime n = DateTime.Now;

        var outputDirectory = new DirectoryInfo(
            $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");

        outputDirectory.Create();

        return outputDirectory;
    }

    private static void ConvertToHtml(FileInfo docxFile, DirectoryInfo outputDirectory)
    {
        using var memoryStream = new MemoryStream();
        byte[] byteArray = File.ReadAllBytes(docxFile.FullName);
        memoryStream.Write(byteArray, 0, byteArray.Length);

        using WordprocessingDocument wordDocument = WordprocessingDocument.Open(memoryStream, true);

        string pageTitle = GetPageTitle(wordDocument, docxFile.Name);

        var htmlFile = new FileInfo(docxFile.Name.Replace(".docx", ".html"));
        htmlFile = new FileInfo(Path.Combine(outputDirectory.FullName, htmlFile.Name));

        string imageDirectoryPath = htmlFile.FullName.Substring(0, htmlFile.FullName.Length - 5) + "_files";
        var imageDirectory = new DirectoryInfo(imageDirectoryPath);
        var imageCounter = 0;

        var settings = new WmlToHtmlConverterSettings
        {
            // TODO: Determine max-width from size of content area.
            AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
            PageTitle = pageTitle,
            FabricateCssClasses = true,
            CssClassPrefix = "pt-",
            RestrictToSupportedLanguages = false,
            RestrictToSupportedNumberingFormats = false,
            ImageHandler = imageInfo => GetImageElement(imageInfo, imageDirectory, ref imageCounter),
        };

        XElement htmlElement = WmlToHtmlConverter.ConvertToHtml(wordDocument, settings);

        // Produce a HTML document with <!DOCTYPE html > declaration to tell the browser
        // we are using HTML5.
        var htmlDocument = new XDocument(new XDocumentType("html", null, null, null), htmlElement);

        // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
        // XEntity. PtOpenXmlUtil.cs define the XEntity class. See
        // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
        // for detailed explanation.
        //
        // If you further transform the XML tree returned by ConvertToHtmlTransform, you
        // must do it correctly, or entities will not be serialized properly.

        var htmlString = htmlDocument.ToString(SaveOptions.DisableFormatting);
        File.WriteAllText(htmlFile.FullName, htmlString, Encoding.UTF8);
    }

    private static string GetPageTitle(WordprocessingDocument wordDocument, string defaultPageTitle)
    {
        var title = (string) wordDocument.CoreFilePropertiesPart?.GetXDocument().Descendants(DC.title).FirstOrDefault();

        return title ?? defaultPageTitle;
    }

    private static XElement GetImageElement(ImageInfo imageInfo, DirectoryInfo imageDirectory, ref int imageCounter)
    {
        ImageFormat imageFormat = GetImageFormat(imageInfo);

        if (imageFormat == null)
        {
            return null;
        }

        string extension = imageFormat.ToString().ToLowerInvariant();
        var imageFileName = $"image{++imageCounter}.{extension}";

        try
        {
            // Ensure the image directory exists.
            imageDirectory.Create();

            imageInfo.Bitmap.Save($"{imageDirectory.FullName}/{imageFileName}", imageFormat);
        }
        catch (ExternalException)
        {
            return null;
        }

        return new XElement(Xhtml.img,
            new XAttribute(NoNamespace.src, $"{imageDirectory.Name}/{imageFileName}"),
            imageInfo.ImgStyleAttribute,
            imageInfo.AltText != null ? new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
    }

    private static ImageFormat GetImageFormat(ImageInfo imageInfo)
    {
        string extension = imageInfo.ContentType.Split('/')[1].ToLower();

        // Map tiff to gif and x-wmf to wmf.
        // Map unsupported image types to null.
        return extension switch
        {
            "png" => ImageFormat.Png,
            "gif" => ImageFormat.Gif,
            "bmp" => ImageFormat.Bmp,
            "jpeg" => ImageFormat.Jpeg,
            "tiff" => ImageFormat.Gif,
            "x-wmf" => ImageFormat.Wmf,
            _ => null,
        };
    }
}
