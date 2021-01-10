using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

internal class WmlToHtmlConverterHelper
{
    private static void Main()
    {
        var n = DateTime.Now;
        var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
        tempDi.Create();

        /*
         * This example loads each document into a byte array, then into a memory stream, so that the document can be opened for writing without
         * modifying the source document.
         */
        foreach (var file in Directory.GetFiles("../../../", "*.docx"))
        {
            ConvertToHtml(file, tempDi.FullName);
        }
    }

    public static void ConvertToHtml(string file, string outputDirectory)
    {
        var fi = new FileInfo(file);
        Console.WriteLine(fi.Name);
        var byteArray = File.ReadAllBytes(fi.FullName);
        using var memoryStream = new MemoryStream();
        memoryStream.Write(byteArray, 0, byteArray.Length);
        using var wDoc = WordprocessingDocument.Open(memoryStream, true);
        var destFileName = new FileInfo(fi.Name.Replace(".docx", ".html"));
        if (outputDirectory != null && !string.IsNullOrEmpty(outputDirectory))
        {
            var di = new DirectoryInfo(outputDirectory);
            if (!di.Exists)
            {
                throw new OpenXmlPowerToolsException("Output directory does not exist");
            }
            destFileName = new FileInfo(Path.Combine(di.FullName, destFileName.Name));
        }
        var imageDirectoryName = destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";

        var pageTitle = fi.FullName;
        var part = wDoc.CoreFilePropertiesPart;
        if (part != null)
        {
            pageTitle = (string)part.GetXDocument().Descendants(DC.title).FirstOrDefault() ?? fi.FullName;
        }

        // TODO: Determine max-width from size of content area.
        var settings = new WmlToHtmlConverterSettings()
        {
            AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
            PageTitle = pageTitle,
            FabricateCssClasses = true,
            CssClassPrefix = "pt-",
            RestrictToSupportedLanguages = false,
            RestrictToSupportedNumberingFormats = false
        };
        var htmlElement = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

        // Produce HTML document with <!DOCTYPE html > declaration to tell the browser we are using HTML5.
        var html = new XDocument(new XDocumentType("html", null, null, null), htmlElement);

        // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
        // XEntity.  PtOpenXmlUtil.cs define the XEntity class.  See
        // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
        // for detailed explanation.
        //
        // If you further transform the XML tree returned by ConvertToHtmlTransform, you
        // must do it correctly, or entities will not be serialized properly.

        var htmlString = html.ToString(SaveOptions.DisableFormatting);
        File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
    }
}