using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

class OpenXmlRegexExample
{
    static void Main(string[] args)
    {
        var n = DateTime.Now;
        var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
        tempDi.Create();

        var sourceDoc = new FileInfo("../../TestDocument.docx");
        var newDoc = new FileInfo("Modified.docx");
        if (newDoc.Exists)
            newDoc.Delete();
        File.Copy(sourceDoc.FullName, newDoc.FullName);
        using (WordprocessingDocument wDoc = WordprocessingDocument.Open(newDoc.FullName, true))
        {
            int count;
            var xDoc = wDoc.MainDocumentPart.GetXDocument();
            Regex regex;
            IEnumerable<XElement> content;

            content = xDoc.Descendants(W.p);
            regex = new Regex("[.]\x020+");
            count = OpenXmlRegex.Replace(content, regex, "." + Environment.NewLine, null);

            foreach (var para in content)
            {
                var newPara = (XElement)TransformEnvironmentNewLineToParagraph(para);
                para.ReplaceNodes(newPara.Nodes());
            }

            wDoc.MainDocumentPart.PutXDocument();
        }
    }

    private static object TransformEnvironmentNewLineToParagraph(XNode node)
    {
        var element = node as XElement;
        if (element != null)
        {
            if (element.Name == W.p)
            {

            }

            return new XElement(element.Name,
                element.Attributes(),
                element.Nodes().Select(n => TransformEnvironmentNewLineToParagraph(n)));
        }
        return node;
    }
}
