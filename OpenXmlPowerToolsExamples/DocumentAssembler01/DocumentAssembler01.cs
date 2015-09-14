using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace OpenXmlPowerTools
{
    class Program
    {
        static void Main(string[] args)
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            FileInfo templateDoc = new FileInfo("../../TemplateDocument.docx");
            FileInfo dataFile = new FileInfo("../../Data.xml");

            WmlDocument wmlDoc = new WmlDocument(templateDoc.FullName);
            XElement data = XElement.Load(dataFile.FullName);
            bool templateError;
            WmlDocument wmlAssembledDoc = DocumentAssembler.AssembleDocument(wmlDoc, data, out templateError);
            if (templateError)
            {
                Console.WriteLine("Errors in template.");
                Console.WriteLine("See AssembledDoc.docx to determine the errors in the template.");
            }

            FileInfo assembledDoc = new FileInfo(Path.Combine(tempDi.FullName, "AssembledDoc.docx"));
            wmlAssembledDoc.SaveAs(assembledDoc.FullName);
        }
    }
}
