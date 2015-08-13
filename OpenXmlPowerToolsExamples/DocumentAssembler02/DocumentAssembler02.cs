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

            FileInfo assembledDoc = new FileInfo("../../AssembledDoc.docx");
            wmlAssembledDoc.SaveAs(assembledDoc.FullName);
        }
    }
}
