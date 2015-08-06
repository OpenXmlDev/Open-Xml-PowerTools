using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

class TestPmlTextReplacer
{
    static void Main(string[] args)
    {
        File.Delete("../../Test01out.pptx");
        File.Copy("../../Test01.pptx", "../../Test01out.pptx");
        using (PresentationDocument pDoc =
            PresentationDocument.Open("../../Test01out.pptx", true))
        {
            TextReplacer.SearchAndReplace(pDoc, "Hello", "Goodbye", true);
        }
        File.Delete("../../Test02out.pptx");
        File.Copy("../../Test02.pptx", "../../Test02out.pptx");
        using (PresentationDocument pDoc =
            PresentationDocument.Open("../../Test02out.pptx", true))
        {
            TextReplacer.SearchAndReplace(pDoc, "Hello", "Goodbye", true);
        }
        File.Delete("../../Test03out.pptx");
        File.Copy("../../Test03.pptx", "../../Test03out.pptx");
        using (PresentationDocument pDoc =
            PresentationDocument.Open("../../Test03out.pptx", true))
        {
            TextReplacer.SearchAndReplace(pDoc, "Hello", "Goodbye", false);
        }
    }
}
