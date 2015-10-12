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
        var n = DateTime.Now;
        var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
        tempDi.Create();

        File.Copy("../../Test01.pptx", Path.Combine(tempDi.FullName, "Test01out.pptx"));
        using (PresentationDocument pDoc =
            PresentationDocument.Open(Path.Combine(tempDi.FullName, "Test01out.pptx"), true))
        {
            TextReplacer.SearchAndReplace(pDoc, "Hello", "Goodbye", true);
        }
        File.Copy("../../Test02.pptx", Path.Combine(tempDi.FullName, "Test02out.pptx"));
        using (PresentationDocument pDoc =
            PresentationDocument.Open(Path.Combine(tempDi.FullName, "Test02out.pptx"), true))
        {
            TextReplacer.SearchAndReplace(pDoc, "Hello", "Goodbye", true);
        }
        File.Copy("../../Test03.pptx", Path.Combine(tempDi.FullName, "Test03out.pptx"));
        using (PresentationDocument pDoc =
            PresentationDocument.Open(Path.Combine(tempDi.FullName, "Test03out.pptx"), true))
        {
            TextReplacer.SearchAndReplace(pDoc, "Hello", "Goodbye", false);
        }
    }
}
