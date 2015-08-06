using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace OpenXmlPowerTools
{
    public class WordAutomationUtilities
    {
        public static void ProcessFilesUsingWordAutomation(List<string> fileNames)
        {
            Word.Application app = new Word.Application();
            app.Visible = false;
            foreach (string fileName in fileNames)
            {
                FileInfo fi = new FileInfo(fileName);
                try
                {
                    Word.Document doc = app.Documents.Open(fi.FullName);
                    doc.Save();
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Console.WriteLine("Caught unexpected COM exception.");
                    ((Microsoft.Office.Interop.Word._Application)app).Quit();
                    Environment.Exit(0);
                }
            }
            ((Microsoft.Office.Interop.Word._Application)app).Quit();
        }
    }
}
