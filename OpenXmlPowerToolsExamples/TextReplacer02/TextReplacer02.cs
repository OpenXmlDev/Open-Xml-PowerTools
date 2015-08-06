using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace OpenXmlPowerTools
{
    class Program
    {
        static void Main(string[] args)
        {
            DirectoryInfo di = new DirectoryInfo(".");
            foreach (var file in di.GetFiles("*.docx"))
                file.Delete();
            DirectoryInfo di2 = new DirectoryInfo("../../");
            foreach (var file in di2.GetFiles("*.docx"))
                file.CopyTo(di.FullName + "/" + file.Name);

            using (WordprocessingDocument doc = WordprocessingDocument.Open("Test01.docx", true))
                TextReplacer.SearchAndReplace(doc, "the", "this", false);
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open("Test02.docx", true))
                    TextReplacer.SearchAndReplace(doc, "the", "this", false);
            }
            catch (Exception) { }
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open("Test03.docx", true))
                    TextReplacer.SearchAndReplace(doc, "the", "this", false);
            }
            catch (Exception) { }
            using (WordprocessingDocument doc = WordprocessingDocument.Open("Test04.docx", true))
                TextReplacer.SearchAndReplace(doc, "the", "this", true);
            using (WordprocessingDocument doc = WordprocessingDocument.Open("Test05.docx", true))
                TextReplacer.SearchAndReplace(doc, "is on", "is above", true);
            using (WordprocessingDocument doc = WordprocessingDocument.Open("Test06.docx", true))
                TextReplacer.SearchAndReplace(doc, "the", "this", false);
            using (WordprocessingDocument doc = WordprocessingDocument.Open("Test07.docx", true))
                TextReplacer.SearchAndReplace(doc, "the", "this", true);
            using (WordprocessingDocument doc = WordprocessingDocument.Open("Test08.docx", true))
                TextReplacer.SearchAndReplace(doc, "the", "this", true);
            using (WordprocessingDocument doc = WordprocessingDocument.Open("Test09.docx", true))
                TextReplacer.SearchAndReplace(doc, "===== Replace this text =====", "***zzz***", true);
        }
    }
}
