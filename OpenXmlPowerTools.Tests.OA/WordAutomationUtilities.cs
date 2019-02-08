/***************************************************************************

Copyright (c) Microsoft Corporation 2012-2015.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

***************************************************************************/

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using Word = Microsoft.Office.Interop.Word;

namespace OxPt
{
    public class WordAutomationUtilities
    {
        public static void DoConversionViaWord(FileInfo newAltChunkBeforeFi, FileInfo newAltChunkAfterFi, XElement html)
        {
            var blankAltChunkFi = new DirectoryInfo(Path.Combine(TestUtil.SourceDir.FullName, "Blank-altchunk.docx"));
            File.Copy(blankAltChunkFi.FullName, newAltChunkBeforeFi.FullName);
            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(newAltChunkBeforeFi.FullName, true))
            {
                string altChunkId = "AltChunkId1";
                MainDocumentPart mainPart = myDoc.MainDocumentPart;
                AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(
                    "application/xhtml+xml", altChunkId);
                using (Stream chunkStream = chunk.GetStream(FileMode.Create, FileAccess.Write))
                using (StreamWriter stringStream = new StreamWriter(chunkStream))
                    stringStream.Write(html.ToString());
                XElement altChunk = new XElement(W.altChunk,
                    new XAttribute(R.id, altChunkId)
                );
                XDocument mainDocumentXDoc = myDoc.MainDocumentPart.GetXDocument();
                mainDocumentXDoc.Root
                    .Element(W.body)
                    .AddFirst(altChunk);
                myDoc.MainDocumentPart.PutXDocument();
            }

            WordAutomationUtilities.OpenAndSaveAs(newAltChunkBeforeFi.FullName, newAltChunkAfterFi.FullName);

            while (true)
            {
                try
                {
                    using (WordprocessingDocument wDoc = WordprocessingDocument.Open(newAltChunkAfterFi.FullName, true))
                    {
                        SimplifyMarkupSettings settings2 = new SimplifyMarkupSettings
                        {
                            RemoveMarkupForDocumentComparison = true,
                        };
                        MarkupSimplifier.SimplifyMarkup(wDoc, settings2);
                        XElement newRoot = (XElement)RemoveDivTransform(wDoc.MainDocumentPart.GetXDocument().Root);
                        wDoc.MainDocumentPart.GetXDocument().Root.ReplaceWith(newRoot);
                        wDoc.MainDocumentPart.PutXDocumentWithFormatting();
                    }
                    break;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                    continue;
                }
            }
        }

        private static object RemoveDivTransform(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.divId)
                    return null;
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => RemoveDivTransform(n)));
            }
            return node;
        }

        public static void SaveAsHtmlUsingWord(FileInfo src, FileInfo dest)
        {
            Word.Application app = new Word.Application();
            app.Visible = false;
            try
            {
                Word.Document doc = app.Documents.Open(src.FullName);
                doc.SaveAs2(dest.FullName, Word.WdSaveFormat.wdFormatFilteredHTML);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Console.WriteLine("Caught unexpected COM exception.");
                ((Microsoft.Office.Interop.Word._Application)app).Quit();
                Environment.Exit(0);
            }
            ((Microsoft.Office.Interop.Word._Application)app).Quit();
        }

        public static void OpenAndSaveAs(string fromFileName, string toFileName)
        {
            Word.Application app = new Word.Application();
            app.Visible = false;
            FileInfo fi = new FileInfo(fromFileName);
            try
            {
                FileInfo ffi = new FileInfo(fromFileName);
                Word.Document doc = app.Documents.Open(ffi.FullName);
                object FileFormat = Word.WdSaveFormat.wdFormatDocument;
                FileInfo tfi = new FileInfo(toFileName);
                object FileName = tfi.FullName;

                doc.SaveAs(tfi.FullName, Word.WdSaveFormat.wdFormatDocumentDefault);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Console.WriteLine("Caught unexpected COM exception.");
                ((Microsoft.Office.Interop.Word._Application)app).Quit();
                Environment.Exit(0);
            }
            ((Microsoft.Office.Interop.Word._Application)app).Quit();
        }

        public static void OpenAndSaveAs(FileInfo src, FileInfo dest)
        {
            Word.Application app = new Word.Application();
            app.Visible = false;
            try
            {
                Word.Document doc = app.Documents.Open(src.FullName);
                doc.SaveAs2(dest.FullName, Word.WdSaveFormat.wdFormatDocument);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Console.WriteLine("Caught unexpected COM exception.");
                ((Microsoft.Office.Interop.Word._Application)app).Quit();
                Environment.Exit(0);
            }
            ((Microsoft.Office.Interop.Word._Application)app).Quit();
        }
    }
}
