using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace FormattingAssembler01
{
    class FormattingAssembler01
    {
        static void Main(string[] args)
        {
            DirectoryInfo di = new DirectoryInfo("../../");
            foreach (var file in di.GetFiles("*out.docx"))
                file.Delete();
            foreach (var file in di.GetFiles("*.docx"))
            {
                Console.WriteLine(file.Name);
                var newFile = new FileInfo("../../" + file.Name.Replace(".docx", "out.docx"));
                File.Copy(file.FullName, newFile.FullName);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(newFile.FullName, true))
                {
                    FormattingAssemblerSettings settings = new FormattingAssemblerSettings()
                    {
                        ClearStyles = true,
                        RemoveStyleNamesFromParagraphAndRunProperties = true,
                        CreateHtmlConverterAnnotationAttributes = true,
                        OrderElementsPerStandard = true,
                        RestrictToSupportedLanguages = true,
                        RestrictToSupportedNumberingFormats = true,
                    };
                    FormattingAssembler.AssembleFormatting(wDoc, settings);
                }
            }
        }
    }
}

