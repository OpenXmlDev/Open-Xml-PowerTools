using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public class ImageReplacer
    {
        public static bool ReplaceImage(WordprocessingDocument wDoc, string contentControlTag, string newImagePath)
        {
            var mainDocumentPart = wDoc.MainDocumentPart;
            var mdXDoc = mainDocumentPart.GetXDocument();
            var cc = mdXDoc.Descendants(W.sdt)
                .FirstOrDefault(sdt => (string)sdt.Elements(W.sdtPr).Elements(W.tag).Attributes(W.val).FirstOrDefault() == contentControlTag);

            if (cc != null)
            {
                var imageId = (string)cc.Descendants(A.blip).Attributes(R.embed).FirstOrDefault();

                if (imageId != null)
                {
                    ImagePart imagePart = (ImagePart)mainDocumentPart.GetPartById(imageId);
                    ReplaceNewImage(imagePart, newImagePath);
                }
            }

            return false;
        }

        private static void ReplaceNewImage(ImagePart imagePart, string newImagePath)
        {
            byte[] imageBytes = File.ReadAllBytes(newImagePath);
            BinaryWriter writer = new BinaryWriter(imagePart.GetStream());
            writer.Write(imageBytes);
            writer.Close();
        }
    }
}
