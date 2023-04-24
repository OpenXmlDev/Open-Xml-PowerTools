﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    /// <summary>
    /// 图片替换器
    /// </summary>
    public class ImageReplacer
    {
        public byte[] ImageBytes { get; private set; }

        /// <summary>
        /// 图片替换器构造函数
        /// </summary>
        /// <param name="imageBytes">新图片的字节数组</param>
        public ImageReplacer(byte[] imageBytes)
        {
            ImageBytes = imageBytes;
        }

        /// <summary>
        /// 新图片路径
        /// </summary>
        /// <param name="newImagePath"></param>
        public ImageReplacer(string newImagePath)
        {
            ImageBytes = File.ReadAllBytes(newImagePath);
        }

        /// <summary>
        /// 替换图片
        /// </summary>
        /// <param name="wDoc">Word文档对象</param>
        /// <param name="contentControlTag">要替换的控件tag</param>
        /// <param name="newImagePath">新图片的路径</param>
        /// <returns></returns>
        public static bool ReplaceImage(WordprocessingDocument wDoc, string contentControlTag, string newImagePath)
        {
            var replacer = new ImageReplacer(newImagePath);
            return replacer.Replace(wDoc, contentControlTag);
        }

        /// <summary>
        /// 替换图片
        /// </summary>
        /// <param name="wDoc">Word文档对象</param>
        /// <param name="contentControlTag">要替换的控件tag</param>
        /// <param name="imageBytes">新图片的路径</param>
        /// <returns></returns>
        public static bool ReplaceImage(WordprocessingDocument wDoc, string contentControlTag, byte[] imageBytes)
        {
            var replacer = new ImageReplacer(imageBytes);
            return replacer.Replace(wDoc, contentControlTag);
        }

        /// <summary>
        /// 替换图片
        /// </summary>
        /// <param name="wDoc"></param>
        /// <param name="contentControlTag"></param>
        /// <returns></returns>
        private bool Replace(WordprocessingDocument wDoc, string contentControlTag)
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
                    ReplaceNewImage(imagePart, this.ImageBytes);
                    return true;
                }
            }

            return false;
        }

        private void ReplaceNewImage(ImagePart imagePart, byte[] imageBytes)
        {
            BinaryWriter writer = new BinaryWriter(imagePart.GetStream());
            writer.Write(imageBytes);
            writer.Close();
        }

        private void ReplaceNewImage(ImagePart imagePart, string newImagePath)
        {
            byte[] imageBytes = File.ReadAllBytes(newImagePath);
            ReplaceNewImage(imagePart, imageBytes);
        }
    }
}
