using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools.Helpers;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    /// <summary>
    /// 图片替换器
    /// </summary>
    public class ImageUpdater
    {
        private readonly ImageFormat _imageFormat;
        private readonly int? _maxWidth;
        private readonly int? _maxHeight;

        public byte[] ImageBytes { get; private set; }

        /// <summary>
        /// 图片替换器构造函数
        /// </summary>
        /// <param name="imageBytes">新图片的字节数组</param>
        public ImageUpdater(byte[] imageBytes, ImageFormat imageFormat, int? maxWidth, int? maxHeight)
        {
            ImageBytes = imageBytes;
            _imageFormat = imageFormat;
            _maxWidth = maxWidth;
            _maxHeight = maxHeight;
        }

        /// <summary>
        /// 新图片路径
        /// </summary>
        /// <param name="newImagePath"></param>
        public ImageUpdater(string newImagePath, ImageFormat imageFormat, int? maxWidth, int? maxHeight)
        {
            ImageBytes = File.ReadAllBytes(newImagePath);
            _imageFormat = imageFormat;
            _maxWidth = maxWidth;
            _maxHeight = maxHeight;
        }

        /// <summary>
        /// 替换图片
        /// </summary>
        /// <param name="wDoc">Word文档对象</param>
        /// <param name="contentControlTag">要替换的控件tag</param>
        /// <param name="newImagePath">新图片的路径</param>
        /// <returns></returns>
        public static bool ReplaceImage(WordprocessingDocument wDoc,
                                        string contentControlTag,
                                        string newImagePath,
                                        ImageFormat imageFormat,
                                        int? maxWidth = null,
                                        int? maxHeight = null)
        {
            var replacer = new ImageUpdater(newImagePath, imageFormat, maxWidth, maxHeight);
            return replacer.Replace(wDoc, contentControlTag);
        }

        /// <summary>
        /// 替换图片
        /// </summary>
        /// <param name="wDoc">Word文档对象</param>
        /// <param name="contentControlTag">要替换的控件tag</param>
        /// <param name="imageBytes">新图片的路径</param>
        /// <returns></returns>
        public static bool ReplaceImage(WordprocessingDocument wDoc,
                                        string contentControlTag,
                                        byte[] imageBytes,
                                        ImageFormat imageFormat,
                                        int? maxWidth = null,
                                        int? maxHeight = null)
        {
            var replacer = new ImageUpdater(imageBytes, imageFormat, maxWidth, maxHeight);
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
                // 替换imagePart
                var imageId = (string)cc.Descendants(A.blip).Attributes(R.embed).FirstOrDefault();

                if (imageId != null)
                {
                    ImagePart imagePart = (ImagePart)mainDocumentPart.GetPartById(imageId);
                    ReplaceNewImage(imagePart, this.ImageBytes);
                }


                // 修改宽度和高度                
                UpdateImageMetrics(cc);

                // 替换cc                
                var paragraph = cc.Descendants(W.sdtContent).Descendants(W.p).FirstOrDefault();
                if (paragraph != null)
                {
                    cc.ReplaceWith(paragraph);
                }

                mainDocumentPart.PutXDocument();
                return true;
            }

            return false;
        }

        /// <summary>
        /// 更新图片宽高
        /// </summary>
        /// <param name="cc"></param>
        private void UpdateImageMetrics(XElement cc)
        {
            float dpi = 96;

            var imgStream = new MemoryStream(this.ImageBytes);
            var (w, h) = ImageHelper.GetImageMetrics(imgStream, this._imageFormat);

            // 如果设置了最大宽度或最大高度，则按比例缩放
            if (this._maxWidth.HasValue || this._maxHeight.HasValue)
            {
                var ratio = (float)w / h;
                if (this._maxWidth.HasValue && this._maxHeight.HasValue)
                {
                    if (ratio > (float)this._maxWidth.Value / this._maxHeight.Value)
                    {
                        w = this._maxWidth.Value;
                        h = (int)(w / ratio);
                    }
                    else
                    {
                        h = this._maxHeight.Value;
                        w = (int)(h * ratio);
                    }
                }
                else if (this._maxWidth.HasValue)
                {
                    w = this._maxWidth.Value;
                    h = (int)(w / ratio);
                }
                else
                {
                    h = this._maxHeight.Value;
                    w = (int)(h * ratio);
                }
            }

            var cx = (long)(w / dpi * 914400);
            var cy = (long)(h / dpi * 914400);

            var drawing = cc.Descendants(W.drawing).FirstOrDefault();
            if (drawing != null)
            {
                var extent = drawing.Descendants(WP.extent).FirstOrDefault();
                if (extent != null)
                {
                    extent.SetAttributeValue("cx", cx);
                    extent.SetAttributeValue("cy", cy);
                }
                var aExtent = drawing.Descendants(A.graphic).Descendants(Pic.spPr).Descendants(A.ext).FirstOrDefault();
                if (aExtent != null)
                {
                    aExtent.SetAttributeValue("cx", cx);
                    aExtent.SetAttributeValue("cy", cy);
                }
            }
        }

        /// <summary>
        /// 替换新图片
        /// </summary>
        /// <param name="imagePart"></param>
        /// <param name="imageBytes"></param>
        private void ReplaceNewImage(ImagePart imagePart, byte[] imageBytes)
        {
            var stream = imagePart.GetStream();

            BinaryWriter writer = new BinaryWriter(stream);
            writer.Write(imageBytes);
            writer.Close();
        }
    }
}
