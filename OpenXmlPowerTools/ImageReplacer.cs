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

        /// <summary>
        /// 设置的宽度，可以为百分比，如果不填则默认使用页面宽度
        /// </summary>
        private readonly string _width;

        /// <summary>
        /// 设置高度，可以为百分比，如果不填则根据宽度缩放后自适应
        /// </summary>
        private readonly string _height;

        public byte[] ImageBytes { get; private set; }

        /// <summary>
        /// 图片替换器构造函数
        /// </summary>
        /// <param name="imageBytes">新图片的字节数组</param>
        public ImageUpdater(byte[] imageBytes, ImageFormat imageFormat, string width = null, string height = null)
        {
            ImageBytes = imageBytes;
            _imageFormat = imageFormat;
            _width = width;
            _height = height;
        }

        /// <summary>
        /// 新图片路径
        /// </summary>
        /// <param name="newImagePath"></param>
        public ImageUpdater(string newImagePath, ImageFormat imageFormat, string width = null, string height = null)
        {
            ImageBytes = File.ReadAllBytes(newImagePath);
            _imageFormat = imageFormat;
            _width = width;
            _height = height;
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
                                        string width = null,
                                        string height = null)
        {
            var replacer = new ImageUpdater(newImagePath, imageFormat, width, height);
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
                                        string width = null,
                                        string height = null)
        {
            var replacer = new ImageUpdater(imageBytes, imageFormat, width, height);
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
                UpdateImageMetrics(mdXDoc, cc);

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
        private void UpdateImageMetrics(XDocument mdXDoc, XElement cc)
        {
            float dpi = 96;

            // 得到图片宽高
            var imgStream = new MemoryStream(this.ImageBytes);
            var (imgWidth, imgHeight) = ImageHelper.GetImageMetrics(imgStream, this._imageFormat);

            // 得到页面宽高
            double pageWidth = 0;
            double pageHeight = 0;

            var pageSize = mdXDoc.Descendants(W.sectPr).FirstOrDefault()?.Descendants(W.pgSz).FirstOrDefault();
            if (pageSize != null)
            {
                pageWidth = Convert.ToDouble(pageSize.Attribute(W.w.GetName("w")).Value) / 20;
                pageHeight = Convert.ToDouble(pageSize.Attribute(W.w.GetName("h")).Value) / 20;
            }

            // 得到设置宽高，传进来的参数
            int.TryParse(this._width, out var settingWidth);
            int.TryParse(this._height, out var settingHeight);

            // 如果设置的宽度是百分比，那么按百分比计算宽度
            if (this._width != null && this._width.EndsWith("%"))
            {
                // 先计算比例，避免宽度变更后比例也跟着变化。                
                var radio = imgWidth / (float)imgHeight;
                double percent = Convert.ToDouble(this._width.TrimEnd('%')) / 100;
                imgWidth = (int)(pageWidth * percent);

                // 如果没有设置高度，那么按比例计算高度
                if (this._height == null)
                {
                    // 那么按比例计算高度
                    imgHeight = (int)(imgWidth / radio);
                }
            }
            else if (settingWidth > 0) // 如果设置了宽度，那么按设置的宽度来
            {
                imgWidth = settingWidth;
            }

            // 如果度是百分比，那么按百分比计算高度
            if (this._height != null && this._height.EndsWith("%"))
            {
                // 先计算比例，避免宽度变更后比例也跟着变化。
                var radio = imgWidth / (float)imgHeight;
                double percent = Convert.ToDouble(this._height.TrimEnd('%')) / 100;
                imgHeight = (int)(pageHeight * percent);

                // 如果没有设置宽度，那么按比例计算宽度
                if (this._width == null)
                {
                    // 那么按比例计算宽度
                    imgWidth = (int)(imgHeight * radio);
                }
            }
            else if (settingHeight > 0) // 如果设置了高度，那么按设置的高度来
            {
                imgHeight = settingHeight;
            }

            // 换算成EMUS
            var cx = (long)(imgWidth / dpi * 914400);
            var cy = (long)(imgHeight / dpi * 914400);

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

        // 按比例缩放图片的方法，传入图片原始宽度和高度，传入缩放后的宽度
        private static (int width, int height) ScaleImage(int originalWidth, int originalHeight, int scaledWidth)
        {
            var ratio = (float)originalWidth / originalHeight;
            var scaledHeight = (int)(scaledWidth / ratio);
            return (scaledWidth, scaledHeight);
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
