using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;

namespace OpenXmlPowerTools.Helpers
{
    public class ImageHelper
    {
        public static (int width, int height) GetImageMetrics(Stream imageData, ImageFormat imageFormat)
        {
            var width = 0;
            var height = 0;

            if (imageFormat == ImageFormat.Emf)
            {
                using (var metafile = new Metafile(imageData))
                {
                    width = metafile.Width; // 获取宽度
                    height = metafile.Height; // 获取高度                                          
                }
            }
            else
            {
                using (var image = Image.FromStream(imageData))
                {
                    width = image.Width; // 获取宽度
                    height = image.Height; // 获取高度
                }
            }

            return (width, height);
        }

        /// <summary>
        /// 获取图片的宽和高
        /// </summary>
        /// <param name="newImagePath"></param>
        /// <returns></returns>
        public static (int width, int height) GetImageMetrics(string newImagePath, ImageFormat imageFormat)
        {
            var ext = Path.GetExtension(newImagePath).ToLower();
            var stream = File.OpenRead(newImagePath);
            return GetImageMetrics(stream, imageFormat);
        }
    }
}
