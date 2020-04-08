// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Drawing;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public class ImageInfo
    {
        public Bitmap Bitmap { get; set; }
        public XAttribute ImgStyleAttribute { get; set; }
        public string ContentType { get; set; }
        public XElement DrawingElement { get; set; }
        public string AltText { get; set; }
        public static int EmusPerCm => 360000;
        public static int EmusPerInch => 914400;
    }
}