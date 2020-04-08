// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Drawing;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public class ImageInfo
    {
        public Bitmap Bitmap;
        public XAttribute ImgStyleAttribute;
        public string ContentType;
        public XElement DrawingElement;
        public string AltText;

        public const int EmusPerInch = 914400;
        public const int EmusPerCm = 360000;
    }
}