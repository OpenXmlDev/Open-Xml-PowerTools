// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using System;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class WmlComparerExtensions
    {
        public static XElement GetMainDocumentBody(this WordprocessingDocument wordDocument)
        {
            return wordDocument.GetMainDocumentRoot().Element(W.body) ?? throw new ArgumentException("Invalid document.");
        }

        public static XElement GetMainDocumentRoot(this WordprocessingDocument wordDocument)
        {
            return wordDocument.MainDocumentPart?.GetXElement() ?? throw new ArgumentException("Invalid document.");
        }

        public static XElement GetXElement(this OpenXmlPart part)
        {
            return part.GetXDocument()?.Root ?? throw new ArgumentException("Invalid document.");
        }
    }
}