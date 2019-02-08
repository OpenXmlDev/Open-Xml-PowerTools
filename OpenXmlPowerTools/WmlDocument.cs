// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public class PtMainDocumentPart : XElement
    {
        private WmlDocument ParentWmlDocument;

        public PtWordprocessingCommentsPart WordprocessingCommentsPart
        {
            get
            {
                using (MemoryStream ms = new MemoryStream(ParentWmlDocument.DocumentByteArray))
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, false))
                {
                    WordprocessingCommentsPart commentsPart = wDoc.MainDocumentPart.WordprocessingCommentsPart;
                    if (commentsPart == null)
                        return null;
                    XElement partElement = commentsPart.GetXDocument().Root;
                    var childNodes = partElement.Nodes().ToList();
                    foreach (var item in childNodes)
                        item.Remove();
                    return new PtWordprocessingCommentsPart(this.ParentWmlDocument, commentsPart.Uri, partElement.Name, partElement.Attributes(), childNodes);
                }
            }
        }

        public PtMainDocumentPart(WmlDocument wmlDocument, Uri uri, XName name, params object[] values)
            : base(name, values)
        {
            ParentWmlDocument = wmlDocument;
            this.Add(
                new XAttribute(PtOpenXml.Uri, uri),
                new XAttribute(XNamespace.Xmlns + "pt", PtOpenXml.pt)
            );
        }
    }

    public class PtWordprocessingCommentsPart : XElement
    {
        private WmlDocument ParentWmlDocument;

        public PtWordprocessingCommentsPart(WmlDocument wmlDocument, Uri uri, XName name, params object[] values)
            : base(name, values)
        {
            ParentWmlDocument = wmlDocument;
            this.Add(
                new XAttribute(PtOpenXml.Uri, uri),
                new XAttribute(XNamespace.Xmlns + "pt", PtOpenXml.pt)
            );
        }
    }

    public partial class WmlDocument
    {
        public PtMainDocumentPart MainDocumentPart
        {
            get
            {
                using (MemoryStream ms = new MemoryStream(this.DocumentByteArray))
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, false))
                {
                    XElement partElement = wDoc.MainDocumentPart.GetXDocument().Root;
                    var childNodes = partElement.Nodes().ToList();
                    foreach (var item in childNodes)
                        item.Remove();
                    return new PtMainDocumentPart(this, wDoc.MainDocumentPart.Uri, partElement.Name, partElement.Attributes(), childNodes);
                }
            }
        }

        public WmlDocument(WmlDocument other, params XElement[] replacementParts)
            : base(other)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(this))
            {
                using (Package package = streamDoc.GetPackage())
                {
                    foreach (var replacementPart in replacementParts)
                    {
                        XAttribute uriAttribute = replacementPart.Attribute(PtOpenXml.Uri);
                        if (uriAttribute == null)
                            throw new OpenXmlPowerToolsException("Replacement part does not contain a Uri as an attribute");
                        String uri = uriAttribute.Value;
                        var part = package.GetParts().FirstOrDefault(p => p.Uri.ToString() == uri);
                        using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
                        using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
                            replacementPart.Save(partXmlWriter);
                    }
                }
                this.DocumentByteArray = streamDoc.GetModifiedDocument().DocumentByteArray;
            }
        }
    }
}
