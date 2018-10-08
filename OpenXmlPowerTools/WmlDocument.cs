// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public class PtMainDocumentPart : XElement
    {
        private readonly WmlDocument _parentWmlDocument;

        public PtMainDocumentPart(WmlDocument wmlDocument, Uri uri, XName name, params object[] values)
            : base(name, values)
        {
            _parentWmlDocument = wmlDocument;
            Add(
                new XAttribute(PtOpenXml.Uri, uri),
                new XAttribute(XNamespace.Xmlns + "pt", PtOpenXml.pt)
            );
        }

        public PtWordprocessingCommentsPart WordprocessingCommentsPart
        {
            get
            {
                using (var ms = new MemoryStream(_parentWmlDocument.DocumentByteArray))
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, false))
                {
                    WordprocessingCommentsPart commentsPart = wDoc.MainDocumentPart.WordprocessingCommentsPart;
                    if (commentsPart == null) return null;

                    XElement partElement = commentsPart.GetXElement();
                    List<XNode> childNodes = partElement.Nodes().ToList();
                    foreach (XNode item in childNodes)
                    {
                        item.Remove();
                    }

                    return new PtWordprocessingCommentsPart(_parentWmlDocument, commentsPart.Uri, partElement.Name,
                        partElement.Attributes(), childNodes);
                }
            }
        }
    }

    public class PtWordprocessingCommentsPart : XElement
    {
        private WmlDocument _parentWmlDocument;

        public PtWordprocessingCommentsPart(WmlDocument wmlDocument, Uri uri, XName name, params object[] values)
            : base(name, values)
        {
            _parentWmlDocument = wmlDocument;
            Add(
                new XAttribute(PtOpenXml.Uri, uri),
                new XAttribute(XNamespace.Xmlns + "pt", PtOpenXml.pt)
            );
        }
    }

    public partial class WmlDocument
    {
        public WmlDocument(WmlDocument other, params XElement[] replacementParts)
            : base(other)
        {
            using (var streamDoc = new OpenXmlMemoryStreamDocument(this))
            {
                using (Package package = streamDoc.GetPackage())
                {
                    foreach (XElement replacementPart in replacementParts)
                    {
                        XAttribute uriAttribute = replacementPart.Attribute(PtOpenXml.Uri);
                        if (uriAttribute == null)
                            throw new OpenXmlPowerToolsException("Replacement part does not contain a Uri as an attribute");

                        string uri = uriAttribute.Value;
                        PackagePart part = package.GetParts().FirstOrDefault(p => p.Uri.ToString() == uri);
                        using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
                        using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
                            replacementPart.Save(partXmlWriter);
                    }
                }

                DocumentByteArray = streamDoc.GetModifiedDocument().DocumentByteArray;
            }
        }

        public PtMainDocumentPart MainDocumentPart
        {
            get
            {
                using (var ms = new MemoryStream(DocumentByteArray))
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, false))
                {
                    XElement partElement = wDoc.MainDocumentPart.GetXElement();
                    List<XNode> childNodes = partElement.Nodes().ToList();
                    foreach (XNode item in childNodes)
                    {
                        item.Remove();
                    }

                    return new PtMainDocumentPart(this, wDoc.MainDocumentPart.Uri, partElement.Name, partElement.Attributes(),
                        childNodes);
                }
            }
        }
    }
}
