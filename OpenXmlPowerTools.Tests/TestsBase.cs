// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace OpenXmlPowerTools.Tests
{
    /// <summary>
    /// Base class for unit tests providing utility methods.
    /// </summary>
    public class TestsBase
    {
        private const WordprocessingDocumentType DocumentType = WordprocessingDocumentType.Document;

        protected static void CreateEmptyWordprocessingDocument(Stream stream)
        {
            using (var wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                var part = wordDocument.AddMainDocumentPart();
                part.Document = new Document(new Body());
            }
        }
    }
}