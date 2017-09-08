//
// Copyright 2017 Thomas Barnekow
//
// This code is licensed using the Microsoft Public License (Ms-PL). The text of the
// license can be found here:
//
// http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx
//
// Developer: Thomas Barnekow
// Email: thomas@barnekow.info
//

using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.Document = new Document(new Body());
            }
        }
    }
}
