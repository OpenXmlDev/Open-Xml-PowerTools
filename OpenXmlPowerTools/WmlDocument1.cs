using Codeuctivity.OpenXmlPowerTools.Exceptions;
using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace Codeuctivity.OpenXmlPowerTools
{
    public partial class WmlDocument : OpenXmlPowerToolsDocument
    {
        public WmlDocument(OpenXmlPowerToolsDocument original)
            : base(original)
        {
            if (GetDocumentType() != typeof(WordprocessingDocument))
            {
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
            }
        }

        public WmlDocument(OpenXmlPowerToolsDocument original, bool convertToTransitional)
            : base(original, convertToTransitional)
        {
            if (GetDocumentType() != typeof(WordprocessingDocument))
            {
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
            }
        }

        public WmlDocument(string fileName)
            : base(fileName)
        {
            if (GetDocumentType() != typeof(WordprocessingDocument))
            {
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
            }
        }

        public WmlDocument(string fileName, bool convertToTransitional)
            : base(fileName, convertToTransitional)
        {
            if (GetDocumentType() != typeof(WordprocessingDocument))
            {
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
            }
        }

        public WmlDocument(string fileName, byte[] byteArray)
            : base(byteArray)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(WordprocessingDocument))
            {
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
            }
        }

        public WmlDocument(string fileName, byte[] byteArray, bool convertToTransitional)
            : base(byteArray, convertToTransitional)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(WordprocessingDocument))
            {
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
            }
        }

        public WmlDocument(string fileName, MemoryStream memStream)
            : base(fileName, memStream)
        {
        }

        public WmlDocument(string fileName, MemoryStream memStream, bool convertToTransitional)
            : base(fileName, memStream, convertToTransitional)
        {
        }
    }
}