using Codeuctivity.Exceptions;
using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace Codeuctivity
{
    public partial class SmlDocument : OpenXmlPowerToolsDocument
    {
        public SmlDocument(OpenXmlPowerToolsDocument original)
            : base(original)
        {
            if (GetDocumentType() != typeof(SpreadsheetDocument))
            {
                throw new PowerToolsDocumentException("Not a Spreadsheet document.");
            }
        }

        public SmlDocument(OpenXmlPowerToolsDocument original, bool convertToTransitional)
            : base(original, convertToTransitional)
        {
            if (GetDocumentType() != typeof(SpreadsheetDocument))
            {
                throw new PowerToolsDocumentException("Not a Spreadsheet document.");
            }
        }

        public SmlDocument(string fileName)
            : base(fileName)
        {
            if (GetDocumentType() != typeof(SpreadsheetDocument))
            {
                throw new PowerToolsDocumentException("Not a Spreadsheet document.");
            }
        }

        public SmlDocument(string fileName, bool convertToTransitional)
            : base(fileName, convertToTransitional)
        {
            if (GetDocumentType() != typeof(SpreadsheetDocument))
            {
                throw new PowerToolsDocumentException("Not a Spreadsheet document.");
            }
        }

        public SmlDocument(string fileName, byte[] byteArray)
            : base(byteArray)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(SpreadsheetDocument))
            {
                throw new PowerToolsDocumentException("Not a Spreadsheet document.");
            }
        }

        public SmlDocument(string fileName, byte[] byteArray, bool convertToTransitional)
            : base(byteArray, convertToTransitional)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(SpreadsheetDocument))
            {
                throw new PowerToolsDocumentException("Not a Spreadsheet document.");
            }
        }

        public SmlDocument(string fileName, MemoryStream memStream)
            : base(fileName, memStream)
        {
        }

        public SmlDocument(string fileName, MemoryStream memStream, bool convertToTransitional)
            : base(fileName, memStream, convertToTransitional)
        {
        }
    }
}