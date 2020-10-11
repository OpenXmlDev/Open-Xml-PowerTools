

#define TestForUnsupportedDocuments
#define MergeStylesWithSameNames

using System;

namespace OpenXmlPowerTools
{
    public class DocumentBuilderException : Exception
    {
        public DocumentBuilderException(string message) : base(message)
        {
        }

        public DocumentBuilderException(string message, Exception innerException) : base(message, innerException)
        {
        }

        public DocumentBuilderException()
        {
        }
    }
}