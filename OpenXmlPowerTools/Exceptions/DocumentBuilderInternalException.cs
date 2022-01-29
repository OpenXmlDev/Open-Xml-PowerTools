#define TestForUnsupportedDocuments
#define MergeStylesWithSameNames

using System;

namespace Codeuctivity.OpenXmlPowerTools.Exceptions
{
    public class DocumentBuilderInternalException : Exception
    {
        public DocumentBuilderInternalException(string message) : base(message)
        {
        }

        public DocumentBuilderInternalException(string message, Exception innerException) : base(message, innerException)
        {
        }

        public DocumentBuilderInternalException()
        {
        }
    }
}