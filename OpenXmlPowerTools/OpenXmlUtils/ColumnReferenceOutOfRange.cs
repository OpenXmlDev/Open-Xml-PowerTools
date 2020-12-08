using System;

namespace OpenXmlPowerTools
{
    public class ColumnReferenceOutOfRange : Exception
    {
        public ColumnReferenceOutOfRange(string columnReference)
            : base(string.Format("Column reference ({0}) is out of range.", columnReference))
        {
        }

        public ColumnReferenceOutOfRange()
        {
        }

        public ColumnReferenceOutOfRange(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}