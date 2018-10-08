using System;

namespace OpenXmlPowerTools
{
    public class ColumnReferenceOutOfRange : Exception
    {
        public ColumnReferenceOutOfRange(string columnReference)
            : base($"Column reference ({columnReference}) is out of range.")
        {
        }
    }
}
