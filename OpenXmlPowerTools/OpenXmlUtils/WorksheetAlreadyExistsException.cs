using System;

namespace OpenXmlPowerTools
{
    public class WorksheetAlreadyExistsException : Exception
    {
        public WorksheetAlreadyExistsException(string sheetName)
            : base(string.Format("The worksheet ({0}) already exists.", sheetName))
        {
        }

        public WorksheetAlreadyExistsException(string message, Exception innerException) : base(message, innerException)
        {
        }

        public WorksheetAlreadyExistsException()
        {
        }
    }
}