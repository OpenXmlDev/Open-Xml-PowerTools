using System;

namespace OpenXmlPowerTools
{
    public class WorksheetAlreadyExistsException : Exception
    {
        public WorksheetAlreadyExistsException(string sheetName)
            : base($"The worksheet ({sheetName}) already exists.")
        {
        }
    }
}
