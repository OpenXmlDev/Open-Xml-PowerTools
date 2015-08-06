using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenXmlPowerTools;

namespace RevisionAccepterExample
{
    class RevisionAccepterExample
    {
        static void Main(string[] args)
        {
            // Accept all revisions, save result as a new document
            WmlDocument result = RevisionAccepter.AcceptRevisions(new WmlDocument("../../Source1.docx"));
            result.SaveAs("Out1.docx");
        }
    }
}
