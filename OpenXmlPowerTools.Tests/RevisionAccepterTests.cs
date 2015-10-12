/***************************************************************************

Copyright (c) Microsoft Corporation 2012-2015.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

***************************************************************************/

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using Xunit;

namespace OxPt
{
    public class RaTests
    {
        [Theory]
        [InlineData("RA001-Tracked-Revisions-01.docx")]
        [InlineData("RA001-Tracked-Revisions-02.docx")]

        public void RA001(string name)
        {
            FileInfo sourceDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            WmlDocument notAccepted = new WmlDocument(sourceDocx.FullName);
            WmlDocument afterAccepting = RevisionAccepter.AcceptRevisions(notAccepted);
            var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-processed-by-RevisionAccepter.docx")));
            afterAccepting.SaveAs(processedDestDocx.FullName);
        }

    }
}
