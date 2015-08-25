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

#if X64
namespace OpenXmlPowerTools.Tests.X64
#else
namespace OpenXmlPowerTools.Tests
#endif
{
    public class RevisionAccepterTests
    {
        [Theory]
        [InlineData("RA001-Tracked-Revisions-01.docx")]
        [InlineData("RA001-Tracked-Revisions-02.docx")]

        public void RA001_RevisionAccepter(string name)
        {
            FileInfo sourceDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            WmlDocument notAccepted = new WmlDocument(sourceDocx.FullName);
            WmlDocument afterAccepting = RevisionAccepter.AcceptRevisions(notAccepted);
            var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-processed-by-RevisionAccepter.docx")));
            afterAccepting.SaveAs(processedDestDocx.FullName);
        }

    }
}
