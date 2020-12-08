using OpenXmlPowerTools;
using System.IO;
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
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));

            var notAccepted = new WmlDocument(sourceDocx.FullName);
            var afterAccepting = RevisionAccepter.AcceptRevisions(notAccepted);
            var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-processed-by-RevisionAccepter.docx")));
            afterAccepting.SaveAs(processedDestDocx.FullName);
        }
    }
}