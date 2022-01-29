using Codeuctivity.OpenXmlPowerTools;
using System.IO;
using Xunit;

namespace Codeuctivity.Tests
{
    public class PtUtilTests
    {
        [Theory(Skip = "This is failing on AppVeyor")]
        [InlineData("PU/PU001-Test001.mht")]
        public void PU001(string name)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceMht = new FileInfo(Path.Combine(sourceDir.FullName, name));
            var src = File.ReadAllText(sourceMht.FullName);
            var p = MhtParser.Parse(src);
            Assert.True(p.ContentType != null);
            Assert.True(p.MimeVersion != null);
            Assert.True(p.Parts.Length != 0);
            Assert.DoesNotContain(p.Parts, part => part.ContentType == null || part.ContentLocation == null);
        }
    }
}