

using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace OxPt
{
    public class PbTests
    {
        [Fact]
        public void PB001_Formatting()
        {
            var name1 = "PB001-Input1.pptx";
            var name2 = "PB001-Input2.pptx";
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name1));
            var source2Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name2));
            var sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(source1Pptx.FullName), 1, true),
                new SlideSource(new PmlDocument(source2Pptx.FullName), 0, true),
            };
            var processedDestPptx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "PB001-Formatting.pptx"));
            PresentationBuilder.BuildPresentation(sources, processedDestPptx.FullName);
        }

        [Fact]
        public void PB002_Formatting()
        {
            var name2 = "PB001-Input2.pptx";
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source2Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name2));
            var sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(source2Pptx.FullName), 0, true),
            };
            var processedDestPptx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "PB002-Formatting.pptx"));
            PresentationBuilder.BuildPresentation(sources, processedDestPptx.FullName);
        }

        [Fact]
        public void PB003_Formatting()
        {
            var name1 = "PB001-Input1.pptx";
            var name2 = "PB001-Input3.pptx";
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name1));
            var source2Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name2));
            var sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(source1Pptx.FullName), 1, true),
                new SlideSource(new PmlDocument(source2Pptx.FullName), 0, true),
            };
            var processedDestPptx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "PB003-Formatting.pptx"));
            PresentationBuilder.BuildPresentation(sources, processedDestPptx.FullName);
        }

        [Fact]
        public void PB004_Formatting()
        {
            var name1 = "PB001-Input1.pptx";
            var name2 = "PB001-Input3.pptx";
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name1));
            var source2Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name2));
            var sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(source2Pptx.FullName), 0, true),
                new SlideSource(new PmlDocument(source1Pptx.FullName), 1, true),
            };
            var processedDestPptx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "PB004-Formatting.pptx"));
            PresentationBuilder.BuildPresentation(sources, processedDestPptx.FullName);
        }

        [Fact]
        public void PB005_Formatting()
        {
            var name1 = "PB001-Input1.pptx";
            var name2 = "PB001-Input3.pptx";
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name1));
            var source2Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name2));
            var sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(source2Pptx.FullName), 0, 0, true),
                new SlideSource(new PmlDocument(source1Pptx.FullName), 1, true),
                new SlideSource(new PmlDocument(source2Pptx.FullName), 0, true),
            };
            var processedDestPptx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "PB005-Formatting.pptx"));
            PresentationBuilder.BuildPresentation(sources, processedDestPptx.FullName);
        }

#if NETCOREAPP2_0
        [Fact(Skip="Bug in netcore 2.0 : https://github.com/OfficeDev/Open-Xml-PowerTools/pull/238#issuecomment-412375570")]
#else
#if NETCOREAPP3_1
        [Fact(Skip="Bug since netcore 2.0 : https://github.com/OfficeDev/Open-Xml-PowerTools/pull/238#issuecomment-412375570")]
#else

        [Fact]
#endif
#endif
        public void PB006_VideoFormats()
        {
            // This presentation contains videos with content types video/mp4, video/quicktime, video/unknown, video/x-ms-asf, and video/x-msvideo.
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourcePptx = new FileInfo(Path.Combine(sourceDir.FullName, "PP006-Videos.pptx"));

            var oldMediaDataContentTypes = GetMediaDataContentTypes(sourcePptx);
            var sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(sourcePptx.FullName), true),
            };
            var processedDestPptx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "PB006-Videos.pptx"));
            PresentationBuilder.BuildPresentation(sources, processedDestPptx.FullName);

            var newMediaDataContentTypes = GetMediaDataContentTypes(processedDestPptx);

            Assert.Equal(oldMediaDataContentTypes, newMediaDataContentTypes);
        }

        private static string[] GetMediaDataContentTypes(FileInfo fi)
        {
            using (var ptDoc = PresentationDocument.Open(fi.FullName, false))
            {
                return ptDoc.PresentationPart.SlideParts.SelectMany(
                        p => p.DataPartReferenceRelationships.Select(d => d.DataPart.ContentType))
                    .Distinct()
                    .OrderBy(m => m)
                    .ToArray();
            }
        }
    }
}