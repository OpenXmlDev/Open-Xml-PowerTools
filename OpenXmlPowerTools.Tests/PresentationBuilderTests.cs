// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

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
using DocumentFormat.OpenXml.Validation;
using OpenXmlPowerTools;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class PbTests
    {
        [Fact]
        public void PB001_Formatting()
        {
            string name1 = "PB001-Input1.pptx";
            string name2 = "PB001-Input2.pptx";
            FileInfo source1Pptx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
            FileInfo source2Pptx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

            List<SlideSource> sources = null;
            sources = new List<SlideSource>()
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
            string name2 = "PB001-Input2.pptx";
           FileInfo source2Pptx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

            List<SlideSource> sources = null;
            sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(source2Pptx.FullName), 0, true),
            };
            var processedDestPptx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "PB002-Formatting.pptx"));
            PresentationBuilder.BuildPresentation(sources, processedDestPptx.FullName);
        }

        [Fact]
        public void PB003_Formatting()
        {
            string name1 = "PB001-Input1.pptx";
            string name2 = "PB001-Input3.pptx";
            FileInfo source1Pptx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
            FileInfo source2Pptx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

            List<SlideSource> sources = null;
            sources = new List<SlideSource>()
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
            string name1 = "PB001-Input1.pptx";
            string name2 = "PB001-Input3.pptx";
            FileInfo source1Pptx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
            FileInfo source2Pptx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

            List<SlideSource> sources = null;
            sources = new List<SlideSource>()
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
            string name1 = "PB001-Input1.pptx";
            string name2 = "PB001-Input3.pptx";
            FileInfo source1Pptx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
            FileInfo source2Pptx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

            List<SlideSource> sources = null;
            sources = new List<SlideSource>()
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
        [Fact]
#endif
        public void PB006_VideoFormats()
        {
            // This presentation contains videos with content types video/mp4, video/quicktime, video/unknown, video/x-ms-asf, and video/x-msvideo.
            FileInfo sourcePptx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, "PP006-Videos.pptx"));

            var oldMediaDataContentTypes = GetMediaDataContentTypes(sourcePptx);

            List<SlideSource> sources = null;
            sources = new List<SlideSource>()
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
            using (PresentationDocument ptDoc = PresentationDocument.Open(fi.FullName, false))
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

#endif
