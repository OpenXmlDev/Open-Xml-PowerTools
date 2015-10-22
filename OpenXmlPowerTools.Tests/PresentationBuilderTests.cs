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
using DocumentFormat.OpenXml.Validation;
using OpenXmlPowerTools;
using Xunit;

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

    }
}
