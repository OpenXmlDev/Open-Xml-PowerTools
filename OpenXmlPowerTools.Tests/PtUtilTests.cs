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
using OpenXmlPowerTools;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class PtUtilTests
    {
        [Theory(Skip = "This is failing on AppVeyor")]
        [InlineData("PU/PU001-Test001.mht")]
        public void PU001(string name)
        {
            FileInfo sourceMht = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));
            var src = File.ReadAllText(sourceMht.FullName);
            var p = MhtParser.Parse(src);
            Assert.True(p.ContentType != null);
            Assert.True(p.MimeVersion != null);
            Assert.True(p.Parts.Length != 0);
            Assert.DoesNotContain(p.Parts, part => part.ContentType == null || part.ContentLocation == null);
        }

    }
}

#endif
