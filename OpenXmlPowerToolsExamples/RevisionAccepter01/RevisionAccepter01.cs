// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using OpenXmlPowerTools;
using System;
using System.IO;

namespace RevisionAccepterExample
{
    internal class RevisionAccepterExample
    {
        private static void Main(string[] args)
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            // Accept all revisions, save result as a new document
            var result = RevisionAccepter.AcceptRevisions(new WmlDocument("../../Source1.docx"));
            result.SaveAs(Path.Combine(tempDi.FullName, "Out1.docx"));
        }
    }
}