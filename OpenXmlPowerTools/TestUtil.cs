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
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXmlPowerTools
{
    public class TestUtil
    {
        public static DirectoryInfo SourceDir = new DirectoryInfo("../../../TestFiles/");
        private static bool? s_DeleteTempFiles = null;

        public static bool DeleteTempFiles
        {
            get
            {
                if (s_DeleteTempFiles != null)
                    return (bool)s_DeleteTempFiles;
                FileInfo donotdelete = new FileInfo("donotdelete.txt");
                s_DeleteTempFiles = !donotdelete.Exists;
                return (bool)s_DeleteTempFiles;
            }
        }

        private static DirectoryInfo s_TempDir = null;
        public static DirectoryInfo TempDir
        {
            get
            {
                if (s_TempDir != null)
                    return s_TempDir;
                else
                {
                    var now = DateTime.Now;
                    var tempDirName = String.Format("Test-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", now.Year - 2000, now.Month, now.Day, now.Hour, now.Minute, now.Second);
                    s_TempDir = new DirectoryInfo(Path.Combine(".", tempDirName));
                    s_TempDir.Create();
                    return s_TempDir;
                }
            }
        }
    }
}
