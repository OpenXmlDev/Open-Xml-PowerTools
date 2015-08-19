using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

#if X64
namespace OpenXmlPowerTools.Tests.X64
#else
namespace OpenXmlPowerTools.Tests
#endif
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

#if false
    class TestUtil
    {
        public static DirectoryInfo SourceDir = new DirectoryInfo("../../../TestFiles/");
        public static DirectoryInfo TempDir = null;

        public static void TempDirSetup()
        {
            if (TempDir == null)
            {
                var homeDrive = Environment.GetEnvironmentVariable("HOMEDRIVE");
                var homePath = Environment.GetEnvironmentVariable("HOMEPATH");
                var now = DateTime.Now;
                var tempDirName = String.Format("OxPt-Test-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", now.Year - 2000, now.Month, now.Day, now.Hour, now.Minute, now.Second);
                TempDir = new DirectoryInfo(Path.Combine(homeDrive, homePath, "Documents", tempDirName));
                TempDir.Create();
            }
        }

    }
#endif
}
