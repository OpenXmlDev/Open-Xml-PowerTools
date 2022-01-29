using System;
using System.IO;

namespace Codeuctivity.Tests
{
    public class TestUtil
    {
        static TestUtil()
        {
            var now = DateTime.Now;
            var tempDirName = $"Test-{now.Year - 2000:00}-{now.Month:00}-{now.Day:00}-{now.Hour:00}{now.Minute:00}{now.Second:00}-{Guid.NewGuid()}";
            var tempDir = new DirectoryInfo(Path.Combine(Path.GetTempPath(), tempDirName));
            tempDir.Create();
            TempDir = tempDir;
        }

        /// <summary>
        /// Lookin into /tmp or %temp% for test output
        /// </summary>
        public static DirectoryInfo TempDir { get; private set; }
    }
}