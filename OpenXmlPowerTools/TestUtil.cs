// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Diagnostics;
using System.IO;

namespace OpenXmlPowerTools
{
    public static class TestUtil
    {
        public static readonly DirectoryInfo SourceDir = new DirectoryInfo("../../../../TestFiles/");

        private static bool? _deleteTempFiles;

        private static DirectoryInfo _tempDir;

        public static bool DeleteTempFiles
        {
            get
            {
                if (_deleteTempFiles != null) return (bool) _deleteTempFiles;

                var donotdelete = new FileInfo("donotdelete.txt");
                _deleteTempFiles = !donotdelete.Exists;

                return (bool) _deleteTempFiles;
            }
        }

        public static DirectoryInfo TempDir
        {
            get
            {
                if (_tempDir != null) return _tempDir;

                DateTime now = DateTime.Now;
                string tempDirName =
                    $"Test-{now.Year - 2000:00}-{now.Month:00}-{now.Day:00}-{now.Hour:00}{now.Minute:00}{now.Second:00}";

                _tempDir = new DirectoryInfo(Path.Combine(".", tempDirName));
                _tempDir.Create();

                return _tempDir;
            }
        }

        public static void NotePad(string str)
        {
            string guidName = Guid.NewGuid().ToString().Replace("-", "") + ".txt";
            var fi = new FileInfo(Path.Combine(TempDir.FullName, guidName));
            File.WriteAllText(fi.FullName, str);

            var notepadExe = new FileInfo(@"C:\Program Files (x86)\Notepad++\notepad++.exe");
            if (!notepadExe.Exists)
            {
                notepadExe = new FileInfo(@"C:\Program Files\Notepad++\notepad++.exe");
            }

            if (!notepadExe.Exists)
            {
                notepadExe = new FileInfo(@"C:\Windows\System32\notepad.exe");
            }

            ExecutableRunner.RunExecutable(notepadExe.FullName, fi.FullName, TempDir.FullName);
        }

        public static void KDiff3(FileInfo oldFi, FileInfo newFi)
        {
            var kdiffExe = new FileInfo(@"C:\Program Files (x86)\KDiff3\kdiff3.exe");
            ExecutableRunner.RunResults unused =
                ExecutableRunner.RunExecutable(kdiffExe.FullName, oldFi.FullName + " " + newFi.FullName, TempDir.FullName);
        }

        public static void Explorer(DirectoryInfo di)
        {
            Process.Start(di.FullName);
        }
    }
}
