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
using System.Management.Automation;
using System.IO;
using System.Xml;
using System.Collections.ObjectModel;

namespace OpenXmlPowerTools.Commands
{
    public class PowerToolsReadOnlyCmdlet : PSCmdlet
    {
        private OpenXmlPowerToolsDocument[] documents;
        internal string[] fileNameReferences;

        #region Parameters

        /// <summary>
        /// Specify the Document parameter
        /// </summary>
        [Parameter(
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Specifies the documents to be processed.")]
        public OpenXmlPowerToolsDocument[] Document
        {
            get
            {
                return documents;
            }
            set
            {
                documents = value;
            }
        }

        /// <summary>
        /// Specify the Path parameter
        /// </summary>
        [Parameter(Position = 0,
            Mandatory = false,
            HelpMessage = "Specifies the path to the documents to be processed")]
        [ValidateNotNullOrEmpty]
        public string[] Path
        {
            set
            {
                fileNameReferences = value;
            }
        }
        #endregion

        internal IEnumerable<OpenXmlPowerToolsDocument> AllDocuments(string action)
        {
            if (fileNameReferences != null)
            {
                foreach (var path in fileNameReferences)
                {
                    Collection<PathInfo> fileList;
                    try
                    {
                        fileList = SessionState.Path.GetResolvedPSPathFromPSPath(path);
                    }
                    catch (ItemNotFoundException e)
                    {
                        WriteError(new ErrorRecord(e, "OpenXmlPowerToolsError", ErrorCategory.OpenError, path));
                        continue;
                    }
                    foreach (var file in fileList)
                    {
                        OpenXmlPowerToolsDocument document;
                        try
                        {
                            document = OpenXmlPowerToolsDocument.FromFileName(file.Path);
                        }
                        catch (Exception e)
                        {
                            WriteError(new ErrorRecord(e, "OpenXmlPowerToolsError", ErrorCategory.OpenError, file));
                            continue;
                        }
                        yield return document;
                    }
                }
            }
            else if (Document != null)
            {
                foreach (OpenXmlPowerToolsDocument document in Document)
                {
                    OpenXmlPowerToolsDocument specificDoc;
                    try
                    {
                        specificDoc = OpenXmlPowerToolsDocument.FromDocument(document);
                    }
                    catch (Exception e)
                    {
                        WriteError(new ErrorRecord(e, "OpenXmlPowerToolsError", ErrorCategory.InvalidType, document));
                        continue;
                    }
                    yield return specificDoc;
                }
            }
        }
    }

    public class PowerToolsModifierCmdlet : PSCmdlet
    {
        private OpenXmlPowerToolsDocument[] documents;
        internal string[] fileNameReferences;
        protected bool passThru = false;
        private string outputFolder;

        #region Parameters

        [Parameter(
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Specifies the documents to be processed.")]
        public OpenXmlPowerToolsDocument[] Document
        {
            get
            {
                return documents;
            }
            set
            {
                documents = value;
            }
        }

        [Parameter(Position = 0,
            Mandatory = false,
            HelpMessage = "Specifies the path to the documents to be processed")]
        [ValidateNotNullOrEmpty]
        public string[] Path
        {
            set
            {
                fileNameReferences = value;
            }
        }

        [Parameter(Position = 1,
            Mandatory = false,
            ValueFromPipeline = false,
            HelpMessage = "Path of folder to store result documents")
        ]
        public string OutputFolder
        {
            get
            {
                return outputFolder;
            }
            set
            {
                outputFolder = SessionState.Path.Combine(SessionState.Path.CurrentLocation.Path, value);
            }
        }

        [Parameter(
            Mandatory = false,
            HelpMessage = "Use this switch to pipe out the processed documents.")
        ]
        [ValidateNotNullOrEmpty]
        public SwitchParameter PassThru
        {
            get
            {
                return passThru;
            }
            set
            {
                passThru = value;
            }
        }

        #endregion

        internal IEnumerable<OpenXmlPowerToolsDocument> AllDocuments(string action)
        {
            if (fileNameReferences != null)
            {
                foreach (var path in fileNameReferences)
                {
                    Collection<PathInfo> fileList;
                    try
                    {
                        fileList = SessionState.Path.GetResolvedPSPathFromPSPath(path);
                    }
                    catch (ItemNotFoundException e)
                    {
                        WriteError(new ErrorRecord(e, "OpenXmlPowerToolsError", ErrorCategory.OpenError, path));
                        continue;
                    }
                    foreach (var file in fileList)
                    {
                        string target = file.Path;
                        if (OutputFolder != null)
                        {
                            FileInfo temp = new FileInfo(file.Path);
                            target = OutputFolder + "\\" + temp.Name;
                        }
                        if (!File.Exists(target) || ShouldProcess(target, action))
                        {
                            OpenXmlPowerToolsDocument document;
                            try
                            {
                                document = OpenXmlPowerToolsDocument.FromFileName(file.Path);
                            }
                            catch (Exception e)
                            {
                                WriteError(new ErrorRecord(e, "OpenXmlPowerToolsError", ErrorCategory.OpenError, file));
                                continue;
                            }
                            yield return document;
                        }
                    }
                }
            }
            else if (Document != null)
            {
                foreach (OpenXmlPowerToolsDocument document in Document)
                {
                    string target = document.FileName;
                    if (OutputFolder != null)
                    {
                        FileInfo temp = new FileInfo(document.FileName);
                        target = OutputFolder + "\\" + temp.Name;
                    }
                    if (!File.Exists(target) || ShouldProcess(target, action))
                    {
                        OpenXmlPowerToolsDocument specificDoc;
                        try
                        {
                            specificDoc = OpenXmlPowerToolsDocument.FromDocument(document);
                        }
                        catch (Exception e)
                        {
                            WriteError(new ErrorRecord(e, "OpenXmlPowerToolsError", ErrorCategory.InvalidType, document));
                            continue;
                        }
                        yield return specificDoc;
                    }
                }
            }
        }

        // Determines if and where to write the modified document
        internal void OutputDocument(OpenXmlPowerToolsDocument doc)
        {
            if (OutputFolder != null)
            {
                FileInfo file = new FileInfo(doc.FileName);
                string newName = OutputFolder + "\\" + file.Name;
                doc.SaveAs(newName);
            }
            else if (!PassThru)
                doc.Save();

            if (PassThru)
                WriteObject(doc, true);
        }
    }

    public class PowerToolsCreateCmdlet : PSCmdlet
    {
        protected bool passThru = false;

        #region Parameters

        /// <summary>
        /// PassThru parameter
        /// </summary>
        [Parameter(
            Mandatory = false,
            HelpMessage = "Use this switch to pipe out the processed documents.")
        ]
        [ValidateNotNullOrEmpty]
        public SwitchParameter PassThru
        {
            get
            {
                return passThru;
            }
            set
            {
                passThru = value;
            }
        }

        #endregion

        // Determines if and where to write the modified document
        internal void OutputDocument(OpenXmlPowerToolsDocument doc)
        {
            if (PassThru)
                WriteObject(doc, true);
            else
                doc.Save();
        }
    }

    public static class PowerToolsExceptionHandling
    {
        public static ErrorRecord GetExceptionErrorRecord(Exception e, OpenXmlPowerToolsDocument doc)
        {
            ErrorCategory cat = ErrorCategory.NotSpecified;
            if (e is ArgumentException)
                cat = ErrorCategory.InvalidArgument;
            else if (e is InvalidOperationException)
                cat = ErrorCategory.InvalidOperation;
            else if (e is PowerToolsDocumentException)
                cat = ErrorCategory.OpenError;
            else if (e is PowerToolsInvalidDataException || e is XmlException)
                cat = ErrorCategory.InvalidData;
            return new ErrorRecord(e, (cat == ErrorCategory.NotSpecified) ? "General" : "OpenXmlPowerToolsError", cat, doc);
        }
    }
    
    
}
