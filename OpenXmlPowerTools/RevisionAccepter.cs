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
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public class RevisionAccepter
    {
        public static WmlDocument AcceptRevisions(WmlDocument document)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(document))
            {
                using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                {
                    AcceptRevisions(doc);
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }

        public static void AcceptRevisions(WordprocessingDocument doc)
        {
            RevisionProcessor.AcceptRevisions(doc);
        }

        public static bool PartHasTrackedRevisions(OpenXmlPart part)
        {
            return RevisionProcessor.PartHasTrackedRevisions(part);
        }

        public static bool HasTrackedRevisions(WmlDocument document)
        {
            return RevisionProcessor.HasTrackedRevisions(document);
        }

        public static bool HasTrackedRevisions(WordprocessingDocument doc)
        {
            return RevisionProcessor.HasTrackedRevisions(doc);
        }
    }
}
