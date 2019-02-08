// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public static class PowerToolsBlockExtensions
    {
        /// <summary>
        /// Begins a PowerTools Block by (1) removing annotations and, unless the package was
        /// opened in read-only mode, (2) saving the package.
        /// </summary>
        /// <remarks>
        /// Removes <see cref="XDocument" /> and <see cref="XmlNamespaceManager" /> instances
        /// added by <see cref="PtOpenXmlExtensions.GetXDocument(OpenXmlPart)" />,
        /// <see cref="PtOpenXmlExtensions.GetXDocument(OpenXmlPart, out XmlNamespaceManager)" />,
        /// <see cref="PtOpenXmlExtensions.PutXDocument(OpenXmlPart)" />,
        /// <see cref="PtOpenXmlExtensions.PutXDocument(OpenXmlPart, XDocument)" />, and
        /// <see cref="PtOpenXmlExtensions.PutXDocumentWithFormatting(OpenXmlPart)" />.
        /// methods.
        /// </remarks>
        /// <param name="package">
        /// A <see cref="WordprocessingDocument" />, <see cref="SpreadsheetDocument" />,
        /// or <see cref="PresentationDocument" />.
        /// </param>
        public static void BeginPowerToolsBlock(this OpenXmlPackage package)
        {
            if (package == null) throw new ArgumentNullException("package");

            package.RemovePowerToolsAnnotations();
            package.Save();
        }

        /// <summary>
        /// Ends a PowerTools Block by reloading the root elements of all package parts
        /// that were changed by the PowerTools. A part is deemed changed by the PowerTools
        /// if it has an annotation of type <see cref="XDocument" />.
        /// </summary>
        /// <param name="package">
        /// A <see cref="WordprocessingDocument" />, <see cref="SpreadsheetDocument" />,
        /// or <see cref="PresentationDocument" />.
        /// </param>
        public static void EndPowerToolsBlock(this OpenXmlPackage package)
        {
            if (package == null) throw new ArgumentNullException("package");

            foreach (OpenXmlPart part in package.GetAllParts())
            {
                if (part.Annotations<XDocument>().Any() && part.RootElement != null)
                    part.RootElement.Reload();
            }
        }

        private static void RemovePowerToolsAnnotations(this OpenXmlPackage package)
        {
            if (package == null) throw new ArgumentNullException("package");

            foreach (OpenXmlPart part in package.GetAllParts())
            {
                part.RemoveAnnotations<XDocument>();
                part.RemoveAnnotations<XmlNamespaceManager>();
            }
        }
    }
}
