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
using System.Xml;
using System.Xml.Linq;
using OpenXmlPowerTools;

/*******************************************************************************************
 * HtmlToWmlConverter expects the HTML to be passed as an XElement, i.e. as XML.  While the HTML test files that
 * are included in Open-Xml-PowerTools are able to be read as XML, most HTML is not able to be read as XML.
 * The best solution is to use the HtmlAgilityPack, which can parse HTML and save as XML.  The HtmlAgilityPack
 * is licensed under the Ms-PL (same as Open-Xml-PowerTools) so it is convenient to include it in your solution,
 * and thereby you can convert HTML to XML that can be processed by the HtmlToWmlConverter.
 * 
 * A convenient way to get the DLL that has been checked out with HtmlToWmlConverter is to clone the repo at
 * https://github.com/EricWhiteDev/HtmlAgilityPack
 * 
 * That repo contains only the DLL that has been checked out with HtmlToWmlConverter.
 * 
 * Of course, you can also get the HtmlAgilityPack source and compile it to get the DLL.  You can find it at
 * http://codeplex.com/HtmlAgilityPack
 * 
 * We don't include the HtmlAgilityPack in Open-Xml-PowerTools, to simplify installation.  The XUnit tests in
 * this module do not require the HtmlAgilityPack to run.
*******************************************************************************************/ 

#if USE_HTMLAGILITYPACK
using HtmlAgilityPack;
#endif

namespace OpenXmlPowerTools
{
    public class HtmlToWmlReadAsXElement
    {
        public static XElement ReadAsXElement(FileInfo sourceHtmlFi)
        {
            string htmlString = File.ReadAllText(sourceHtmlFi.FullName);
            XElement html = null;
            try
            {
                html = XElement.Parse(htmlString);
            }
#if USE_HTMLAGILITYPACK
            catch (XmlException)
            {
                HtmlDocument hdoc = new HtmlDocument();
                hdoc.Load(sourceHtmlFi.FullName, Encoding.Default);
                hdoc.OptionOutputAsXml = true;
                hdoc.Save(sourceHtmlFi.FullName, Encoding.Default);
                StringBuilder sb = new StringBuilder(File.ReadAllText(sourceHtmlFi.FullName, Encoding.Default));
                sb.Replace("&amp;", "&");
                sb.Replace("&nbsp;", "\xA0");
                sb.Replace("&quot;", "\"");
                sb.Replace("&lt;", "~lt;");
                sb.Replace("&gt;", "~gt;");
                sb.Replace("&#", "~#");
                sb.Replace("&", "&amp;");
                sb.Replace("~lt;", "&lt;");
                sb.Replace("~gt;", "&gt;");
                sb.Replace("~#", "&#");
                File.WriteAllText(sourceHtmlFi.FullName, sb.ToString(), Encoding.Default);
                html = XElement.Parse(sb.ToString());
            }
#else
            catch (XmlException e)
            {
                throw e;
            }
#endif
            html = (XElement)ConvertToNoNamespace(html);
            return html;
        }

        private static object ConvertToNoNamespace(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                return new XElement(element.Name.LocalName,
                    element.Attributes().Where(a => !a.IsNamespaceDeclaration),
                    element.Nodes().Select(n => ConvertToNoNamespace(n)));
            }
            return node;
        }
    }
}
