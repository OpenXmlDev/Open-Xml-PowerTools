// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;

namespace OpenXmlPowerTools
{
    public static class PtUtils
    {
        public static string SHA1HashStringForUTF8String(string s)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(s);
            var sha1 = SHA1.Create();
            byte[] hashBytes = sha1.ComputeHash(bytes);
            return HexStringFromBytes(hashBytes);
        }

        public static string SHA1HashStringForByteArray(byte[] bytes)
        {
            var sha1 = SHA1.Create();
            byte[] hashBytes = sha1.ComputeHash(bytes);
            return HexStringFromBytes(hashBytes);
        }

        public static string HexStringFromBytes(byte[] bytes)
        {
            var sb = new StringBuilder();
            foreach (byte b in bytes)
            {
                var hex = b.ToString("x2");
                sb.Append(hex);
            }
            return sb.ToString();
        }

        public static string NormalizeDirName(string dirName)
        {
            string d = dirName.Replace('\\', '/');
            if (d[dirName.Length - 1] != '/' && d[dirName.Length - 1] != '\\')
                return d + "/";

            return d;
        }

        public static string MakeValidXml(string p)
        {
            return p.Any(c => c < 0x20)
                ? p.Select(c => c < 0x20 ? string.Format("_{0:X}_", (int) c) : c.ToString()).StringConcatenate()
                : p;
        }

        public static void AddElementIfMissing(XDocument partXDoc, XElement existing, string newElement)
        {
            if (existing != null)
                return;

            XElement newXElement = XElement.Parse(newElement);
            newXElement.Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
            if (partXDoc.Root != null) partXDoc.Root.Add(newXElement);
        }
    }

    public class MhtParser
    {
        public string MimeVersion;
        public string ContentType;
        public MhtParserPart[] Parts;

        public class MhtParserPart
        {
            public string ContentLocation;
            public string ContentTransferEncoding;
            public string ContentType;
            public string CharSet;
            public string Text;
            public byte[] Binary;
        }

        public static MhtParser Parse(string src)
        {
            string mimeVersion = null;
            string contentType = null;
            string boundary = null;

            string[] lines = src.Split(new[] { Environment.NewLine }, StringSplitOptions.None);


            var priambleKeyWords = new[]
            {
                "MIME-VERSION:",
                "CONTENT-TYPE:",
            };

            var priamble = lines.TakeWhile(l =>
            {
                var s = l.ToUpper();
                return priambleKeyWords.Any(pk => s.StartsWith(pk));
            }).ToArray();

            foreach (var item in priamble)
            {
                if (item.ToUpper().StartsWith("MIME-VERSION:"))
                    mimeVersion = item.Substring("MIME-VERSION:".Length).Trim();
                else if (item.ToUpper().StartsWith("CONTENT-TYPE:"))
                {
                    var contentTypeLine = item.Substring("CONTENT-TYPE:".Length).Trim();
                    var spl = contentTypeLine.Split(';').Select(z => z.Trim()).ToArray();
                    foreach (var s in spl)
                    {
                        if (s.StartsWith("boundary"))
                        {
                            var begText = "boundary=\"";
                            var begLen = begText.Length;
                            boundary = s.Substring(begLen, s.Length - begLen - 1).TrimStart('-');
                            continue;
                        }
                        if (contentType == null)
                        {
                            contentType = s;
                            continue;
                        }
                        throw new OpenXmlPowerToolsException("Unexpected content in MHTML");
                    }
                }
            }

            var grouped = lines
                .Skip(priamble.Length)
                .GroupAdjacent(l =>
                {
                    var b = l.TrimStart('-') == boundary;
                    return b;
                })
                .Where(g => g.Key == false)
                .ToArray();

            var parts = grouped.Select(rp =>
                {
                    var partPriambleKeyWords = new[]
                    {
                        "CONTENT-LOCATION:",
                        "CONTENT-TRANSFER-ENCODING:",
                        "CONTENT-TYPE:",
                    };

                    var partPriamble = rp.TakeWhile(l =>
                    {
                        var s = l.ToUpper();
                        return partPriambleKeyWords.Any(pk => s.StartsWith(pk));
                    }).ToArray();

                    string contentLocation = null;
                    string contentTransferEncoding = null;
                    string partContentType = null;
                    string partCharSet = null;
                    byte[] partBinary = null;

                    foreach (var item in partPriamble)
                    {
                        if (item.ToUpper().StartsWith("CONTENT-LOCATION:"))
                            contentLocation = item.Substring("CONTENT-LOCATION:".Length).Trim();
                        else if (item.ToUpper().StartsWith("CONTENT-TRANSFER-ENCODING:"))
                            contentTransferEncoding = item.Substring("CONTENT-TRANSFER-ENCODING:".Length).Trim();
                        else if (item.ToUpper().StartsWith("CONTENT-TYPE:"))
                            partContentType = item.Substring("CONTENT-TYPE:".Length).Trim();
                    }

                    var blankLinesAtBeginning = rp
                        .Skip(partPriamble.Length)
                        .TakeWhile(l => l == "")
                        .Count();

                    var partText = rp
                        .Skip(partPriamble.Length)
                        .Skip(blankLinesAtBeginning)
                        .Select(l => l + Environment.NewLine)
                        .StringConcatenate();

                    if (partContentType != null && partContentType.Contains(";"))
                    {
                        string thisPartContentType = null;
                        var spl = partContentType.Split(';').Select(s => s.Trim()).ToArray();
                        foreach (var s in spl)
                        {
                            if (s.StartsWith("charset"))
                            {
                                var begText = "charset=\"";
                                var begLen = begText.Length;
                                partCharSet = s.Substring(begLen, s.Length - begLen - 1);
                                continue;
                            }
                            if (thisPartContentType == null)
                            {
                                thisPartContentType = s;
                                continue;
                            }
                            throw new OpenXmlPowerToolsException("Unexpected content in MHTML");
                        }
                        partContentType = thisPartContentType;
                    }

                    if (contentTransferEncoding != null && contentTransferEncoding.ToUpper() == "BASE64")
                    {
                        partBinary = Convert.FromBase64String(partText);
                    }

                    return new MhtParserPart()
                    {
                        ContentLocation = contentLocation,
                        ContentTransferEncoding = contentTransferEncoding,
                        ContentType = partContentType,
                        CharSet = partCharSet,
                        Text = partText,
                        Binary = partBinary,
                    };
                })
                .Where(p => p.ContentType != null)
                .ToArray();

            return new MhtParser()
            {
                ContentType = contentType,
                MimeVersion = mimeVersion,
                Parts = parts,
            };
        }
    }

    public class Normalizer
    {
        public static XDocument Normalize(XDocument source, XmlSchemaSet schema)
        {
            bool havePSVI = false;
            // validate, throw errors, add PSVI information
            if (schema != null)
            {
                source.Validate(schema, null, true);
                havePSVI = true;
            }
            return new XDocument(
                source.Declaration,
                source.Nodes().Select(n =>
                {
                    // Remove comments, processing instructions, and text nodes that are
                    // children of XDocument.  Only white space text nodes are allowed as
                    // children of a document, so we can remove all text nodes.
                    if (n is XComment || n is XProcessingInstruction || n is XText)
                        return null;
                    XElement e = n as XElement;
                    if (e != null)
                        return NormalizeElement(e, havePSVI);
                    return n;
                }
                )
            );
        }

        public static bool DeepEqualsWithNormalization(XDocument doc1, XDocument doc2,
            XmlSchemaSet schemaSet)
        {
            XDocument d1 = Normalize(doc1, schemaSet);
            XDocument d2 = Normalize(doc2, schemaSet);
            return XNode.DeepEquals(d1, d2);
        }

        private static IEnumerable<XAttribute> NormalizeAttributes(XElement element,
            bool havePSVI)
        {
            return element.Attributes()
                    .Where(a => !a.IsNamespaceDeclaration &&
                        a.Name != Xsi.schemaLocation &&
                        a.Name != Xsi.noNamespaceSchemaLocation)
                    .OrderBy(a => a.Name.NamespaceName)
                    .ThenBy(a => a.Name.LocalName)
                    .Select(
                        a =>
                        {
                            if (havePSVI)
                            {
                                var dt = a.GetSchemaInfo().SchemaType.TypeCode;
                                switch (dt)
                                {
                                    case XmlTypeCode.Boolean:
                                        return new XAttribute(a.Name, (bool)a);
                                    case XmlTypeCode.DateTime:
                                        return new XAttribute(a.Name, (DateTime)a);
                                    case XmlTypeCode.Decimal:
                                        return new XAttribute(a.Name, (decimal)a);
                                    case XmlTypeCode.Double:
                                        return new XAttribute(a.Name, (double)a);
                                    case XmlTypeCode.Float:
                                        return new XAttribute(a.Name, (float)a);
                                    case XmlTypeCode.HexBinary:
                                    case XmlTypeCode.Language:
                                        return new XAttribute(a.Name,
                                            ((string)a).ToLower());
                                }
                            }
                            return a;
                        }
                    );
        }

        private static XNode NormalizeNode(XNode node, bool havePSVI)
        {
            // trim comments and processing instructions from normalized tree
            if (node is XComment || node is XProcessingInstruction)
                return null;
            XElement e = node as XElement;
            if (e != null)
                return NormalizeElement(e, havePSVI);
            // Only thing left is XCData and XText, so clone them
            return node;
        }

        private static XElement NormalizeElement(XElement element, bool havePSVI)
        {
            if (havePSVI)
            {
                var dt = element.GetSchemaInfo();
                switch (dt.SchemaType.TypeCode)
                {
                    case XmlTypeCode.Boolean:
                        return new XElement(element.Name,
                            NormalizeAttributes(element, havePSVI),
                            (bool)element);
                    case XmlTypeCode.DateTime:
                        return new XElement(element.Name,
                            NormalizeAttributes(element, havePSVI),
                            (DateTime)element);
                    case XmlTypeCode.Decimal:
                        return new XElement(element.Name,
                            NormalizeAttributes(element, havePSVI),
                            (decimal)element);
                    case XmlTypeCode.Double:
                        return new XElement(element.Name,
                            NormalizeAttributes(element, havePSVI),
                            (double)element);
                    case XmlTypeCode.Float:
                        return new XElement(element.Name,
                            NormalizeAttributes(element, havePSVI),
                            (float)element);
                    case XmlTypeCode.HexBinary:
                    case XmlTypeCode.Language:
                        return new XElement(element.Name,
                            NormalizeAttributes(element, havePSVI),
                            ((string)element).ToLower());
                    default:
                        return new XElement(element.Name,
                            NormalizeAttributes(element, havePSVI),
                            element.Nodes().Select(n => NormalizeNode(n, havePSVI))
                        );
                }
            }
            else
            {
                return new XElement(element.Name,
                    NormalizeAttributes(element, havePSVI),
                    element.Nodes().Select(n => NormalizeNode(n, havePSVI))
                );
            }
        }
    }

    public class FileUtils
    {
        public static DirectoryInfo GetDateTimeStampedDirectoryInfo(string prefix)
        {
            DateTime now = DateTime.Now;
            string dirName =
                prefix +
                string.Format("-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", now.Year - 2000, now.Month, now.Day, now.Hour,
                    now.Minute, now.Second);
            return new DirectoryInfo(dirName);
        }

        public static FileInfo GetDateTimeStampedFileInfo(string prefix, string suffix)
        {
            DateTime now = DateTime.Now;
            string fileName =
                prefix +
                string.Format("-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", now.Year - 2000, now.Month, now.Day, now.Hour,
                    now.Minute, now.Second) +
                suffix;
            return new FileInfo(fileName);
        }

        public static void ThreadSafeCreateDirectory(DirectoryInfo dir)
        {
            while (true)
            {
                if (dir.Exists)
                    break;

                try
                {
                    dir.Create();
                    break;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                }
            }
        }

        public static void ThreadSafeCopy(FileInfo sourceFile, FileInfo destFile)
        {
            while (true)
            {
                if (destFile.Exists)
                    break;

                try
                {
                    File.Copy(sourceFile.FullName, destFile.FullName);
                    break;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                }
            }
        }

        public static void ThreadSafeCreateEmptyTextFileIfNotExist(FileInfo file)
        {
            while (true)
            {
                if (file.Exists)
                    break;

                try
                {
                    File.WriteAllText(file.FullName, "");
                    break;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                }
            }
        }

#if !NET35
        internal static void ThreadSafeAppendAllLines(FileInfo file, string[] strings)
        {
            while (true)
            {
                try
                {
                    File.AppendAllLines(file.FullName, strings);
                    break;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                }
            }
        }
#endif

        public static List<string> GetFilesRecursive(DirectoryInfo dir, string searchPattern)
        {
            var fileList = new List<string>();
            GetFilesRecursiveInternal(dir, searchPattern, fileList);
            return fileList;
        }

        private static void GetFilesRecursiveInternal(DirectoryInfo dir, string searchPattern, List<string> fileList)
        {
            fileList.AddRange(dir.GetFiles(searchPattern).Select(file => file.FullName));
            foreach (DirectoryInfo subdir in dir.GetDirectories())
                GetFilesRecursiveInternal(subdir, searchPattern, fileList);
        }

        public static List<string> GetFilesRecursive(DirectoryInfo dir)
        {
            var fileList = new List<string>();
            GetFilesRecursiveInternal(dir, fileList);
            return fileList;
        }

        private static void GetFilesRecursiveInternal(DirectoryInfo dir, List<string> fileList)
        {
            fileList.AddRange(dir.GetFiles().Select(file => file.FullName));
            foreach (DirectoryInfo subdir in dir.GetDirectories())
                GetFilesRecursiveInternal(subdir, fileList);
        }

        public static void CopyStream(Stream source, Stream target)
        {
            const int bufSize = 0x4096;
            var buf = new byte[bufSize];
            int bytesRead;
            while ((bytesRead = source.Read(buf, 0, bufSize)) > 0)
                target.Write(buf, 0, bytesRead);
        }
    }

    public static class PtExtensions
    {
        public static XElement GetXElement(this XmlNode node)
        {
            var xDoc = new XDocument();
            using (XmlWriter xmlWriter = xDoc.CreateWriter())
                node.WriteTo(xmlWriter);
            return xDoc.Root;
        }

        public static XmlNode GetXmlNode(this XElement element)
        {
            var xmlDoc = new XmlDocument();
            using (XmlReader xmlReader = element.CreateReader())
                xmlDoc.Load(xmlReader);
            return xmlDoc;
        }

        public static XDocument GetXDocument(this XmlDocument document)
        {
            var xDoc = new XDocument();
            using (XmlWriter xmlWriter = xDoc.CreateWriter())
                document.WriteTo(xmlWriter);

            XmlDeclaration decl = document.ChildNodes.OfType<XmlDeclaration>().FirstOrDefault();
            if (decl != null)
                xDoc.Declaration = new XDeclaration(decl.Version, decl.Encoding, decl.Standalone);

            return xDoc;
        }

        public static XmlDocument GetXmlDocument(this XDocument document)
        {
            var xmlDoc = new XmlDocument();
            using (XmlReader xmlReader = document.CreateReader())
            {
                xmlDoc.Load(xmlReader);
                if (document.Declaration != null)
                {
                    XmlDeclaration dec = xmlDoc.CreateXmlDeclaration(document.Declaration.Version,
                        document.Declaration.Encoding, document.Declaration.Standalone);
                    xmlDoc.InsertBefore(dec, xmlDoc.FirstChild);
                }
            }
            return xmlDoc;
        }

        public static string StringConcatenate(this IEnumerable<string> source)
        {
            return source.Aggregate(
                new StringBuilder(),
                (sb, s) => sb.Append(s),
                sb => sb.ToString());
        }

        public static string StringConcatenate<T>(this IEnumerable<T> source, Func<T, string> projectionFunc)
        {
            return source.Aggregate(
                new StringBuilder(),
                (sb, i) => sb.Append(projectionFunc(i)),
                sb => sb.ToString());
        }

        public static IEnumerable<TResult> PtZip<TFirst, TSecond, TResult>(
            this IEnumerable<TFirst> first,
            IEnumerable<TSecond> second,
            Func<TFirst, TSecond, TResult> func)
        {
            using (IEnumerator<TFirst> ie1 = first.GetEnumerator())
            using (IEnumerator<TSecond> ie2 = second.GetEnumerator())
                while (ie1.MoveNext() && ie2.MoveNext())
                    yield return func(ie1.Current, ie2.Current);
        }

        public static IEnumerable<IGrouping<TKey, TSource>> GroupAdjacent<TSource, TKey>(
            this IEnumerable<TSource> source,
            Func<TSource, TKey> keySelector)
        {
            TKey last = default(TKey);
            var haveLast = false;
            var list = new List<TSource>();

            foreach (TSource s in source)
            {
                TKey k = keySelector(s);
                if (haveLast)
                {
                    if (!k.Equals(last))
                    {
                        yield return new GroupOfAdjacent<TSource, TKey>(list, last);

                        list = new List<TSource> { s };
                        last = k;
                    }
                    else
                    {
                        list.Add(s);
                        last = k;
                    }
                }
                else
                {
                    list.Add(s);
                    last = k;
                    haveLast = true;
                }
            }
            if (haveLast)
                yield return new GroupOfAdjacent<TSource, TKey>(list, last);
        }

        private static void InitializeSiblingsReverseDocumentOrder(XElement element)
        {
            XElement prev = null;
            foreach (XElement e in element.Elements())
            {
                e.AddAnnotation(new SiblingsReverseDocumentOrderInfo { PreviousSibling = prev });
                prev = e;
            }
        }

        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        public static IEnumerable<XElement> SiblingsBeforeSelfReverseDocumentOrder(
            this XElement element)
        {
            if (element.Annotation<SiblingsReverseDocumentOrderInfo>() == null)
                InitializeSiblingsReverseDocumentOrder(element.Parent);
            XElement current = element;
            while (true)
            {
                XElement previousElement = current
                    .Annotation<SiblingsReverseDocumentOrderInfo>()
                    .PreviousSibling;
                if (previousElement == null)
                    yield break;

                yield return previousElement;

                current = previousElement;
            }
        }

        private static void InitializeDescendantsReverseDocumentOrder(XElement element)
        {
            XElement prev = null;
            foreach (XElement e in element.Descendants())
            {
                e.AddAnnotation(new DescendantsReverseDocumentOrderInfo { PreviousElement = prev });
                prev = e;
            }
        }

        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        public static IEnumerable<XElement> DescendantsBeforeSelfReverseDocumentOrder(
            this XElement element)
        {
            if (element.Annotation<DescendantsReverseDocumentOrderInfo>() == null)
                InitializeDescendantsReverseDocumentOrder(element.AncestorsAndSelf().Last());
            XElement current = element;
            while (true)
            {
                XElement previousElement = current
                    .Annotation<DescendantsReverseDocumentOrderInfo>()
                    .PreviousElement;
                if (previousElement == null)
                    yield break;

                yield return previousElement;

                current = previousElement;
            }
        }

        private static void InitializeDescendantsTrimmedReverseDocumentOrder(XElement element, XName trimName)
        {
            XElement prev = null;
            foreach (XElement e in element.DescendantsTrimmed(trimName))
            {
                e.AddAnnotation(new DescendantsTrimmedReverseDocumentOrderInfo { PreviousElement = prev });
                prev = e;
            }
        }

        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        public static IEnumerable<XElement> DescendantsTrimmedBeforeSelfReverseDocumentOrder(
            this XElement element, XName trimName)
        {
            if (element.Annotation<DescendantsTrimmedReverseDocumentOrderInfo>() == null)
            {
                XElement ances = element.AncestorsAndSelf(W.txbxContent).FirstOrDefault() ??
                                 element.AncestorsAndSelf().Last();
                InitializeDescendantsTrimmedReverseDocumentOrder(ances, trimName);
            }

            XElement current = element;
            while (true)
            {
                XElement previousElement = current
                    .Annotation<DescendantsTrimmedReverseDocumentOrderInfo>()
                    .PreviousElement;
                if (previousElement == null)
                    yield break;

                yield return previousElement;

                current = previousElement;
            }
        }

        public static string ToStringNewLineOnAttributes(this XElement element)
        {
            var settings = new XmlWriterSettings
            {
                Indent = true,
                OmitXmlDeclaration = true,
                NewLineOnAttributes = true
            };
            var stringBuilder = new StringBuilder();
            using (var stringWriter = new StringWriter(stringBuilder))
            using (XmlWriter xmlWriter = XmlWriter.Create(stringWriter, settings))
                element.WriteTo(xmlWriter);
            return stringBuilder.ToString();
        }

        public static IEnumerable<XElement> DescendantsTrimmed(this XElement element,
            XName trimName)
        {
            return DescendantsTrimmed(element, e => e.Name == trimName);
        }

        public static IEnumerable<XElement> DescendantsTrimmed(this XElement element,
            Func<XElement, bool> predicate)
        {
            Stack<IEnumerator<XElement>> iteratorStack = new Stack<IEnumerator<XElement>>();
            iteratorStack.Push(element.Elements().GetEnumerator());
            while (iteratorStack.Count > 0)
            {
                while (iteratorStack.Peek().MoveNext())
                {
                    XElement currentXElement = iteratorStack.Peek().Current;
                    if (predicate(currentXElement))
                    {
                        yield return currentXElement;
                        continue;
                    }
                    yield return currentXElement;
                    iteratorStack.Push(currentXElement.Elements().GetEnumerator());
                }
                iteratorStack.Pop();
            }
        }

        public static IEnumerable<TResult> Rollup<TSource, TResult>(
            this IEnumerable<TSource> source,
            TResult seed,
            Func<TSource, TResult, TResult> projection)
        {
            TResult nextSeed = seed;
            foreach (TSource src in source)
            {
                TResult projectedValue = projection(src, nextSeed);
                nextSeed = projectedValue;
                yield return projectedValue;
            }
        }

        public static IEnumerable<TResult> Rollup<TSource, TResult>(
            this IEnumerable<TSource> source,
            TResult seed,
            Func<TSource, TResult, int, TResult> projection)
        {
            TResult nextSeed = seed;
            int index = 0;
            foreach (TSource src in source)
            {
                TResult projectedValue = projection(src, nextSeed, index++);
                nextSeed = projectedValue;
                yield return projectedValue;
            }
        }

        public static IEnumerable<TSource> SequenceAt<TSource>(this TSource[] source, int index)
        {
            int i = index;
            while (i < source.Length)
                yield return source[i++];
        }

        public static IEnumerable<T> SkipLast<T>(this IEnumerable<T> source, int count)
        {
            var saveList = new Queue<T>();
            var saved = 0;
            foreach (T item in source)
            {
                if (saved < count)
                {
                    saveList.Enqueue(item);
                    ++saved;
                    continue;
                }

                saveList.Enqueue(item);
                yield return saveList.Dequeue();
            }
        }

        public static bool? ToBoolean(this XAttribute a)
        {
            if (a == null)
                return null;

            string s = ((string) a).ToLower();
            switch (s)
            {
                case "1":
                    return true;
                case "0":
                    return false;
                case "true":
                    return true;
                case "false":
                    return false;
                case "on":
                    return true;
                case "off":
                    return false;
                default:
                    return (bool) a;
            }
        }

        private static string GetQName(XElement xe)
        {
            string prefix = xe.GetPrefixOfNamespace(xe.Name.Namespace);
            if (xe.Name.Namespace == XNamespace.None || prefix == null)
                return xe.Name.LocalName;

            return prefix + ":" + xe.Name.LocalName;
        }

        private static string GetQName(XAttribute xa)
        {
            string prefix = xa.Parent != null ? xa.Parent.GetPrefixOfNamespace(xa.Name.Namespace) : null;
            if (xa.Name.Namespace == XNamespace.None || prefix == null)
                return xa.Name.ToString();

            return prefix + ":" + xa.Name.LocalName;
        }

        private static string NameWithPredicate(XElement el)
        {
            if (el.Parent != null && el.Parent.Elements(el.Name).Count() != 1)
                return GetQName(el) + "[" +
                    (el.ElementsBeforeSelf(el.Name).Count() + 1) + "]";
            else
                return GetQName(el);
        }

        public static string StrCat<T>(this IEnumerable<T> source,
            string separator)
        {
            return source.Aggregate(new StringBuilder(),
                       (sb, i) => sb
                           .Append(i.ToString())
                           .Append(separator),
                       s => s.ToString());
        }

        public static string GetXPath(this XObject xobj)
        {
            if (xobj.Parent == null)
            {
                var doc = xobj as XDocument;
                if (doc != null)
                    return ".";

                var el = xobj as XElement;
                if (el != null)
                    return "/" + NameWithPredicate(el);

                var xt = xobj as XText;
                if (xt != null)
                    return null;

                //
                //the following doesn't work because the XPath data
                //model doesn't include white space text nodes that
                //are children of the document.
                //
                //return
                //    "/" +
                //    (
                //        xt
                //        .Document
                //        .Nodes()
                //        .OfType<XText>()
                //        .Count() != 1 ?
                //        "text()[" +
                //        (xt
                //        .NodesBeforeSelf()
                //        .OfType<XText>()
                //        .Count() + 1) + "]" :
                //        "text()"
                //    );
                //
                var com = xobj as XComment;
                if (com != null && com.Document != null)
                    return
                        "/" +
                        (
                            com
                                .Document
                                .Nodes()
                                .OfType<XComment>()
                                .Count() != 1
                                ? "comment()[" +
                                  (com
                                       .NodesBeforeSelf()
                                       .OfType<XComment>()
                                       .Count() + 1) +
                                  "]"
                                : "comment()"
                        );

                var pi = xobj as XProcessingInstruction;
                if (pi != null)
                    return
                        "/" +
                        (
                            pi.Document != null && pi.Document.Nodes().OfType<XProcessingInstruction>().Count() != 1
                                ? "processing-instruction()[" +
                                  (pi
                                       .NodesBeforeSelf()
                                       .OfType<XProcessingInstruction>()
                                       .Count() + 1) +
                                  "]"
                                : "processing-instruction()"
                        );

                return null;
            }
            else
            {
                var el = xobj as XElement;
                if (el != null)
                {
                    return
                        "/" +
                        el
                            .Ancestors()
                            .InDocumentOrder()
                            .Select(e => NameWithPredicate(e))
                            .StrCat("/") +
                        NameWithPredicate(el);
                }

                var at = xobj as XAttribute;
                if (at != null && at.Parent != null)
                    return
                        "/" +
                        at
                            .Parent
                            .AncestorsAndSelf()
                            .InDocumentOrder()
                            .Select(e => NameWithPredicate(e))
                            .StrCat("/") +
                        "@" + GetQName(at);

                var com = xobj as XComment;
                if (com != null && com.Parent != null)
                    return
                        "/" +
                        com
                            .Parent
                            .AncestorsAndSelf()
                            .InDocumentOrder()
                            .Select(e => NameWithPredicate(e))
                            .StrCat("/") +
                        (
                            com
                                .Parent
                                .Nodes()
                                .OfType<XComment>()
                                .Count() != 1
                                ? "comment()[" +
                                  (com
                                       .NodesBeforeSelf()
                                       .OfType<XComment>()
                                       .Count() + 1) + "]"
                                : "comment()"
                        );

                var cd = xobj as XCData;
                if (cd != null && cd.Parent != null)
                    return
                        "/" +
                        cd
                            .Parent
                            .AncestorsAndSelf()
                            .InDocumentOrder()
                            .Select(e => NameWithPredicate(e))
                            .StrCat("/") +
                        (
                            cd
                                .Parent
                                .Nodes()
                                .OfType<XText>()
                                .Count() != 1
                                ? "text()[" +
                                  (cd
                                       .NodesBeforeSelf()
                                       .OfType<XText>()
                                       .Count() + 1) + "]"
                                : "text()"
                        );

                var tx = xobj as XText;
                if (tx != null && tx.Parent != null)
                    return
                        "/" +
                        tx
                            .Parent
                            .AncestorsAndSelf()
                            .InDocumentOrder()
                            .Select(e => NameWithPredicate(e))
                            .StrCat("/") +
                        (
                            tx
                                .Parent
                                .Nodes()
                                .OfType<XText>()
                                .Count() != 1
                                ? "text()[" +
                                  (tx
                                       .NodesBeforeSelf()
                                       .OfType<XText>()
                                       .Count() + 1) + "]"
                                : "text()"
                        );

                var pi = xobj as XProcessingInstruction;
                if (pi != null && pi.Parent != null)
                    return
                        "/" +
                        pi
                            .Parent
                            .AncestorsAndSelf()
                            .InDocumentOrder()
                            .Select(e => NameWithPredicate(e))
                            .StrCat("/") +
                        (
                            pi
                                .Parent
                                .Nodes()
                                .OfType<XProcessingInstruction>()
                                .Count() != 1
                                ? "processing-instruction()[" +
                                  (pi
                                       .NodesBeforeSelf()
                                       .OfType<XProcessingInstruction>()
                                       .Count() + 1) + "]"
                                : "processing-instruction()"
                        );

                return null;
            }
        }
    }

    public class ExecutableRunner
    {
        public class RunResults
        {
            public int ExitCode;
            public Exception RunException;
            public StringBuilder Output;
            public StringBuilder Error;
        }

        public static RunResults RunExecutable(string executablePath, string arguments, string workingDirectory)
        {
            RunResults runResults = new RunResults
            {
                Output = new StringBuilder(),
                Error = new StringBuilder(),
                RunException = null
            };
            try
            {
                if (File.Exists(executablePath))
                {
                    using (Process proc = new Process())
                    {
                        proc.StartInfo.FileName = executablePath;
                        proc.StartInfo.Arguments = arguments;
                        proc.StartInfo.WorkingDirectory = workingDirectory;
                        proc.StartInfo.UseShellExecute = false;
                        proc.StartInfo.RedirectStandardOutput = true;
                        proc.StartInfo.RedirectStandardError = true;
                        proc.OutputDataReceived +=
                            (o, e) => runResults.Output.Append(e.Data).Append(Environment.NewLine);
                        proc.ErrorDataReceived +=
                            (o, e) => runResults.Error.Append(e.Data).Append(Environment.NewLine);
                        proc.Start();
                        proc.BeginOutputReadLine();
                        proc.BeginErrorReadLine();
                        proc.WaitForExit();
                        runResults.ExitCode = proc.ExitCode;
                    }
                }
                else
                {
                    throw new ArgumentException("Invalid executable path.", "executablePath");
                }
            }
            catch (Exception e)
            {
                runResults.RunException = e;
            }
            return runResults;
        }
    }

    public class SiblingsReverseDocumentOrderInfo
    {
        public XElement PreviousSibling;
    }

    public class DescendantsReverseDocumentOrderInfo
    {
        public XElement PreviousElement;
    }

    public class DescendantsTrimmedReverseDocumentOrderInfo
    {
        public XElement PreviousElement;
    }

    public class GroupOfAdjacent<TSource, TKey> : IGrouping<TKey, TSource>
    {
        public GroupOfAdjacent(List<TSource> source, TKey key)
        {
            GroupList = source;
            Key = key;
        }

        public TKey Key { get; set; }
        private List<TSource> GroupList { get; set; }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<TSource>) this).GetEnumerator();
        }

        IEnumerator<TSource> IEnumerable<TSource>.GetEnumerator()
        {
            return ((IEnumerable<TSource>) GroupList).GetEnumerator();
        }
    }

    public static class PtBucketTimer
    {
        private class BucketInfo
        {
            public int Count;
            public TimeSpan Time;
        }

        public static string LastBucket = null;
        private static DateTime LastTime;
        private static Dictionary<string, BucketInfo> Buckets;

        public static void Bucket(string bucket)
        {
            DateTime now = DateTime.Now;
            if (LastBucket != null)
                AddToBuckets(now);
            LastBucket = bucket;
            LastTime = now;
        }

        public static void End()
        {
            DateTime now = DateTime.Now;
            if (LastBucket != null)
                AddToBuckets(now);
            LastBucket = null;
        }

        private static void AddToBuckets(DateTime now)
        {
            TimeSpan d = now - LastTime;

            if (Buckets.ContainsKey(LastBucket))
            {
                Buckets[LastBucket].Count += 1;
                Buckets[LastBucket].Time += d;
            }
            else
            {
                Buckets.Add(LastBucket, new BucketInfo()
                {
                    Count = 1,
                    Time = d,
                });
            }
            LastTime = now;
        }

        public static string DumpBucketsByKey()
        {
            StringBuilder sb = new StringBuilder();
            foreach (var bucket in Buckets.OrderBy(b => b.Key))
            {
                string ts = bucket.Value.Time.ToString();
                if (ts.Contains('.'))
                    ts = ts.Substring(0, ts.Length - 5);
                string s = bucket.Key.PadRight(80, '-') + "  " + string.Format("{0:00000000}", bucket.Value.Count) + "  " + ts;
                sb.Append(s + Environment.NewLine);
            }
            TimeSpan total = Buckets
                .Aggregate(TimeSpan.Zero, (t, b) => t + b.Value.Time);
            var tz = total.ToString();
            sb.Append(string.Format("Total: {0}", tz.Substring(0, tz.Length - 5)));
            return sb.ToString();
        }

        public static string DumpBucketsToCsvByKey()
        {
            StringBuilder sb = new StringBuilder();
            foreach (var bucket in Buckets.OrderBy(b => b.Key))
            {
                string ts = bucket.Value.Time.TotalMilliseconds.ToString();
                if (ts.Contains('.'))
                    ts = ts.Substring(0, ts.Length - 5);
                string s = bucket.Key + "," + bucket.Value.Count.ToString() + "," + ts;
                sb.Append(s + Environment.NewLine);
            }
            return sb.ToString();
        }

        public static string DumpBucketsByTime()
        {
            StringBuilder sb = new StringBuilder();
            foreach (var bucket in Buckets.OrderBy(b => b.Value.Time))
            {
                string ts = bucket.Value.Time.ToString();
                if (ts.Contains('.'))
                    ts = ts.Substring(0, ts.Length - 5);
                string s = bucket.Key.PadRight(80, '-') + "  " + string.Format("{0:00000000}", bucket.Value.Count) + "  " + ts;
                sb.Append(s + Environment.NewLine);
            }
            TimeSpan total = Buckets
                .Aggregate(TimeSpan.Zero, (t, b) => t + b.Value.Time);
            var tz = total.ToString();
            sb.Append(string.Format("Total: {0}", tz.Substring(0, tz.Length - 5)));
            return sb.ToString();
        }

        public static void Init()
        {
            LastBucket = null;
            Buckets = new Dictionary<string, BucketInfo>();
        }
    }
    
    public class XEntity : XText
    {
        public override void WriteTo(XmlWriter writer)
        {
            if (Value.Substring(0, 1) == "#")
            {
                string e = string.Format("&{0};", Value);
                writer.WriteRaw(e);
            }
            else
                writer.WriteEntityRef(Value);
        }

        public XEntity(string value) : base(value)
        {
        }
    }

    public static class Xsi
    {
        public static XNamespace xsi = "http://www.w3.org/2001/XMLSchema-instance";

        public static XName schemaLocation = xsi + "schemaLocation";
        public static XName noNamespaceSchemaLocation = xsi + "noNamespaceSchemaLocation";
    }
}
