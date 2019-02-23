using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public partial class DocumentAssembler : DocumentAssemblerBase
    {
        public static Task<AssembleResult> AssembleDocumentAsync(WmlDocument templateDoc, XmlDocument data)
        {
            XDocument xDoc = data.GetXDocument();
            return AssembleDocumentAsync(templateDoc, xDoc.Root);
        }

        public static async Task<AssembleResult> AssembleDocumentAsync(WmlDocument templateDoc, XElement data)
        {
            var dataSource = new AsyncXmlDataContext(data);
            byte[] byteArray = templateDoc.DocumentByteArray;
            bool templateError = false;
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, (int)byteArray.Length);
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, true))
                {
                    templateError = await AssembleDocumentAsync(wordDoc, dataSource);
                }
                return new AssembleResult(new WmlDocument("TempFileName.docx", mem.ToArray()), templateError);
            }
        }
    }

    public partial class DocumentAssemblerBase
    {
        protected static async Task<bool> AssembleDocumentAsync(WordprocessingDocument wordDoc, IAsyncDataContext data)
        {
            if (RevisionAccepter.HasTrackedRevisions(wordDoc))
                throw new OpenXmlPowerToolsException("Invalid DocumentAssembler template - contains tracked revisions");

            var te = new TemplateError();
            var partTasks = wordDoc.ContentParts().Select(part => ProcessTemplatePartAsync(data, te, part));
            await Task.WhenAll(partTasks);
            return te.HasError;
        }

        protected static async Task<AssembleResult> AssembleDocumentAsync(WmlDocument templateDoc, string outputFilename, IAsyncDataContext dataSource)
        {
            WmlDocument assembledDocument = null;
            bool templateError = false;
            byte[] byteArray = templateDoc.DocumentByteArray;
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, (int)byteArray.Length); // copy template file (binary) into memory -- I guess so the template/file handle isn't held/locked?
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, true)) // read & parse that byte array into OXML document (also in memory)
                {
                    templateError = await AssembleDocumentAsync(wordDoc, dataSource);
                }
                assembledDocument = new WmlDocument(outputFilename, mem.ToArray());
            }
            return new AssembleResult(assembledDocument, templateError);
        }

        protected static async Task<bool> PrepareTemplateAsync(WordprocessingDocument wordDoc, IMetadataParser fieldParser)
        {
            if (RevisionAccepter.HasTrackedRevisions(wordDoc))
                throw new OpenXmlPowerToolsException("Invalid DocumentAssembler template - contains tracked revisions");

            await Task.Yield(); // force asynchrony (for testing, mostly)

            var te = new TemplateError();
            foreach (var part in wordDoc.ContentParts())
            {
                XDocument xDoc = part.GetXDocument();
                xDoc.Elements().First().ReplaceWith(PrepareTemplatePart(fieldParser, te, xDoc.Root));
                part.PutXDocument(); // if we were processing template parts in parallel (as we do during assembly), this would need to be lock{}ed
            }
            return te.HasError;
        }
        private static readonly object s_partLock = new object();

        private static async Task ProcessTemplatePartAsync(IAsyncDataContext data, TemplateError te, OpenXmlPart part)
        {
            XDocument xDoc = part.GetXDocument();
            var xDocRoot = PrepareTemplatePart(data, te, xDoc.Root);

            // do the actual content replacement
            xDocRoot = (XElement)await ContentReplacementTransformAsync(xDocRoot, data, te);

            xDoc.Elements().First().ReplaceWith(xDocRoot);
            // work around apparent issues with thread safety when replacing the content of a part within a package
            lock (s_partLock)
            {
                part.PutXDocument();
            }
        }

        static async Task<object> ContentReplacementTransformAsync(XNode node, IAsyncDataContext data, TemplateError templateError)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == PA.Content)
                {
                    XElement para = element.Descendants(W.p).FirstOrDefault();
                    XElement run = element.Descendants(W.r).FirstOrDefault();

                    var selector = (string)element.Attribute(PA.Select);
                    var optionalString = (string)element.Attribute(PA.Optional);
                    bool optional = (optionalString != null && optionalString.ToLower() == "true");

                    string newValue;
                    try
                    {
                        newValue = await data.EvaluateTextAsync(selector, optional);
                    }
                    catch (EvaluationException e)
                    {
                        return CreateContextErrorMessage(element, "EvaluationException: " + e.Message, templateError);
                    }

                    return ReplaceValue(para, run, newValue);
                }
                if (element.Name == PA.Repeat)
                {
                    string selector = (string)element.Attribute(PA.Select);
                    var optionalString = (string)element.Attribute(PA.Optional);
                    bool optional = (optionalString != null && optionalString.ToLower() == "true");

                    IAsyncDataContext[] repeatingData;
                    try
                    {
                        repeatingData = await data.EvaluateListAsync(selector, optional);
                    }
                    catch (EvaluationException e)
                    {
                        return CreateContextErrorMessage(element, "EvaluationException: " + e.Message, templateError);
                    }
                    if (!repeatingData.Any())
                    {
                        return null;
                        //XElement para = element.Descendants(W.p).FirstOrDefault();
                        //if (para != null)
                        //    return new XElement(W.p, new XElement(W.r));
                        //else
                        //    return new XElement(W.r);
                    }
                    var newContentTasks = repeatingData.Select(async d =>
                        {
                            var contentTasks = element
                                .Elements()
                                .Select(e => ContentReplacementTransformAsync(e, d, templateError));
                            var content = await Task.WhenAll(contentTasks);
                            await d.ReleaseAsync();
                            return content;
                        });
                    var newContent = await Task.WhenAll(newContentTasks);
                    return newContent;
                }
                if (element.Name == PA.Table)
                {
                    IAsyncDataContext[] tableData;
                    try
                    {
                        tableData = await data.EvaluateListAsync((string)element.Attribute(PA.Select), true);
                    }
                    catch (EvaluationException e)
                    {
                        return CreateContextErrorMessage(element, "EvaluationException: " + e.Message, templateError);
                    }
                    if (tableData.Count() == 0)
                        return CreateContextErrorMessage(element, "Table Select returned no data", templateError);
                    XElement table = element.Element(W.tbl);
                    XElement protoRow = table.Elements(W.tr).Skip(1).FirstOrDefault();
                    var footerRowsBeforeTransform = table
                        .Elements(W.tr)
                        .Skip(2)
                        .ToList();
                    var footerRowTasks = footerRowsBeforeTransform
                        .Select(x => ContentReplacementTransformAsync(x, data, templateError));
                    var footerRows = await Task.WhenAll(footerRowTasks);
                    if (protoRow == null)
                        return CreateContextErrorMessage(element, string.Format("Table does not contain a prototype row"), templateError);
                    protoRow.Descendants(W.bookmarkStart).Remove();
                    protoRow.Descendants(W.bookmarkEnd).Remove();
                    var dataRowTasks = tableData.Select(async d =>
                        {
                            var cellTasks = protoRow.Elements(W.tc)
                                .Select(async tc =>
                                {
                                    XElement paragraph = tc.Elements(W.p).FirstOrDefault();
                                    XElement cellRun = paragraph.Elements(W.r).FirstOrDefault();
                                    string selector = paragraph.Value;
                                    string newValue = null;
                                    try
                                    {
                                        newValue = await d.EvaluateTextAsync(selector, false);
                                    }
                                    catch (EvaluationException e)
                                    {
                                        XElement errorCell = new XElement(W.tc,
                                            tc.Elements().Where(z => z.Name != W.p),
                                            new XElement(W.p,
                                                paragraph.Element(W.pPr),
                                                CreateRunErrorMessage(e.Message, templateError)));
                                        return errorCell;
                                    }

                                    XElement newCell = new XElement(W.tc,
                                        tc.Elements().Where(z => z.Name != W.p),
                                        new XElement(W.p,
                                            paragraph.Element(W.pPr),
                                            new XElement(W.r,
                                                cellRun != null ? cellRun.Element(W.rPr) : new XElement(W.rPr),  //if the cell was empty there is no cellrun
                                                new XElement(W.t, newValue))));
                                    return newCell;
                                });
                            var rowContent = new XElement(W.tr,
                                protoRow.Elements().Where(r => r.Name != W.tc),
                                await Task.WhenAll(cellTasks));
                            await d.ReleaseAsync();
                            return rowContent;
                        });
                    XElement newTable = new XElement(W.tbl,
                        table.Elements().Where(e => e.Name != W.tr),
                        table.Elements(W.tr).FirstOrDefault(),
                        await Task.WhenAll(dataRowTasks),
                        footerRows
                        );
                    return newTable;
                }
                if (element.Name == PA.Conditional)
                {
                    string selector = (string)element.Attribute(PA.Select);
                    var match = (string)element.Attribute(PA.Match);
                    var notMatch = (string)element.Attribute(PA.NotMatch);
                    bool testValue;

                    try
                    {
                        testValue = await data.EvaluateBoolAsync(selector, match, notMatch);
                    }
                    catch (EvaluationException e)
                    {
                        return CreateContextErrorMessage(element, e.Message, templateError);
                    }

                    if (testValue)
                    {
                        var contentTasks = element.Elements().Select(e => ContentReplacementTransformAsync(e, data, templateError));
                        var content = await Task.WhenAll(contentTasks);
                        return content;
                    }
                    return null;
                }
                var childNodeTasks = element.Nodes().Select(n => ContentReplacementTransformAsync(n, data, templateError));
                return new XElement(element.Name,
                    element.Attributes(),
                    await Task.WhenAll(childNodeTasks));
            }
            return node;
        }
    }

    public class AsyncXmlDataContext : XmlMetadataParser, IAsyncDataContext
    {
        private XElement _element;

        public AsyncXmlDataContext(XElement data)
        {
            _element = data;
        }

        public async Task<IAsyncDataContext[]> EvaluateListAsync(string selector, bool optional)
        {
            await Task.Yield(); // make this async -- silly, but really just for testing
            IEnumerable<XElement> repeatingData;
            try
            {
                repeatingData = _element.XPathSelectElements(selector);
            }
            catch (XPathException e)
            {
                throw new EvaluationException("XPathException: " + e.Message);
            }
            var newContent = repeatingData.Select(d => new AsyncXmlDataContext(d)).ToArray();
            if (!newContent.Any())
            {
                if (!optional)
                    throw new EvaluationException("Repeat: Select returned no data");
            }
            return newContent;
        }

        public async Task<string> EvaluateTextAsync(string xPath, bool optional)
        {
            await Task.Yield(); // make this async -- silly, but really just for testing
            object xPathSelectResult;
            try
            {
                //support some cells in the table may not have an xpath expression.
                if (String.IsNullOrWhiteSpace(xPath)) return String.Empty;

                xPathSelectResult = _element.XPathEvaluate(xPath);
            }
            catch (XPathException e)
            {
                throw new EvaluationException("XPathException: " + e.Message, e);
            }

            if ((xPathSelectResult is IEnumerable) && !(xPathSelectResult is string))
            {
                var selectedData = ((IEnumerable)xPathSelectResult).Cast<XObject>();
                if (!selectedData.Any())
                {
                    if (optional) return string.Empty;
                    throw new EvaluationException(string.Format("XPath expression ({0}) returned no results", xPath));
                }
                if (selectedData.Count() > 1)
                {
                    throw new EvaluationException(string.Format("XPath expression ({0}) returned more than one node", xPath));
                }

                XObject selectedDatum = selectedData.First();

                if (selectedDatum is XElement) return ((XElement)selectedDatum).Value;

                if (selectedDatum is XAttribute) return ((XAttribute)selectedDatum).Value;
            }

            return xPathSelectResult.ToString();

        }

        public async Task<bool> EvaluateBoolAsync(string xPath, string match, string notMatch)
        {
            if (match == null && notMatch == null)
                throw new EvaluationException("Conditional: Must specify either Match or NotMatch");
            if (match != null && notMatch != null)
                throw new EvaluationException("Conditional: Cannot specify both Match and NotMatch");

            string testValue = await EvaluateTextAsync(xPath, false);

            return (match != null && testValue == match) || (notMatch != null && testValue != notMatch);
        }

        public async Task ReleaseAsync()
        {
            await Task.Yield(); // make this async -- silly, but really just for testing
            _element = null;
        }
    }

    public class AssembleResult
    {
        public WmlDocument Document { get; private set; }
        public bool HasErrors { get; private set; }

        internal AssembleResult(WmlDocument document, bool hasErrors)
        {
            Document = document;
            HasErrors = hasErrors;
        }

        public void Deconstruct(out WmlDocument document, out bool hasErrors)
        {
            document = Document;
            hasErrors = HasErrors;
        }
    }
}
