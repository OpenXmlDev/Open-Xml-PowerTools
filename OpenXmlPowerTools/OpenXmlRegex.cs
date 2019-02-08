// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public class OpenXmlRegex
    {
        private const string DontConsolidate = "DontConsolidate";

        private static readonly HashSet<XName> RevTrackMarkupWithId = new HashSet<XName>
        {
            W.cellDel,
            W.cellIns,
            W.cellMerge,
            W.customXmlDelRangeEnd,
            W.customXmlDelRangeStart,
            W.customXmlInsRangeEnd,
            W.customXmlInsRangeStart,
            W.customXmlMoveFromRangeEnd,
            W.customXmlMoveFromRangeStart,
            W.customXmlMoveToRangeEnd,
            W.customXmlMoveToRangeStart,
            W.del,
            W.ins,
            W.moveFrom,
            W.moveFromRangeEnd,
            W.moveFromRangeStart,
            W.moveTo,
            W.moveToRangeEnd,
            W.moveToRangeStart,
            W.pPrChange,
            W.rPrChange,
            W.sectPrChange,
            W.tblGridChange,
            W.tblPrChange,
            W.tblPrExChange,
            W.tcPrChange
        };

        public static int Match(IEnumerable<XElement> content, Regex regex)
        {
            return ReplaceInternal(content, regex, null, null, false, null, true);
        }

        /// <summary>
        /// If callback == null Then returns count of matches in the content
        /// If callback != null Then Match calls Found for each match
        /// </summary>
        public static int Match(IEnumerable<XElement> content, Regex regex, Action<XElement, Match> found)
        {
            return ReplaceInternal(content, regex, null,
                (x, m) =>
                {
                    if (found != null) found.Invoke(x, m);
                    return true;
                },
                false, null, true);
        }

        /// <summary>
        /// If replacement == "new content" && callback == null
        ///     Then replaces all matches
        /// If replacement == "" && callback == null)
        ///     Then deletes all matches
        /// If replacement == "new content" && callback != null)
        ///     Then the callback can return true / false to indicate whether to replace or not
        /// If the callback returns true once, and false on all subsequent calls, then this method replaces only the first found.
        /// If replacement == "" && callback != null)
        ///     Then the callback can return true / false to indicate whether to delete or not
        /// </summary>
        public static int Replace(IEnumerable<XElement> content, Regex regex, string replacement,
            Func<XElement, Match, bool> doReplacement)
        {
            return ReplaceInternal(content, regex, replacement, doReplacement, false, null, true);
        }

        /// <summary>
        /// This overload enables not coalescing content, which is necessary for DocumentAssembler.
        /// </summary>
        public static int Replace(IEnumerable<XElement> content, Regex regex, string replacement,
            Func<XElement, Match, bool> doReplacement, bool coalesceContent)
        {
            return ReplaceInternal(content, regex, replacement, doReplacement, false, null, coalesceContent);
        }

        /// <summary>
        /// If replacement == "new content" && callback == null
        ///     Then replaces all matches
        /// If replacement == "" && callback == null)
        ///     Then deletes all matches
        /// If replacement == "new content" && callback != null)
        ///     Then the callback can return true / false to indicate whether to replace or not
        /// If the callback returns true once, and false on all subsequent calls, then this method replaces only the first found.
        /// If replacement == "" && callback != null)
        ///     Then the callback can return true / false to indicate whether to delete or not
        /// If trackRevisions == true
        ///     Then replacement is done using revision tracking markup, with author as the revision tracking author
        /// If trackRevisions == true for a PPTX
        ///     Then code throws an exception
        /// </summary>
        public static int Replace(IEnumerable<XElement> content, Regex regex, string replacement,
            Func<XElement, Match, bool> doReplacement, bool trackRevisions, string author)
        {
            return ReplaceInternal(content, regex, replacement, doReplacement, trackRevisions, author, true);
        }

        private static int ReplaceInternal(IEnumerable<XElement> content, Regex regex, string replacement,
            Func<XElement, Match, bool> callback, bool trackRevisions, string revisionTrackingAuthor,
            bool coalesceContent)
        {
            if (content == null) throw new ArgumentNullException("content");
            if (regex == null) throw new ArgumentNullException("regex");

            IEnumerable<XElement> contentList = content as IList<XElement> ?? content.ToList();

            XElement first = contentList.FirstOrDefault();
            if (first == null)
                return 0;

            if (first.Name.Namespace == W.w)
            {
                if (!contentList.Any())
                    return 0;

                var replInfo = new ReplaceInternalInfo { Count = 0 };
                foreach (XElement c in contentList)
                {
                    var newC = (XElement) WmlSearchAndReplaceTransform(c, regex, replacement, callback, trackRevisions,
                        revisionTrackingAuthor, replInfo, coalesceContent);
                    c.ReplaceNodes(newC.Nodes());
                }

                XElement root = contentList.First().AncestorsAndSelf().Last();
                int nextId = new[] { 0 }
                                 .Concat(root
                                     .Descendants()
                                     .Where(d => RevTrackMarkupWithId.Contains(d.Name))
                                     .Attributes(W.id)
                                     .Select(a => (int) a))
                                 .Max() + 1;
                IEnumerable<XElement> revTrackingWithoutId = root
                    .DescendantsAndSelf()
                    .Where(d => RevTrackMarkupWithId.Contains(d.Name) && (d.Attribute(W.id) == null));
                foreach (XElement item in revTrackingWithoutId)
                    item.Add(new XAttribute(W.id, nextId++));

                List<IGrouping<int, XElement>> revTrackingWithDuplicateIds = root
                    .DescendantsAndSelf()
                    .Where(d => RevTrackMarkupWithId.Contains(d.Name))
                    .GroupBy(d => (int) d.Attribute(W.id))
                    .Where(g => g.Count() > 1)
                    .ToList();
                foreach (IGrouping<int, XElement> group in revTrackingWithDuplicateIds)
                    foreach (XElement gc in group.Skip(1))
                    {
                        XAttribute xAttribute = gc.Attribute(W.id);
                        if (xAttribute != null) xAttribute.Value = nextId.ToString();
                        nextId++;
                    }

                return replInfo.Count;
            }

            if ((first.Name.Namespace == P.p) || (first.Name.Namespace == A.a))
            {
                if (trackRevisions)
                    throw new OpenXmlPowerToolsException("PPTX does not support revision tracking");

                var counter = new ReplaceInternalInfo { Count = 0 };
                foreach (XElement c in contentList)
                {
                    var newC = (XElement) PmlSearchAndReplaceTransform(c, regex, replacement, callback, counter);
                    c.ReplaceNodes(newC.Nodes());
                }

                return counter.Count;
            }

            return 0;
        }

        private static object WmlSearchAndReplaceTransform(XNode node, Regex regex, string replacement,
            Func<XElement, Match, bool> callback, bool trackRevisions, string revisionTrackingAuthor,
            ReplaceInternalInfo replInfo, bool coalesceContent)
        {
            var element = node as XElement;
            if (element == null) return node;

            if (element.Name == W.p)
            {
                XElement paragraph = element;

                string preliminaryContent = paragraph
                    .DescendantsTrimmed(W.txbxContent)
                    .Where(d => d.Name == W.r && (d.Parent == null || d.Parent.Name != W.del))
                    .Select(UnicodeMapper.RunToString)
                    .StringConcatenate();
                if (regex.IsMatch(preliminaryContent))
                {
                    var paragraphWithSplitRuns = new XElement(W.p,
                        paragraph.Attributes(),
                        paragraph.Nodes().Select(n => WmlSearchAndReplaceTransform(n, regex, replacement, callback,
                            trackRevisions, revisionTrackingAuthor, replInfo, coalesceContent)));

                    IEnumerable<XElement> runsTrimmed = paragraphWithSplitRuns
                        .DescendantsTrimmed(W.txbxContent)
                        .Where(d => d.Name == W.r && (d.Parent == null || d.Parent.Name != W.del));

                    var charsAndRuns = runsTrimmed
                        .Select(r => new { Ch = UnicodeMapper.RunToString(r), r })
                        .ToList();

                    string content = charsAndRuns.Select(t => t.Ch).StringConcatenate();
                    XElement[] alignedRuns = charsAndRuns.Select(t => t.r).ToArray();

                    MatchCollection matchCollection = regex.Matches(content);
                    replInfo.Count += matchCollection.Count;

                    // Process Match
                    if (replacement == null)
                    {
                        if (callback == null) return paragraph;

                        foreach (Match match in matchCollection.Cast<Match>())
                            callback(paragraph, match);

                        return paragraph;
                    }

                    // Process Replace
                    foreach (Match match in matchCollection.Cast<Match>())
                    {
                        if (match.Length == 0) continue;
                        if ((callback != null) && !callback(paragraph, match)) continue;

                        List<XElement> runCollection = alignedRuns
                            .Skip(match.Index)
                            .Take(match.Length)
                            .ToList();

                        // uses the Skip / Take special semantics of array to implement efficient finding of sub array

                        XElement firstRun = runCollection.First();
                        XElement firstRunProperties = firstRun.Elements(W.rPr).FirstOrDefault();

                        // save away first run properties

                        if (trackRevisions)
                        {
                            if (replacement != "")
                            {
                                // We coalesce runs as some methods, e.g., in DocumentAssembler,
                                // will try to find the replacement string even though they
                                // set coalesceContent to false.
                                string newTextValue = match.Result(replacement);
                                List<XElement> newRuns = UnicodeMapper.StringToCoalescedRunList(newTextValue,
                                    firstRunProperties);
                                var newIns = new XElement(W.ins,
                                    new XAttribute(W.author, revisionTrackingAuthor),
                                    new XAttribute(W.date, DateTime.UtcNow.ToString("s") + "Z"),
                                    newRuns);

                                if (firstRun.Parent != null && firstRun.Parent.Name == W.ins)
                                    firstRun.Parent.AddBeforeSelf(newIns);
                                else
                                    firstRun.AddBeforeSelf(newIns);
                            }

                            foreach (XElement run in runCollection)
                            {
                                bool isInIns = run.Parent != null && run.Parent.Name == W.ins;
                                if (isInIns)
                                {
                                    XElement parentIns = run.Parent;
                                    XElement grandParentParagraph = parentIns.Parent;
                                    if (grandParentParagraph != null)
                                    {
                                        if ((string) parentIns.Attributes(W.author).FirstOrDefault() ==
                                            revisionTrackingAuthor)
                                        {
                                            List<XElement> parentInsSiblings = grandParentParagraph
                                                .Elements()
                                                .Where(c => c != parentIns)
                                                .ToList();
                                            grandParentParagraph.ReplaceNodes(parentInsSiblings);
                                        }
                                        else
                                        {
                                            List<XElement> parentInsSiblings = grandParentParagraph
                                                .Elements()
                                                .Select(c => c == parentIns
                                                    ? new XElement(W.ins,
                                                        parentIns.Attributes(),
                                                        new XElement(W.del,
                                                            new XAttribute(W.author, revisionTrackingAuthor),
                                                            new XAttribute(W.date, DateTime.UtcNow.ToString("s") + "Z"),
                                                            parentIns.Elements().Select(TransformToDelText)))
                                                    : c)
                                                .ToList();
                                            grandParentParagraph.ReplaceNodes(parentInsSiblings);
                                        }
                                    }
                                }
                                else
                                {
                                    var delRun = new XElement(W.del,
                                        new XAttribute(W.author, revisionTrackingAuthor),
                                        new XAttribute(W.date, DateTime.UtcNow.ToString("s") + "Z"),
                                        TransformToDelText(run));
                                    run.ReplaceWith(delRun);
                                }
                            }
                        }
                        else // not tracked revisions
                        {
                            foreach (XElement runToDelete in runCollection.Skip(1).ToList())
                                if (runToDelete.Parent != null && runToDelete.Parent.Name == W.ins)
                                    runToDelete.Parent.Remove();
                                else
                                    runToDelete.Remove();

                            // We coalesce runs as some methods, e.g., in DocumentAssembler,
                            // will try to find the replacement string even though they
                            // set coalesceContent to false.
                            string newTextValue = match.Result(replacement);
                            List<XElement> newRuns = UnicodeMapper.StringToCoalescedRunList(newTextValue,
                                firstRunProperties);
                            if (firstRun.Parent != null && firstRun.Parent.Name == W.ins)
                                firstRun.Parent.ReplaceWith(newRuns);
                            else
                                firstRun.ReplaceWith(newRuns);
                        }
                    }

                    return coalesceContent
                        ? WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(paragraphWithSplitRuns)
                        : paragraphWithSplitRuns;
                }

                var newParagraph = new XElement(W.p,
                    paragraph.Attributes(),
                    paragraph.Nodes().Select(n =>
                    {
                        var e = n as XElement;
                        if (e == null) return n;

                        if (e.Name == W.pPr)
                            return e;
                        if (((e.Name == W.r) && e.Elements(W.t).Any()) || e.Elements(W.tab).Any())
                            return e;
                        if ((e.Name == W.ins) && e.Elements(W.r).Elements(W.t).Any())
                            return e;

                        return WmlSearchAndReplaceTransform(e, regex, replacement, callback,
                            trackRevisions, revisionTrackingAuthor, replInfo, coalesceContent);
                    }));
                return coalesceContent
                    ? WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(newParagraph) // CoalesceContent(newParagraph)
                    : newParagraph;
            }

            if (element.Name == W.ins && element.Elements(W.r).Any())
            {
                List<object> collectionOfCollections = element
                    .Elements()
                    .Select(n => WmlSearchAndReplaceTransform(n, regex, replacement, callback, trackRevisions,
                        revisionTrackingAuthor, replInfo, coalesceContent))
                    .ToList();
                List<object> collectionOfIns = collectionOfCollections
                    .Select(c =>
                    {
                        var elements = c as IEnumerable<XElement>;
                        return elements != null
                            ? elements.Select(ixc => new XElement(W.ins, element.Attributes(), ixc))
                            : c;
                    })
                    .ToList();
                return collectionOfIns;
            }

            if (element.Name == W.r)
            {
                return element.Elements()
                    .Where(e => e.Name != W.rPr)
                    .Select(e => e.Name == W.t
                        ? ((string) e).Select(c =>
                            new XElement(W.r,
                                element.Elements(W.rPr),
                                new XElement(W.t, XmlUtil.GetXmlSpaceAttribute(c), c)))
                        : new[] { new XElement(W.r, element.Elements(W.rPr), e) })
                    .SelectMany(t => t);
            }

            return new XElement(element.Name,
                element.Attributes(),
                element.Nodes()
                    .Select(n => WmlSearchAndReplaceTransform(n, regex, replacement, callback, trackRevisions,
                        revisionTrackingAuthor, replInfo, coalesceContent)));
        }

        private static object TransformToDelText(XNode node)
        {
            var element = node as XElement;
            if (element == null) return node;

            if (element.Name == W.t)
                return new XElement(W.delText,
                    XmlUtil.GetXmlSpaceAttribute(element.Value),
                    element.Value);

            return new XElement(element.Name,
                element.Attributes(),
                element.Nodes().Select(TransformToDelText));
        }

        private static object PmlSearchAndReplaceTransform(XNode node, Regex regex, string replacement,
            Func<XElement, Match, bool> callback, ReplaceInternalInfo counter)
        {
            var element = node as XElement;
            if (element == null) return node;

            if (element.Name == A.p)
            {
                XElement paragraph = element;
                string contents = element.Descendants(A.t).Select(t => (string) t).StringConcatenate();
                if (!regex.IsMatch(contents))
                    return new XElement(element.Name, element.Attributes(), element.Nodes());

                var paragraphWithSplitRuns = new XElement(A.p,
                    paragraph.Attributes(),
                    paragraph.Nodes()
                        .Select(n => PmlSearchAndReplaceTransform(n, regex, replacement, callback, counter)));

                List<XElement> runsTrimmed = paragraphWithSplitRuns
                    .Descendants(A.r)
                    .ToList();

                var charsAndRuns = runsTrimmed
                    .Select(r =>
                        r.Element(A.t) != null
                            ? new { Ch = r.Element(A.t).Value, r }
                            : new { Ch = "\x01", r })
                    .ToList();

                string content = charsAndRuns.Select(t => t.Ch).StringConcatenate();
                XElement[] alignedRuns = charsAndRuns.Select(t => t.r).ToArray();

                MatchCollection matchCollection = regex.Matches(content);
                counter.Count += matchCollection.Count;
                if (replacement == null)
                {
                    foreach (Match match in matchCollection.Cast<Match>())
                        callback(paragraph, match);
                }
                else
                {
                    foreach (Match match in matchCollection.Cast<Match>())
                    {
                        if ((callback != null) && !callback(paragraph, match)) continue;

                        List<XElement> runCollection = alignedRuns
                            .Skip(match.Index)
                            .Take(match.Length)
                            .ToList();

                        // uses the Skip / Take special semantics of array to implement efficient finding of sub array

                        XElement firstRun = runCollection.First();

                        // save away first run because we want the run properties

                        runCollection.Skip(1).Remove();

                        // binds to Remove(this IEnumerable<XElement> elements), which is an extension

                        // in LINQ to XML that uses snapshot semantics and removes every element from
                        // its parent.

                        var newFirstRun = new XElement(A.r,
                            firstRun.Element(A.rPr),
                            new XElement(A.t, replacement));

                        // creates a new run with proper run properties

                        firstRun.ReplaceWith(newFirstRun);

                        // finds firstRun in its parent's list of children, unparents firstRun,

                        // sets newFirstRun's parent to firstRuns old parent, and inserts in the list
                        // of children at the right place.
                    }
                    XElement paragraphWithReplacedRuns = paragraphWithSplitRuns;

                    IEnumerable<IGrouping<string, XElement>> groupedAdjacentRunsWithIdenticalFormatting =
                        paragraphWithReplacedRuns
                            .Elements()
                            .GroupAdjacent(ce =>
                            {
                                if (ce.Name != A.r)
                                    return DontConsolidate;
                                if ((ce.Elements().Count(e => e.Name != A.rPr) != 1) || (ce.Element(A.t) == null))
                                    return DontConsolidate;

                                XElement rPr = ce.Element(A.rPr);
                                return rPr == null ? "" : rPr.ToString(SaveOptions.None);
                            });
                    var paragraphWithConsolidatedRuns = new XElement(A.p,
                        groupedAdjacentRunsWithIdenticalFormatting.Select(g =>
                        {
                            if (g.Key == DontConsolidate)
                                return (object) g;

                            string textValue = g.Select(r => r.Element(A.t).Value).StringConcatenate();
                            XAttribute xs = XmlUtil.GetXmlSpaceAttribute(textValue);
                            return new XElement(A.r,
                                g.First().Elements(A.rPr),
                                new XElement(A.t, xs, textValue));
                        }));
                    paragraph = paragraphWithConsolidatedRuns;
                }

                return paragraph;
            }

            if ((element.Name == A.r) && element.Elements(A.t).Any())
            {
                return element.Elements()
                    .Where(e => e.Name != A.rPr)
                    .Select(e =>
                    {
                        if (e.Name == A.t)
                        {
                            var s = (string) e;
                            IEnumerable<XElement> collectionOfSubRuns = s.Select(c => new XElement(A.r,
                                element.Elements(A.rPr),
                                new XElement(A.t, XmlUtil.GetXmlSpaceAttribute(c), c)));
                            return (object) collectionOfSubRuns;
                        }

                        return new XElement(A.r,
                            element.Elements(A.rPr),
                            e);
                    });
            }

            return new XElement(element.Name,
                element.Attributes(),
                element.Nodes().Select(n => PmlSearchAndReplaceTransform(n, regex, replacement, callback, counter)));
        }

        private class ReplaceInternalInfo
        {
            public int Count;
        }
    }
}
