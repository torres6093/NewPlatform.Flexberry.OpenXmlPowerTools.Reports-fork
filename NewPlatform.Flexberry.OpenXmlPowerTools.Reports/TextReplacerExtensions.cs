namespace NewPlatform.Flexberry.Reports
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Xml.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using OpenXmlPowerTools;

    public class TextReplacerExtensions
    {
        private class MatchSemaphore
        {
            public int MatchId;

            public MatchSemaphore(int matchId)
            {
                MatchId = matchId;
            }
        }

        private static XObject CloneWithAnnotation(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                XElement newElement = new XElement(
                    element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => CloneWithAnnotation(n)));
                if (element.Annotation<MatchSemaphore>() != null) newElement.AddAnnotation(element.Annotation<MatchSemaphore>());
            }
            return node;
        }

        public static WmlDocument SearchAndReplace(WmlDocument doc, string search, string replace, bool matchCase)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            {
                using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
                {
                    SearchAndReplace(document, search, replace, matchCase);
                }

                return streamDoc.GetModifiedWmlDocument();
            }
        }

        public static void SearchAndReplace(
            WordprocessingDocument wordDoc,
            string search,
            string replace,
            bool matchCase,
            XElement rPrFormat = null)
        {
            if (RevisionAccepter.HasTrackedRevisions(wordDoc))
                throw new InvalidDataException(
                    "Search and replace will not work with documents " + "that contain revision tracking.");
            XDocument xDoc;
            xDoc = wordDoc.MainDocumentPart.DocumentSettingsPart.GetXDocument();
            if (xDoc.Descendants(W.trackRevisions).Any()) throw new InvalidDataException("Revision tracking is turned on for document.");

            xDoc = wordDoc.MainDocumentPart.GetXDocument();
            WmlSearchAndReplaceInXDocument(xDoc, search, replace, matchCase, rPrFormat);
            wordDoc.MainDocumentPart.PutXDocument();
            foreach (var part in wordDoc.MainDocumentPart.HeaderParts)
            {
                xDoc = part.GetXDocument();
                WmlSearchAndReplaceInXDocument(xDoc, search, replace, matchCase, rPrFormat);
                part.PutXDocument();
            }

            foreach (var part in wordDoc.MainDocumentPart.FooterParts)
            {
                xDoc = part.GetXDocument();
                WmlSearchAndReplaceInXDocument(xDoc, search, replace, matchCase, rPrFormat);
                part.PutXDocument();
            }

            if (wordDoc.MainDocumentPart.EndnotesPart != null)
            {
                xDoc = wordDoc.MainDocumentPart.EndnotesPart.GetXDocument();
                WmlSearchAndReplaceInXDocument(xDoc, search, replace, matchCase, rPrFormat);
                wordDoc.MainDocumentPart.EndnotesPart.PutXDocument();
            }

            if (wordDoc.MainDocumentPart.FootnotesPart != null)
            {
                xDoc = wordDoc.MainDocumentPart.FootnotesPart.GetXDocument();
                WmlSearchAndReplaceInXDocument(xDoc, search, replace, matchCase, rPrFormat);
                wordDoc.MainDocumentPart.FootnotesPart.PutXDocument();
            }
        }

        private static void WmlSearchAndReplaceInXDocument(
            XDocument xDocument,
            string search,
            string replace,
            bool matchCase,
            XElement rPrFormat = null)
        {
            XElement newRoot = (XElement)WmlSearchAndReplaceTransform(xDocument.Root, search, replace, matchCase, rPrFormat);
            xDocument.Elements().First().ReplaceWith(newRoot);
        }

        public static object WmlSearchAndReplaceTransform(XNode node, string search, string replace, bool matchCase, XElement rPrFormat = null)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.p)
                {
                    string contents = element.Descendants(W.t).Select(t => (string)t).StringConcatenate();
                    if (contents.Contains(search) || (!matchCase && contents.ToUpper().Contains(search.ToUpper())))
                    {
                        XElement paragraphWithSplitRuns = new XElement(
                            W.p,
                            element.Attributes(),
                            element.Nodes().Select(n => WmlSearchAndReplaceTransform(n, search, replace, matchCase, rPrFormat)));
                        XElement[] subRunArray = paragraphWithSplitRuns.Elements(W.r).Where(
                            e =>
                            {
                                XElement subRunElement = e.Elements().FirstOrDefault(el => el.Name != W.rPr);
                                if (subRunElement == null)
                                {
                                    return false;
                                }

                                return W.SubRunLevelContent.Contains(subRunElement.Name);
                            }).ToArray();
                        int paragraphChildrenCount = subRunArray.Length;
                        int matchId = 1;
                        foreach (
                            var pc in
                                subRunArray.Take(paragraphChildrenCount - (search.Length - 1))
                                    .Select((c, i) => new { Child = c, Index = i, }))
                        {
                            var subSequence = subRunArray.SequenceAt(pc.Index).Take(search.Length);
                            var zipped = subSequence.Zip(
                                search,
                                (pcp, c) => new { ParagraphChildProjection = pcp, CharacterToCompare = c, });
                            bool dontMatch = zipped.Any(
                                z =>
                                {
                                    if (z.ParagraphChildProjection.Annotation<MatchSemaphore>() != null) return true;
                                    bool b;
                                    if (matchCase) b = z.ParagraphChildProjection.Value != z.CharacterToCompare.ToString();
                                    else
                                        b = z.ParagraphChildProjection.Value.ToUpper()
                                            != z.CharacterToCompare.ToString().ToUpper();
                                    return b;
                                });
                            bool match = !dontMatch;
                            if (match)
                            {
                                foreach (var item in subSequence) item.AddAnnotation(new MatchSemaphore(matchId));
                                ++matchId;
                            }
                        }

                        // The following code is locally impure, as this is the most expressive way to write it.
                        XElement paragraphWithReplacedRuns = (XElement)CloneWithAnnotation(paragraphWithSplitRuns);
                        for (int id = 1; id < matchId; ++id)
                        {
                            List<XElement> elementsToReplace = paragraphWithReplacedRuns.Elements().Where(
                                e =>
                                {
                                    var sem = e.Annotation<MatchSemaphore>();
                                    if (sem == null) return false;
                                    return sem.MatchId == id;
                                }).ToList();

                            var replaceWithBreaks = new List<object>();

                            var newLineIndex = replace.IndexOf(Environment.NewLine);
                            int startIndex = 0;
                            while (newLineIndex != -1)
                            {
                                replaceWithBreaks.Add(
                                    new XElement(W.t, replace.Substring(startIndex, newLineIndex - startIndex)));
                                replaceWithBreaks.Add(new XElement(W.br));

                                startIndex = newLineIndex + Environment.NewLine.Length;
                                newLineIndex = replace.IndexOf(Environment.NewLine, startIndex);
                            }

                            replaceWithBreaks.Add(
                                new XElement(W.t, replace.Substring(startIndex, replace.Length - startIndex)));

                            var existRunProp = elementsToReplace.First().Elements(W.rPr);

                            if (rPrFormat != null)
                            {
                                foreach (XElement runProp in existRunProp)
                                {
                                    foreach (XElement newrPrNode in rPrFormat.Nodes())
                                    {
                                        runProp.Elements(newrPrNode.Name).Remove();
                                        runProp.Add(new XElement(newrPrNode));
                                    }
                                }
                            }

                            replaceWithBreaks.Insert(0, existRunProp);

                            elementsToReplace.First().AddBeforeSelf(new XElement(W.r, replaceWithBreaks.ToArray()));
                            elementsToReplace.Remove();
                        }

                        var groupedAdjacentRunsWithIdenticalFormatting =
                            paragraphWithReplacedRuns.Elements().GroupAdjacent(
                                ce =>
                                {
                                    if (ce.Name != W.r) return "DontConsolidate";
                                    if (ce.Elements().Where(e => e.Name != W.rPr).Count() != 1
                                        || ce.Element(W.t) == null) return "DontConsolidate";
                                    if (ce.Element(W.rPr) == null) return "";
                                    return ce.Element(W.rPr).ToString(SaveOptions.None);
                                });
                        XElement paragraphWithConsolidatedRuns = new XElement(
                            W.p,
                            groupedAdjacentRunsWithIdenticalFormatting.Select(
                                g =>
                                {
                                    if (g.Key == "DontConsolidate") return (object)g;
                                    string textValue = g.Select(r => r.Element(W.t).Value).StringConcatenate();
                                    XAttribute xs = null;
                                    if ((textValue.Length > 0 && textValue[0] == ' ')
                                        || (textValue.Length > 1 && textValue[textValue.Length - 1] == ' '))
                                    {
                                        xs = new XAttribute(XNamespace.Xml + "space", "preserve");
                                    }
                                    return new XElement(
                                        W.r,
                                        g.First().Elements(W.rPr),
                                        new XElement(W.t, xs, textValue));
                                }));
                        return paragraphWithConsolidatedRuns;
                    }
                    return element;
                }

                if (element.Name == W.r && element.Elements(W.t).Any())
                {
                    var collectionOfRuns = element.Elements().Where(e => e.Name != W.rPr).Select(
                        e =>
                        {
                            if (e.Name == W.t)
                            {
                                string s = (string)e;
                                IEnumerable<XElement> collectionOfSubRuns = s.Select(
                                    c =>
                                    {
                                        XElement newRun = new XElement(
                                                W.r,
                                                element.Elements(W.rPr),
                                                new XElement(
                                                    W.t,
                                                    c == ' '
                                                        ? new XAttribute(XNamespace.Xml + "space", "preserve")
                                                        : null,
                                                    c));
                                        return newRun;
                                    });
                                return (object)collectionOfSubRuns;
                            }
                            else
                            {
                                XElement newRun = new XElement(W.r, element.Elements(W.rPr), e);
                                return newRun;
                            }
                        });
                    return collectionOfRuns;
                }
                return new XElement(
                    element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => WmlSearchAndReplaceTransform(n, search, replace, matchCase, rPrFormat)));
            }
            return node;
        }

        public static XElement WmlFindFirst(XNode node, string search, XName xname = null, bool matchCase = false, int start = 0)
        {
            if (xname == null) xname = W.p;

            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == xname)
                {
                    string contents = element.Descendants(W.t).Select(t => (string)t).StringConcatenate();
                    if (contents.Contains(search) || (!matchCase && contents.ToUpper().Contains(search.ToUpper())))
                    {
                        return element;
                    }
                }

                foreach (var xNode in element.Nodes())
                {
                    var p = WmlFindFirst(xNode, search, xname, matchCase, start);
                    if (p != null)
                    {
                        return p;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Найти и заменить кучу вхождений
        /// <remarks>
        /// Используется дорогая, в плане производительности, процедура PutXDococument, 
        /// для исключения многих её вызовов добавлена эта процедура
        /// </remarks>
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <param name="replaces"></param>
        /// <param name="matchCase"></param>
        public static void SearchAndReplace(
            WordprocessingDocument wordDoc,
            Dictionary<string, string> replaces,
            bool matchCase)
        {
            if (RevisionAccepter.HasTrackedRevisions(wordDoc))
            {
                throw new InvalidDataException(
                    "Search and replace will not work with documents " + "that contain revision tracking.");
            }

            XDocument xDoc = wordDoc.MainDocumentPart.DocumentSettingsPart.GetXDocument();
            if (xDoc.Descendants(W.trackRevisions).Any())
            {
                throw new InvalidDataException("Revision tracking is turned on for document.");
            }

            xDoc = wordDoc.MainDocumentPart.GetXDocument();

            foreach (string key in replaces.Keys)
            {
                WmlSearchAndReplaceInXDocument(xDoc, key, replaces[key], matchCase);
            }

            wordDoc.MainDocumentPart.PutXDocument();

            foreach (var part in wordDoc.MainDocumentPart.HeaderParts)
            {
                xDoc = part.GetXDocument();
                foreach (string key in replaces.Keys)
                {
                    WmlSearchAndReplaceInXDocument(xDoc, key, replaces[key], matchCase);
                }

                part.PutXDocument();
            }

            foreach (var part in wordDoc.MainDocumentPart.FooterParts)
            {
                xDoc = part.GetXDocument();
                foreach (string key in replaces.Keys)
                {
                    WmlSearchAndReplaceInXDocument(xDoc, key, replaces[key], matchCase);
                }

                part.PutXDocument();
            }

            if (wordDoc.MainDocumentPart.EndnotesPart != null)
            {
                xDoc = wordDoc.MainDocumentPart.EndnotesPart.GetXDocument();
                foreach (string key in replaces.Keys)
                {
                    WmlSearchAndReplaceInXDocument(xDoc, key, replaces[key], matchCase);
                }

                wordDoc.MainDocumentPart.EndnotesPart.PutXDocument();
            }
            if (wordDoc.MainDocumentPart.FootnotesPart != null)
            {
                xDoc = wordDoc.MainDocumentPart.FootnotesPart.GetXDocument();
                foreach (string key in replaces.Keys)
                {
                    WmlSearchAndReplaceInXDocument(xDoc, key, replaces[key], matchCase);
                }

                wordDoc.MainDocumentPart.FootnotesPart.PutXDocument();
            }
        }
    }
}
