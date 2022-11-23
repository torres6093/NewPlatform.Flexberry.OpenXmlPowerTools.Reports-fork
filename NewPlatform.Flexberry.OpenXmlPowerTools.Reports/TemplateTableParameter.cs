namespace NewPlatform.Flexberry.Reports
{
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Xml.Linq;

    using DocumentFormat.OpenXml.Packaging;
    using OpenXmlPowerTools;

    /// <summary>
    /// Сложный параметр в шаблоне - таблица, в которой может быть много строк <!--<##Контрагенты  Контрагенты##>-->
    /// </summary>
    public class TemplateTableParameter
    {
        public TemplateTableParameter(string name, string tableRange)
        {
            Name = name;
            InnerParams = new List<TemplateParameter>();

            InnerTableParams = MatchParameters(tableRange);

            var regexField = new Regex(@"<#(?<param>.+?)#>");
            foreach (Match match in regexField.Matches(tableRange))
            {
                var matchValue = match.Groups["param"].Value;

                if ((matchValue.IndexOf("#") < 0 || matchValue.IndexOf(":") > 0 && matchValue.IndexOf("#") > matchValue.IndexOf(":")) && InnerParams.All(p => p.Name != matchValue) && !InInnerTable(matchValue))
                {
                    InnerParams.Add(new TemplateParameter(matchValue));
                }
            }
        }

        public static List<TemplateTableParameter> MatchParameters(string allText)
        {
            var list = new List<TemplateTableParameter>();

            if (allText.IndexOf("<##") < 0)
            {
                return list;
            }

            if (allText.IndexOf("##>") < 0)
            {
                return list;
            }

            allText = allText.Substring(allText.IndexOf("<##"));
            if (allText.LastIndexOf("##>") + 3 < allText.Length)
            {
                allText = allText.Remove(allText.LastIndexOf("##>") + 3);
            }

            var regexTable = new Regex(@"<##(?<param>[\w\d]+)");

            int beginning = 0;
            while (regexTable.IsMatch(allText, beginning))
            {
                var match = regexTable.Match(allText, beginning);
                var matchValue = match.Groups["param"].Value;

                var start = allText.IndexOf("<##" + matchValue);
                var end = allText.IndexOf(matchValue + "##>");

                var inner = string.Empty;
                if (start > -1 && end > -1 && list.All(p => p.Name != matchValue))
                {
                    inner = allText.Remove(end);
                    inner = inner.Substring(start).Replace("<##" + matchValue, string.Empty);

                    list.Add(new TemplateTableParameter(matchValue, inner));
                }

                beginning = end;
            }

            return list;
        }

        public string Name { get; private set; }

        public List<TemplateParameter> InnerParams { get; private set; }

        public List<TemplateTableParameter> InnerTableParams { get; private set; }

        private bool InInnerTable(string name)
        {
            if (InnerTableParams != null && InnerTableParams.Count > 0)
            {
                return InnerTableParams.Any(tp => tp.InnerParams.Any(p => p.FullName == name) || tp.InInnerTable(name));
            }

            return false;
        }

        public void ReplaceInXNode(XNode node, List<Dictionary<string, object>> rows)
        {
            string templateStartTableTag = "<##" + Name;
            string templateEndTableTag = Name + "##>";


            var startNode = TextReplacerExtensions.WmlFindFirst(node, templateStartTableTag);
            var endNode = TextReplacerExtensions.WmlFindFirst(node, templateEndTableTag);

            if (startNode == null || endNode == null)
            {
                return;
            }

            XElement commonParent = null;
            var startParent = startNode.Parent;
            var startFirstChild = startNode;
            var endFirstChild = endNode;
            while (startParent != null && commonParent == null)
            {
                var endParent = endNode.Parent;
                endFirstChild = endNode;
                while (endParent != null && commonParent == null)
                {
                    if (endParent == startParent)
                    {
                        commonParent = startParent;
                    }
                    else
                    {
                        endFirstChild = endParent;
                        endParent = endParent.Parent;
                    }
                }

                if (commonParent != null)
                {
                    continue;
                }

                startFirstChild = startParent;
                startParent = startParent.Parent;
            }

            Debug.Assert(commonParent != null, "Неправильный документ");

            foreach (Dictionary<string, object> row in rows)
            {
                var images = row.Keys.Where(x => x.StartsWith(TemplateImageParameter.ImgBookmarkPrefix)).ToList();

                var iteratorNode = (XNode)startFirstChild;
                do
                {
                    var newNode = new XElement((XElement)iteratorNode);
                    foreach (var innerParam in InnerParams)
                    {
                        var replace = row.ContainsKey(innerParam.Name) ? row[innerParam.Name] : null;

                        newNode =
                            TextReplacerExtensions.WmlSearchAndReplaceTransform(
                                newNode,
                                "<#" + innerParam.FullName + "#>",
                                replace == null ? string.Empty : innerParam.FormatObject(replace),
                                true) as XElement;
                    }

                    newNode = TextReplacerExtensions.WmlSearchAndReplaceTransform(newNode, templateStartTableTag, string.Empty, true) as XElement;
                    newNode = TextReplacerExtensions.WmlSearchAndReplaceTransform(newNode, templateEndTableTag, string.Empty, true) as XElement;

                    foreach (var image in images)
                    {
                        var hyperlink = TextReplacerExtensions.WmlFindFirst(newNode, image, W.hyperlink, true);
                        hyperlink?.SetAttributeValue("imgPath", ((ImageParameter)row[image]).FileName);
                        hyperlink?.SetAttributeValue("imgWidth", ((ImageParameter)row[image]).Width);
                        hyperlink?.SetAttributeValue("imgHeight", ((ImageParameter)row[image]).Height);
                    }

                    startFirstChild.AddBeforeSelf(newNode);

                    // бывает что все находится в одном паренте
                    if (iteratorNode == endFirstChild) break;

                    iteratorNode = iteratorNode.NextNode;
                }
                while (iteratorNode != endFirstChild);

                foreach (var innerTable in InnerTableParams)
                {
                    innerTable.ReplaceInXNode(
                        commonParent,
                        row.ContainsKey(innerTable.Name) ? (List<Dictionary<string, object>>)row[innerTable.Name] : new List<Dictionary<string, object>>());
                }
            }

            var newEndNode = new XElement(endFirstChild);
            foreach (var innerParam in InnerParams)
            {
                newEndNode =
                    TextReplacerExtensions.WmlSearchAndReplaceTransform(
                        newEndNode,
                        "<#" + innerParam.FullName + "#>",
                        string.Empty,
                        true) as XElement;
            }

            newEndNode = TextReplacerExtensions.WmlSearchAndReplaceTransform(newEndNode, templateStartTableTag, string.Empty, true) as XElement;
            newEndNode = TextReplacerExtensions.WmlSearchAndReplaceTransform(newEndNode, templateEndTableTag, string.Empty, true) as XElement;
            var newEndNodeContent = newEndNode.Descendants(W.t).Select(t => (string)t).StringConcatenate().Trim();
            if (!string.IsNullOrEmpty(newEndNodeContent))
            {
                startFirstChild.AddBeforeSelf(newEndNode);
            }

            // удалить все из парента
            var nextNode = (XNode)startFirstChild;
            while (nextNode != endFirstChild)
            {
                var nodeToDelete = nextNode;
                nextNode = nextNode.NextNode;
                nodeToDelete.Remove();
            }

            nextNode.Remove();
        }

        /// <summary>
        /// Заменить таблицу в указанном документе
        /// </summary>
        /// <param name="document"></param>
        /// <param name="parameters"></param>
        public void ReplaceInDocument(WordprocessingDocument document, List<Dictionary<string, object>> rows)
        {
            ReplaceInXNode(document.MainDocumentPart.GetXDocument().Root, rows);
        }
    }
}
