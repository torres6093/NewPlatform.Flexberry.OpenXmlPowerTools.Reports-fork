namespace NewPlatform.Flexberry.Reports
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Xml.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Drawing;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using global::OpenXmlPowerTools;
    using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
    using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
    using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
    using Run = DocumentFormat.OpenXml.Wordprocessing.Run;

    public class DocxReport
    {
        public readonly string NewLine = Environment.NewLine;

        private WmlDocument templateDoc;

        private readonly List<Source> resultDocs = new List<Source>();

        private readonly List<TemplateParameter> templateParameters = new List<TemplateParameter>();

        private List<TemplateTableParameter> templateTableParameters = new List<TemplateTableParameter>();

        private readonly List<TemplateImageParameter> templateImageParameters = new List<TemplateImageParameter>();

        /// <summary>
        /// Найти все параметры в по указанному паттерну в тексте
        /// </summary>
        /// <param name="text"></param>
        /// <param name="pattern"></param>
        /// <param name="patternGroup"></param>
        private List<string> MatchParameters(string text, string pattern, string patternGroup)
        {
            var regex = new Regex(pattern);
            return regex.Matches(text).Cast<Match>().Select(m => m.Groups[patternGroup].Value).ToList();
        }

        /// <summary>
        /// Открываем шаблон, ищем параметры
        /// </summary>
        public DocxReport(string templateFilePath)
        {
            // открыть файл шаблона
            var templateDocument = WordprocessingDocument.Open(templateFilePath, false);

            var mainPart = templateDocument.MainDocumentPart;

            // формируем сложные-табличные шаблоны (<##Имя ... <#Поле#> ... Имя##>)
            var allText = mainPart.Document.InnerText;
            templateTableParameters = TemplateTableParameter.MatchParameters(allText);

            // формируем простые шаблоны ([Имя])
            foreach (
                var parameter in
                    MatchParameters(allText, @"\[(?<param>.+?)\]", "param")
                        .Where(parameter => templateParameters.All(p => p.FullName != parameter)))
            {
                templateParameters.Add(new TemplateParameter(parameter));
            }


            foreach (var headerPart in mainPart.HeaderParts)
            {
                foreach (var parameter in
                    MatchParameters(headerPart.Header.InnerText, @"\[(?<param>.+?)\]", "param")
                        .Where(parameter => templateParameters.All(p => p.FullName != parameter)))
                {
                    templateParameters.Add(new TemplateParameter(parameter));
                }
            }

            foreach (var footerPart in mainPart.FooterParts)
            {
                foreach (var parameter in
                    MatchParameters(footerPart.Footer.InnerText, @"\[(?<param>.+?)\]", "param")
                        .Where(parameter => templateParameters.All(p => p.FullName != parameter)))
                {
                    templateParameters.Add(new TemplateParameter(parameter));
                }
            }

            // формируем шаблоны - рисунки
            foreach (var bookMark in mainPart.RootElement.Descendants<BookmarkStart>())
            {
                if (bookMark.Name.ToString().StartsWith(TemplateImageParameter.ImgBookmarkPrefix))
                {
                    templateImageParameters.Add(new TemplateImageParameter(bookMark.Name));
                }
            }

            templateDoc = new WmlDocument(templateFilePath);
        }

        /// <summary>
        /// Сохраняем и закрываем документы
        /// </summary>
        public void SaveAs(string saveFilePath)
        {
            DocumentBuilder.BuildDocument(resultDocs, saveFilePath);
        }

        public void SaveAs(Stream stream)
        {
            var wmlDocument = DocumentBuilder.BuildDocument(resultDocs);
            wmlDocument.WriteByteArray(stream);
        }

        /// <summary>
        /// Создать замены на основе созданных параметров и тех которые есть в шаблоне
        /// </summary>
        /// <param name="inputParameters"></param>
        /// <param name="replace"></param>
        protected void FormReplaceByExistsParameters(
            Dictionary<string, object> inputParameters,
            Dictionary<string, string> replace)
        {
            foreach (var templateParameter in templateParameters)
            {
                if (inputParameters.ContainsKey(templateParameter.Name))
                {
                    replace.Add(
                        "[" + templateParameter.FullName + "]", templateParameter.FormatObject(inputParameters[templateParameter.Name]));
                    continue;
                }

                // Может быть идет значение какого-либо поля параметра
                var parameterName = templateParameter.Name;
                if (parameterName.Contains("."))
                {
                    var objectName = parameterName.Substring(0, parameterName.IndexOf("."));

                    // если такого параметра нет значит пользователь скармливает сайту какую-то фигню, игнорируем его
                    if (!inputParameters.ContainsKey(objectName))
                    {
                        continue;
                    }

                    object dataObject = inputParameters[objectName];

                    // ковыряем значения полей через reflection
                    var fieldName = parameterName.Remove(0, objectName.Length + 1);

                    // пользователь опять может подсунуть какую-нибудь гадость
                    while (fieldName.Contains(".") && dataObject != null)
                    {
                        objectName = fieldName.Substring(0, fieldName.IndexOf("."));
                        dataObject = GetFieldValueWithReflection(dataObject, objectName);

                        fieldName = fieldName.Remove(0, objectName.Length + 1);
                    }

                    if (dataObject != null)
                    {
                        var fieldValue = GetFieldValueWithReflection(dataObject, fieldName);
                        replace.Add("[" + templateParameter.FullName + "]", templateParameter.FormatObject(fieldValue));
                    }
                    else
                    {
                        replace.Add("[" + templateParameter.FullName + "]", string.Empty);
                    }
                }
            }
        }

        protected void CheckTableParameters(TemplateTableParameter table, Dictionary<string, object> parameters, ref string message)
        {
            if (!parameters.ContainsKey(table.Name))
            {
                var r = string.Format("Не найден табличный параметр {0}; ", table.Name);
                if (!message.Contains(r)) message += r;
                return;
            }

            var curTableParameter = (List<Dictionary<string, object>>)parameters[table.Name];

            foreach (var param in table.InnerParams)
            {
                if (curTableParameter.Any(row => !row.ContainsKey(param.Name)))
                {
                    var r = string.Format("Не найден параметр {0} табличного параметра {1}; ", param.Name, table.Name);
                    if (!message.Contains(r)) message += r;
                }
            }

            foreach (var param in table.InnerTableParams)
            {
                foreach (var tparam in curTableParameter)
                {
                    CheckTableParameters(param, tparam, ref message);
                }
            }
        }

        public string BuildWithParameters(Dictionary<string, object> parameters)
        {
            // Сформировать параметры необходимые для текущего шаблона
            var replace = new Dictionary<string, string>();

            FormReplaceByExistsParameters(parameters, replace);

            var result = templateParameters.Where(parameter => !replace.ContainsKey("[" + parameter.FullName + "]")).Aggregate(string.Empty, (current, parameter) => current + string.Format("Не найден параметр {0}; ", parameter.Name));

            // создаем документ на основе шаблона
            WmlDocument wmlDoc;

            using (var streamDoc = new OpenXmlMemoryStreamDocument(templateDoc))
            {
                using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
                {
                    foreach (TemplateTableParameter table in templateTableParameters)
                    {
                        CheckTableParameters(table, parameters, ref result);

                        var rows = parameters.ContainsKey(table.Name)
                                       ? (List<Dictionary<string, object>>)parameters[table.Name]
                                       : new List<Dictionary<string, object>>();
                        table.ReplaceInDocument(document, rows);
                    }

                    TextReplacerExtensions.SearchAndReplace(document, replace, false);

                    // с закладками оказалось неудобно в случаях, когда необходимо вставить картинку в табличный параметр - закладки не копируются
                    // оставили этот код для совместимости
                    foreach (var templateImageParameter in templateImageParameters)
                    {
                        if (parameters.ContainsKey(templateImageParameter.Name))
                        {
                            var images = (List<ImageParameter>)parameters[templateImageParameter.Name];
                            foreach (var image in images)
                            {
                                InsertAPicture(document, image, templateImageParameter.FullName);
                            }
                        }
                    }

                    // тут вставляются картинки на место hyperlink, которые начинаются со спецпрефикса TemplateImageParameter.ImgBookmarkPrefix (imgTemplate)
                    AddImageParameters(document);
                }

                wmlDoc = streamDoc.GetModifiedWmlDocument();
            }

            // форматирование применять только к одному, сразу по всем отчетам при большом количестве договоров оно вываливается в StackOverflow
            wmlDoc = ApplyFormatting(wmlDoc);
            resultDocs.Add(new Source(wmlDoc, true));

            return result;
        }

        public static void InsertAPicture(WordprocessingDocument wordprocessingDocument, ImageParameter image, string bookMarkName)
        {
            var mainPart = wordprocessingDocument.MainDocumentPart;
            var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

            using (var stream = new FileStream(image.FileName, FileMode.Open))
            {
                imagePart.FeedData(stream);
                stream.Close();
            }

            AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart), bookMarkName, image.Width, image.Height);
            File.Delete(image.FileName);
        }

        private static Drawing GetImage(string name, int width, int height, string relationshipId)
        {
            int emuWidth = width * 9525;
            int emuHeight = height * 9525;

            var element =
                    new Drawing(
                        new DW.Inline(
                            new DW.Extent() { Cx = emuWidth, Cy = emuHeight },
                            new DW.EffectExtent()
                            {
                                LeftEdge = 0L,
                                TopEdge = 0L,
                                RightEdge = 0L,
                                BottomEdge = 0L
                            },
                            new DW.DocProperties()
                            {
                                Id = (UInt32Value)1U,
                                Name = name
                            },
                            new DW.NonVisualGraphicFrameDrawingProperties(new GraphicFrameLocks()
                            {
                                NoChangeAspect = true
                            }),
                            new Graphic(
                                new GraphicData(
                                    new PIC.Picture(
                                        new PIC.NonVisualPictureProperties(
                                            new PIC.NonVisualDrawingProperties()
                                            {
                                                Id = (UInt32Value)0U,
                                                Name = name
                                            },
                                            new PIC.NonVisualPictureDrawingProperties()),
                                        new PIC.BlipFill(
                                            new Blip(
                                                new BlipExtensionList(
                                                    new BlipExtension()
                                                    {
                                                        Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                    })
                                                )
                                            {
                                                Embed = relationshipId,
                                                CompressionState =
                                                    BlipCompressionValues.Print
                                            },
                                            new Stretch(
                                                new FillRectangle())),
                                        new PIC.ShapeProperties(
                                            new Transform2D(
                                                new Offset() { X = 0L, Y = 0L },
                                                new Extents() { Cx = emuWidth, Cy = emuHeight }),
                                            new PresetGeometry(new AdjustValueList()
                                                )
                                            { Preset = ShapeTypeValues.Rectangle }))
                                    )
                                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                            )
                        {
                            DistanceFromTop = (UInt32Value)0U,
                            DistanceFromBottom = (UInt32Value)0U,
                            DistanceFromLeft = (UInt32Value)0U,
                            DistanceFromRight = (UInt32Value)0U
                        });

            return element;
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId, string bookMarkName, int width, int height)
        {
            var element = GetImage(bookMarkName, width, height, relationshipId);

            var runImg = new Run();
            runImg.Append(element);

            var parImg = new Paragraph();
            parImg.Append(runImg);

            IDictionary<string, BookmarkStart> bookMarkMap = new Dictionary<string, BookmarkStart>();
            foreach (BookmarkStart bookMarkStart in wordDoc.MainDocumentPart.RootElement.Descendants<BookmarkStart>().Where(bookMarkStart => bookMarkStart.Name.Value == bookMarkName))
            {
                bookMarkMap[bookMarkStart.Name] = bookMarkStart;
            }

            foreach (BookmarkStart bookmarkStart in bookMarkMap.Values)
            {
                bookmarkStart.Parent.InsertAfterSelf(parImg);
            }
        }

        private static void AddImageParameters(WordprocessingDocument wordDoc)
        {
            foreach (
                DocumentFormat.OpenXml.Wordprocessing.Hyperlink hyperlink in
                    wordDoc.MainDocumentPart.RootElement.Descendants<DocumentFormat.OpenXml.Wordprocessing.Hyperlink>()
                        .Where(hyperlink => hyperlink.InnerText.StartsWith(TemplateImageParameter.ImgBookmarkPrefix)))
            {
                var mainPart = wordDoc.MainDocumentPart;
                var imagePart = mainPart.AddImagePart(ImagePartType.Bmp);

                var path = hyperlink.ExtendedAttributes.FirstOrDefault(x => x.LocalName == "imgPath").Value;

                if (string.IsNullOrEmpty(path)) continue;

                var width = int.Parse(hyperlink.ExtendedAttributes.FirstOrDefault(x => x.LocalName == "imgWidth").Value);
                var height = int.Parse(hyperlink.ExtendedAttributes.FirstOrDefault(x => x.LocalName == "imgHeight").Value);

                using (var stream = new FileStream(path, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                    stream.Close();
                }

                var element = GetImage(hyperlink.InnerText, width, height, mainPart.GetIdOfPart(imagePart));

                var runImg = new Run();
                runImg.Append(element);

                var parImg = new Paragraph();
                parImg.Append(runImg);

                hyperlink.Parent.InsertAfterSelf(parImg);
                hyperlink.Remove();
            }
        }

        protected WmlDocument ApplyFormatting(WmlDocument wmlDoc)
        {
            WmlDocument resultDoc = null;

            List<string> formattedBoldParameters;
            List<string> formattedNonBoldParameters;
            List<string> formattedRedParameters;
            List<string> formattedItalicParameters;
            List<string> formattedParagraphParameters;
            List<string> formattedDelParameters;
            List<ComplexFormatParameter> formattedSizeParameters;

            // заменить отформатированные куски
            // именно в таком порядке - вначале получение всех форматированных частей, затем их замена, доступ к document.InnerText вызывает перечитывание документа
            using (var streamDoc = new OpenXmlMemoryStreamDocument(wmlDoc))
            {
                using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
                {
                    var innerText = document.MainDocumentPart.Document.InnerText;
                    formattedBoldParameters =
                        MatchParameters(innerText, @"<b>(?<param>.*?)\</b>", "param")
                            .Distinct()
                            .ToList();

                    formattedNonBoldParameters =
                        MatchParameters(
                            innerText,
                            @"<nb>(?<param>.*?)</nb>",
                            "param").Distinct().ToList();

                    formattedRedParameters =
                        MatchParameters(
                            innerText,
                            @"<red>(?<param>.*?)</red>",
                            "param").Distinct().ToList();


                    var regex = new Regex(@"<(?<size>s(?<sd>\d+))>(?<param>.*?)</\k<size>>");
                    formattedSizeParameters =
                        regex.Matches(innerText)
                            .Cast<Match>()
                            .Select(
                                m =>
                                new ComplexFormatParameter
                                {
                                    Text = m.Groups["param"].Value,
                                    Additional = m.Groups["sd"].Value
                                })
                            .ToList();

                    formattedItalicParameters =
                        MatchParameters(innerText, @"<i>(?<param>.*?)\</i>", "param")
                            .Distinct()
                            .ToList();

                    formattedParagraphParameters =
                        MatchParameters(innerText, @"<Pr>(?<param>.*?)\</Pr>", "param")
                            .Distinct()
                            .ToList();

                    formattedDelParameters =
                        MatchParameters(innerText, @"<d>(?<param>.*?)\</d>", "param")
                            .Distinct()
                            .ToList();
                }
            }

            using (var streamDoc = new OpenXmlMemoryStreamDocument(wmlDoc))
            {
                using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
                {
                    foreach (string param in formattedBoldParameters)
                    {
                        TextReplacerExtensions.SearchAndReplace(
                            document,
                            "<b>" + param + "</b>",
                            param,
                            true,
                            new XElement(W.rPr, new XElement(W.b)));
                    }

                    foreach (string param in formattedNonBoldParameters)
                    {
                        TextReplacerExtensions.SearchAndReplace(
                            document,
                            "<nb>" + param + "</nb>",
                            param,
                            true,
                            new XElement(W.rPr, new XElement(W.b, new XAttribute(W.val, false))));
                    }

                    // тут можно поступить как с размером шрифта <colorFF0000>красный текст</colorFF0000>, если вдруг потребуется
                    foreach (string param in formattedRedParameters)
                    {
                        TextReplacerExtensions.SearchAndReplace(
                            document,
                            "<red>" + param + "</red>",
                            param,
                            true,
                            new XElement(W.rPr, new XElement(W.color, new XElement(W.val, "FF0000"))));
                    }

                    foreach (ComplexFormatParameter sizeParameter in formattedSizeParameters)
                    {
                        int size = int.Parse(sizeParameter.Additional) * 2;
                        TextReplacerExtensions.SearchAndReplace(
                            document,
                            string.Format("<s{0}>{1}</s{0}>", sizeParameter.Additional, sizeParameter.Text),
                            sizeParameter.Text,
                            true,
                            new XElement(
                                W.rPr,
                                new XElement(W.sz, new XAttribute(W.val, size)),
                                new XElement(W.szCs, new XAttribute(W.val, size))));
                    }

                    foreach (string param in formattedItalicParameters)
                    {
                        TextReplacerExtensions.SearchAndReplace(
                            document,
                            "<i>" + param + "</i>",
                            param,
                            true,
                            new XElement(W.rPr, new XElement(W.i)));
                    }

                    foreach (string param in formattedParagraphParameters)
                    {
                        var node = document.MainDocumentPart.GetXDocument().Root;
                        var oldNode = TextReplacerExtensions.WmlFindFirst(node, "<Pr>" + param + "</Pr>");
                        if (oldNode.Name == W.p)
                        {
                            var par = MatchParameters(param, @"<p>(?<param>.*?)\</p>", "param").Distinct().ToList();
                            var pPr = oldNode.Element(W.pPr);
                            var rPr = oldNode.Element(W.r).Element(W.rPr);
                            foreach (var p in par)
                            {
                                var newNode = new XElement(W.p, pPr, new XElement(W.r, rPr, new XElement(W.t, p)));
                                oldNode.AddBeforeSelf(newNode);
                            }
                            oldNode.Remove();
                            document.MainDocumentPart.PutXDocument();
                        }
                    }

                    foreach (string param in formattedDelParameters)
                    {
                        if (string.IsNullOrWhiteSpace(param))
                        {
                            var node = document.MainDocumentPart.GetXDocument().Root;
                            var oldNode = TextReplacerExtensions.WmlFindFirst(node, "<d>" + param + "</d>");
                            if (oldNode.Name == W.p)
                            {
                                oldNode.Remove();
                                document.MainDocumentPart.PutXDocument();
                            }
                        }
                    }
                }

                resultDoc = streamDoc.GetModifiedWmlDocument();
            }

            return resultDoc;
        }

        /// <summary>
        /// Получить значение поля используя Reflection
        /// </summary>
        /// <param name="dataObject"></param>
        /// <param name="fieldName"></param>
        /// <returns></returns>
        protected object GetFieldValueWithReflection(object dataObject, string fieldName)
        {
            var propeties = dataObject.GetType().GetProperties();
            var property = propeties.FirstOrDefault(f => f.Name == fieldName);
            if (property != null)
            {
                return property.GetValue(dataObject, null);
            }

            return null;
        }
    }
}
