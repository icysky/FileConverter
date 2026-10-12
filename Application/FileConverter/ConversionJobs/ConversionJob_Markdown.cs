// <copyright file="ConversionJob_Markdown.cs" company="AAllard">License: http://www.gnu.org/licenses/gpl.html GPL version 3.</copyright>

namespace FileConverter.ConversionJobs
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    using FileConverter.Diagnostics;

    using Markdig;
    using Markdig.Helpers;
    using Markdig.Syntax;
    using Markdig.Syntax.Inlines;

    using MarkdigDocument = Markdig.Syntax.MarkdownDocument;
    using MTable = Markdig.Extensions.Tables.Table;
    using MTableRow = Markdig.Extensions.Tables.TableRow;
    using MTableCell = Markdig.Extensions.Tables.TableCell;
    using MEmphasis = Markdig.Syntax.Inlines.EmphasisInline;
    using WColor = DocumentFormat.OpenXml.Wordprocessing.Color;
    using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
    using ADrawing = DocumentFormat.OpenXml.Drawing;
    using Debug = FileConverter.Diagnostics.Debug;

    public class ConversionJob_Markdown : ConversionJob
    {
        private const long MaximumImageWidthInEmu = 5943600; // ~6.5 inches.
        private const long EmuPerPixel = 9525; // 96 dpi.

        private const int PageWidthInTwips = 11906; // A4 portrait width.
        private const int PageHeightInTwips = 16838; // A4 portrait height.
        private const int PageMarginInTwips = 1418; // 2.5 cm margins.
        private const int UsablePageWidthInTwips = PageWidthInTwips - (2 * PageMarginInTwips); // 9070 twips.

        private static readonly Markdig.MarkdownPipeline markdownPipeline = new Markdig.MarkdownPipelineBuilder()
            .UseAdvancedExtensions()
            .Build();

        private MainDocumentPart currentMainDocumentPart;
        private int currentImageId;
        private int nextNumberingInstanceId = 1;
        private readonly Dictionary<ListBlock, int> orderedListNumberingIds = new Dictionary<ListBlock, int>();

        public ConversionJob_Markdown()
        {
        }

        public ConversionJob_Markdown(ConversionPreset conversionPreset, string inputFilePath) : base(conversionPreset, inputFilePath)
        {
        }

        protected override void Convert()
        {
            this.UserState = Properties.Resources.ConversionStateReadDocument;

            string markdownContent = File.ReadAllText(this.InputFilePath);

            this.UserState = Properties.Resources.ConversionStateConversion;
            Debug.Log($"Convert markdown document '{this.InputFilePath}' to docx.");

            MarkdigDocument document = Markdig.Markdown.Parse(markdownContent, markdownPipeline);

            this.currentImageId = 0;
            this.nextNumberingInstanceId = 1;
            this.orderedListNumberingIds.Clear();

            using (WordprocessingDocument docxDocument = WordprocessingDocument.Create(this.OutputFilePath, WordprocessingDocumentType.Document))
            {
                this.currentMainDocumentPart = docxDocument.AddMainDocumentPart();
                this.currentMainDocumentPart.Document = new Document(new Body());
                Body body = this.currentMainDocumentPart.Document.Body;

                this.CreateStylesPart(this.currentMainDocumentPart);
                this.CreateNumberingPart(this.currentMainDocumentPart);

                List<OpenXmlElement> elements = new List<OpenXmlElement>();
                this.RenderBlocks(document, elements, new RenderingContext());

                foreach (OpenXmlElement element in elements)
                {
                    body.AppendChild(element);
                }

                // Page setup. Orientation is stated explicitly so that the document never inherits
                // a previous page setup from the host application (Word/WPS "last used" settings).
                body.AppendChild(new SectionProperties(
                    new PageSize { Width = PageWidthInTwips, Height = PageHeightInTwips, Orient = PageOrientationValues.Portrait }, // A4 portrait.
                    new PageMargin { Top = PageMarginInTwips, Right = PageMarginInTwips, Bottom = PageMarginInTwips, Left = PageMarginInTwips, Header = 709, Footer = 709, Gutter = 0 }));

                this.currentMainDocumentPart.Document.Save();
            }

            this.currentMainDocumentPart = null;

            Debug.Log($"Markdown document converted: {this.OutputFilePath}.");
        }

        private class RenderingContext
        {
            public int QuoteDepth
            {
                get;
                set;
            }
        }

        private class InlineRenderingContext
        {
            public bool Bold
            {
                get;
                set;
            }

            public bool Italic
            {
                get;
                set;
            }

            public bool Strikethrough
            {
                get;
                set;
            }

            public string HyperlinkUrl
            {
                get;
                set;
            }

            public InlineRenderingContext Clone()
            {
                return new InlineRenderingContext
                {
                    Bold = this.Bold,
                    Italic = this.Italic,
                    Strikethrough = this.Strikethrough,
                    HyperlinkUrl = this.HyperlinkUrl,
                };
            }
        }

        private void RenderBlocks(ContainerBlock container, List<OpenXmlElement> elements, RenderingContext context)
        {
            this.RenderBlocks((IEnumerable<Block>)container, elements, context);
        }

        private void RenderBlocks(IEnumerable<Block> blocks, List<OpenXmlElement> elements, RenderingContext context)
        {
            foreach (Block block in blocks)
            {
                if (this.CancelIsRequested)
                {
                    return;
                }

                HeadingBlock headingBlock = block as HeadingBlock;
                if (headingBlock != null)
                {
                    elements.Add(this.RenderHeading(headingBlock));
                    continue;
                }

                ParagraphBlock paragraphBlock = block as ParagraphBlock;
                if (paragraphBlock != null)
                {
                    elements.Add(this.RenderParagraph(paragraphBlock, context, null));
                    continue;
                }

                ListBlock listBlock = block as ListBlock;
                if (listBlock != null)
                {
                    this.RenderList(listBlock, elements, context);
                    continue;
                }

                CodeBlock codeBlock = block as CodeBlock;
                if (codeBlock != null)
                {
                    elements.Add(this.RenderCodeBlock(codeBlock));
                    continue;
                }

                QuoteBlock quoteBlock = block as QuoteBlock;
                if (quoteBlock != null)
                {
                    context.QuoteDepth++;
                    this.RenderBlocks(quoteBlock, elements, context);
                    context.QuoteDepth--;
                    continue;
                }

                ThematicBreakBlock thematicBreakBlock = block as ThematicBreakBlock;
                if (thematicBreakBlock != null)
                {
                    elements.Add(this.RenderThematicBreak());
                    continue;
                }

                MTable table = block as MTable;
                if (table != null)
                {
                    elements.Add(this.RenderTable(table, context));
                    continue;
                }

                if (block is HtmlBlock)
                {
                    Debug.Log($"Warning: html block ignored in '{this.InputFilePath}'.");
                    continue;
                }

                ContainerBlock containerBlock = block as ContainerBlock;
                if (containerBlock != null)
                {
                    this.RenderBlocks((IEnumerable<Block>)containerBlock, elements, context);
                    continue;
                }

                Debug.Log($"Warning: unsupported markdown block type '{block.GetType().Name}' ignored.");
            }
        }

        private Paragraph RenderHeading(HeadingBlock block)
        {
            List<OpenXmlElement> elements = new List<OpenXmlElement>();
            this.RenderInlines(block.Inline, elements, new InlineRenderingContext());

            Paragraph paragraph = new Paragraph();
            paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "Heading" + Math.Min(Math.Max(block.Level, 1), 6) });
            paragraph.Append(elements);
            return paragraph;
        }

        private Paragraph RenderParagraph(ParagraphBlock block, RenderingContext context, NumberingProperties numbering)
        {
            List<OpenXmlElement> elements = new List<OpenXmlElement>();
            this.RenderInlines(block.Inline, elements, new InlineRenderingContext());

            Paragraph paragraph = new Paragraph();

            if (numbering != null || context.QuoteDepth > 0)
            {
                ParagraphProperties paragraphProperties = new ParagraphProperties();

                if (numbering != null)
                {
                    paragraphProperties.AppendChild(numbering);
                }

                if (context.QuoteDepth > 0)
                {
                    paragraphProperties.AppendChild(new ParagraphBorders(new LeftBorder { Val = BorderValues.Single, Size = 12, Space = 8, Color = "A6A6A6" }));
                    paragraphProperties.AppendChild(new Indentation { Left = (720 * context.QuoteDepth).ToString() });
                }

                paragraph.ParagraphProperties = paragraphProperties;
            }

            paragraph.Append(elements);
            return paragraph;
        }

        private void RenderList(ListBlock listBlock, List<OpenXmlElement> elements, RenderingContext context)
        {
            int level = 0;
            Block ancestor = listBlock.Parent;
            while (ancestor != null)
            {
                ancestor = ancestor.Parent;
                if (ancestor is ListBlock)
                {
                    level++;
                }
            }

            int numberingId;
            if (listBlock.IsOrdered)
            {
                if (!this.orderedListNumberingIds.TryGetValue(listBlock, out numberingId))
                {
                    numberingId = this.CreateOrderedListNumberingInstance();
                    this.orderedListNumberingIds.Add(listBlock, numberingId);
                }
            }
            else
            {
                numberingId = 1; // Bullet list numbering instance created in the numbering part.
            }

            foreach (Block block in listBlock)
            {
                ListItemBlock listItem = block as ListItemBlock;
                if (listItem == null)
                {
                    continue;
                }

                foreach (Block itemBlock in listItem)
                {
                    if (itemBlock is ParagraphBlock paragraphBlock)
                    {
                        NumberingProperties numbering = new NumberingProperties
                        {
                            NumberingLevelReference = new NumberingLevelReference { Val = Math.Min(level, 2) },
                            NumberingId = new NumberingId { Val = numberingId },
                        };

                        elements.Add(this.RenderParagraph(paragraphBlock, context, numbering));
                    }
                    else
                    {
                        this.RenderBlocks(new[] { itemBlock }, elements, context);
                    }
                }
            }
        }

        private Paragraph RenderCodeBlock(CodeBlock block)
        {
            Paragraph paragraph = new Paragraph();
            paragraph.ParagraphProperties = new ParagraphProperties(
                new Shading { Val = ShadingPatternValues.Clear, Fill = "F2F2F2" },
                new SpacingBetweenLines { After = "160", Line = "240", LineRule = LineSpacingRuleValues.Auto },
                new Indentation { Left = "227" });

            bool hasContent = false;
            foreach (StringSlice line in block.Lines.Lines)
            {
                if (hasContent)
                {
                    paragraph.AppendChild(new Run(new Break()));
                }

                string text = line.ToString();
                if (!string.IsNullOrEmpty(text))
                {
                    paragraph.AppendChild(this.CreateCodeRun(text));
                }

                hasContent = true;
            }

            if (!hasContent)
            {
                paragraph.AppendChild(this.CreateCodeRun(" "));
            }

            return paragraph;
        }

        private Paragraph RenderThematicBreak()
        {
            Paragraph paragraph = new Paragraph();
            paragraph.ParagraphProperties = new ParagraphProperties(
                new ParagraphBorders(new BottomBorder { Val = BorderValues.Single, Size = 6, Space = 1, Color = "999999" }),
                new SpacingBetweenLines { After = "240" });
            paragraph.AppendChild(new Run(new Text("")));
            return paragraph;
        }

        private Table RenderTable(MTable table, RenderingContext context)
        {
            Table docxTable = new Table();

            // Count the maximum number of grid columns (a cell can span several columns).
            int columnCount = 1;
            foreach (Block block in table)
            {
                MTableRow row = block as MTableRow;
                if (row == null)
                {
                    continue;
                }

                int rowColumnCount = 0;
                foreach (Block cellBlock in row)
                {
                    MTableCell cell = cellBlock as MTableCell;
                    if (cell != null)
                    {
                        rowColumnCount += cell.ColumnSpan > 1 ? cell.ColumnSpan : 1;
                    }
                }

                columnCount = Math.Max(columnCount, rowColumnCount);
            }

            // Share the usable page width between all columns, so that the table can never be
            // wider than the page (even with many columns or long unbreakable content).
            int columnWidth = UsablePageWidthInTwips / columnCount;
            if (columnWidth < 1)
            {
                columnWidth = 1;
            }

            int lastColumnWidth = UsablePageWidthInTwips - (columnWidth * (columnCount - 1));
            if (lastColumnWidth < 1)
            {
                lastColumnWidth = columnWidth;
            }

            // A fixed layout makes the host application respect the widths below instead of
            // auto-fitting the table to its content (which used to overflow the page).
            docxTable.AppendChild(new TableProperties(
                new TableStyle { Val = "TableGrid" },
                new TableWidth { Width = UsablePageWidthInTwips.ToString(), Type = TableWidthUnitValues.Dxa },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4, Color = "BFBFBF" },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "BFBFBF" },
                    new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "BFBFBF" },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Color = "BFBFBF" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "BFBFBF" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "BFBFBF" }),
                new TableLayout { Type = TableLayoutValues.Fixed },
                new TableLook { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true }));

            TableGrid tableGrid = new TableGrid();
            for (int index = 0; index < columnCount; index++)
            {
                tableGrid.AppendChild(new GridColumn { Width = (index == columnCount - 1 ? lastColumnWidth : columnWidth).ToString() });
            }

            docxTable.AppendChild(tableGrid);

            foreach (Block block in table)
            {
                MTableRow markdownRow = block as MTableRow;
                if (markdownRow == null)
                {
                    continue;
                }

                TableRow docxRow = new TableRow();
                if (markdownRow.IsHeader)
                {
                    docxRow.TableRowProperties = new TableRowProperties(new TableHeader());
                }

                int columnIndex = 0;
                foreach (Block cellBlock in markdownRow)
                {
                    MTableCell markdownCell = cellBlock as MTableCell;
                    if (markdownCell == null)
                    {
                        continue;
                    }

                    int columnSpan = markdownCell.ColumnSpan > 1 ? markdownCell.ColumnSpan : 1;

                    // Cell width must match the columns it covers, otherwise a fixed layout
                    // would squeeze or overflow the table.
                    int cellWidth = 0;
                    for (int index = 0; index < columnSpan && columnIndex + index < columnCount; index++)
                    {
                        cellWidth += columnIndex + index == columnCount - 1 ? lastColumnWidth : columnWidth;
                    }

                    if (cellWidth < 1)
                    {
                        cellWidth = columnWidth;
                    }

                    TableCell cell = new TableCell();
                    cell.TableCellProperties = new TableCellProperties(new TableCellWidth { Width = cellWidth.ToString(), Type = TableWidthUnitValues.Dxa });
                    if (columnSpan > 1)
                    {
                        cell.TableCellProperties.AppendChild(new GridSpan { Val = columnSpan });
                    }

                    columnIndex += columnSpan;

                    if (markdownRow.IsHeader)
                    {
                        cell.TableCellProperties.AppendChild(new Shading { Val = ShadingPatternValues.Clear, Fill = "DEEAF6" });
                    }

                    List<OpenXmlElement> cellElements = new List<OpenXmlElement>();
                    this.RenderBlocks(markdownCell, cellElements, context);

                    if (cellElements.Count == 0)
                    {
                        cell.AppendChild(new Paragraph());
                    }
                    else
                    {
                        foreach (OpenXmlElement cellElement in cellElements)
                        {
                            if (markdownRow.IsHeader)
                            {
                                foreach (Run run in cellElement.Descendants<Run>())
                                {
                                    RunProperties runProperties = run.GetFirstChild<RunProperties>();
                                    if (runProperties == null)
                                    {
                                        runProperties = new RunProperties();
                                        run.PrependChild(runProperties);
                                    }

                                    if (runProperties.Bold == null)
                                    {
                                        runProperties.Bold = new Bold();
                                    }
                                }
                            }

                            cell.AppendChild(cellElement);
                        }
                    }

                    docxRow.AppendChild(cell);
                }

                docxTable.AppendChild(docxRow);
            }

            return docxTable;
        }

        private void RenderInlines(ContainerInline inlines, List<OpenXmlElement> elements, InlineRenderingContext context)
        {
            if (inlines == null)
            {
                return;
            }

            foreach (Inline inline in inlines)
            {
                this.RenderInline(inline, elements, context);
            }
        }

        private void RenderInline(Inline inline, List<OpenXmlElement> elements, InlineRenderingContext context)
        {
            LiteralInline literalInline = inline as LiteralInline;
            if (literalInline != null)
            {
                this.AddTextElement(elements, this.CreateTextRun(literalInline.Content.ToString(), context));
                return;
            }

            CodeInline codeInline = inline as CodeInline;
            if (codeInline != null)
            {
                elements.Add(this.CreateCodeRun(codeInline.Content));
                return;
            }

            MEmphasis emphasis = inline as MEmphasis;
            if (emphasis != null)
            {
                InlineRenderingContext emphasisContext = context.Clone();
                if (emphasis.DelimiterChar == '~')
                {
                    emphasisContext.Strikethrough = true;
                }
                else
                {
                    if (emphasis.DelimiterCount >= 2)
                    {
                        emphasisContext.Bold = true;
                    }

                    if (emphasis.DelimiterCount % 2 == 1)
                    {
                        emphasisContext.Italic = true;
                    }
                }

                this.RenderInlines((ContainerInline)emphasis, elements, emphasisContext);
                return;
            }

            LinkInline linkInline = inline as LinkInline;
            if (linkInline != null)
            {
                if (linkInline.IsImage)
                {
                    Run imageRun = this.CreateImageRun(linkInline);
                    if (imageRun != null)
                    {
                        elements.Add(imageRun);
                    }

                    return;
                }

                InlineRenderingContext linkContext = context.Clone();
                linkContext.HyperlinkUrl = linkInline.Url;

                string linkText = this.ExtractLiteralText(linkInline);
                if (string.IsNullOrEmpty(linkText))
                {
                    linkText = linkInline.Url;
                }

                this.AddTextElement(elements, this.CreateTextRun(linkText, linkContext));
                return;
            }

            LineBreakInline lineBreak = inline as LineBreakInline;
            if (lineBreak != null)
            {
                elements.Add(new Run(new Break()));
                return;
            }

            AutolinkInline autolinkInline = inline as AutolinkInline;
            if (autolinkInline != null)
            {
                InlineRenderingContext autolinkContext = context.Clone();
                autolinkContext.HyperlinkUrl = autolinkInline.IsEmail ? "mailto:" + autolinkInline.Url : autolinkInline.Url;
                this.AddTextElement(elements, this.CreateTextRun(autolinkInline.Url, autolinkContext));
                return;
            }

            HtmlInline htmlInline = inline as HtmlInline;
            if (htmlInline != null)
            {
                Debug.Log($"Warning: html inline tag '{htmlInline.Tag}' ignored in '{this.InputFilePath}'.");
                return;
            }

            HtmlEntityInline htmlEntityInline = inline as HtmlEntityInline;
            if (htmlEntityInline != null)
            {
                string entityText = htmlEntityInline.Original.ToString();
                this.AddTextElement(elements, this.CreateTextRun(entityText, context));
                return;
            }

            ContainerInline containerInline = inline as ContainerInline;
            if (containerInline != null)
            {
                this.RenderInlines(containerInline, elements, context);
                return;
            }

            if (!(inline is Markdig.Extensions.TaskLists.TaskList))
            {
                Debug.Log($"Warning: unsupported markdown inline type '{inline.GetType().Name}' ignored.");
            }
        }

        private void AddTextElement(List<OpenXmlElement> elements, OpenXmlElement element)
        {
            if (element != null)
            {
                elements.Add(element);
            }
        }

        private string ExtractLiteralText(ContainerInline container)
        {
            string text = string.Empty;
            if (container == null)
            {
                return text;
            }

            foreach (Inline inline in container)
            {
                LiteralInline literal = inline as LiteralInline;
                if (literal != null)
                {
                    text += literal.Content.ToString();
                }
                else if (inline is CodeInline code)
                {
                    text += code.Content;
                }
                else if (inline is ContainerInline childContainer)
                {
                    text += this.ExtractLiteralText(childContainer);
                }
            }

            return text;
        }

        private OpenXmlElement CreateTextRun(string text, InlineRenderingContext context)
        {
            if (string.IsNullOrEmpty(text))
            {
                return null;
            }

            RunProperties runProperties = new RunProperties();
            bool hasProperties = false;

            if (context.Bold)
            {
                runProperties.Bold = new Bold();
                hasProperties = true;
            }

            if (context.Italic)
            {
                runProperties.Italic = new Italic();
                hasProperties = true;
            }

            if (context.Strikethrough)
            {
                runProperties.Strike = new Strike();
                hasProperties = true;
            }

            if (!string.IsNullOrEmpty(context.HyperlinkUrl))
            {
                runProperties.Color = new WColor { Val = "0563C1" };
                runProperties.Underline = new Underline { Val = UnderlineValues.Single };
                hasProperties = true;
            }

            Run run = new Run();
            if (hasProperties)
            {
                run.RunProperties = runProperties;
            }

            run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

            if (string.IsNullOrEmpty(context.HyperlinkUrl))
            {
                return run;
            }

            // Web links are rendered as real hyperlinks, anything else stays plain text.
            string url = context.HyperlinkUrl;
            if (!Uri.IsWellFormedUriString(url, UriKind.Absolute) ||
                !(url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                  url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                  url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)))
            {
                return run;
            }

            try
            {
                HyperlinkRelationship relationship = this.currentMainDocumentPart.AddHyperlinkRelationship(new Uri(url, UriKind.Absolute), true);
                return new Hyperlink(run) { Id = relationship.Id, History = true };
            }
            catch (Exception exception)
            {
                Debug.Log($"Warning: fail to create hyperlink '{url}': {exception.Message}");
                return run;
            }
        }

        private Run CreateCodeRun(string text)
        {
            RunProperties runProperties = new RunProperties(
                new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas" },
                new FontSize { Val = "19" },
                new Shading { Val = ShadingPatternValues.Clear, Fill = "F2F2F2" });

            return new Run(runProperties, new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        }

        private Run CreateImageRun(LinkInline imageLink)
        {
            string imageUrl = imageLink.Url;
            string imagePath = Uri.UnescapeDataString(imageUrl);

            if (Path.IsPathRooted(imagePath))
            {
                if (!File.Exists(imagePath))
                {
                    Debug.Log($"Warning: image '{imageUrl}' not found, ignored.");
                    return null;
                }
            }
            else
            {
                imagePath = Path.Combine(Path.GetDirectoryName(this.InputFilePath), imagePath);
                if (!File.Exists(imagePath))
                {
                    Debug.Log($"Warning: image '{imageUrl}' not found, ignored.");
                    return null;
                }
            }

            try
            {
                byte[] imageBytes = File.ReadAllBytes(imagePath);
                string relationshipId = this.GetOrCreateImagePart(imagePath, imageBytes);

                long widthEmu;
                long heightEmu;
                using (MemoryStream stream = new MemoryStream(imageBytes))
                using (Image image = Image.FromStream(stream))
                {
                    widthEmu = image.Width * EmuPerPixel;
                    heightEmu = image.Height * EmuPerPixel;
                    if (widthEmu > MaximumImageWidthInEmu)
                    {
                        heightEmu = heightEmu * MaximumImageWidthInEmu / widthEmu;
                        widthEmu = MaximumImageWidthInEmu;
                    }
                }

                this.currentImageId++;

                string altText = this.ExtractLiteralText(imageLink);
                var picture = new ADrawing.Picture(
                    new ADrawing.NonVisualPictureProperties(
                        new ADrawing.NonVisualDrawingProperties { Id = (UInt32Value)(uint)this.currentImageId, Name = $"Image {this.currentImageId}", Description = altText },
                        new ADrawing.NonVisualPictureDrawingProperties()),
                    new ADrawing.BlipFill(
                        new ADrawing.Blip { Embed = relationshipId },
                        new ADrawing.Stretch(new ADrawing.FillRectangle())),
                    new ADrawing.ShapeProperties(
                        new ADrawing.Transform2D(
                            new ADrawing.Offset { X = 0, Y = 0 },
                            new ADrawing.Extents { Cx = widthEmu, Cy = heightEmu }),
                        new ADrawing.PresetGeometry(new ADrawing.AdjustValueList()) { Preset = ADrawing.ShapeTypeValues.Rectangle }));

                DW.Inline inline = new DW.Inline();

                inline.AppendChild(new DW.Extent { Cx = widthEmu, Cy = heightEmu });
                inline.AppendChild(new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 });
                inline.AppendChild(new DW.DocProperties { Id = (UInt32Value)(uint)this.currentImageId, Name = $"Picture {this.currentImageId}", Description = altText });
                inline.AppendChild(new ADrawing.Graphic(new ADrawing.GraphicData(picture) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }));

                return new Run(new Drawing(inline));
            }
            catch (Exception exception)
            {
                Debug.Log($"Warning: fail to embed image '{imageUrl}': {exception.Message}");
                return null;
            }
        }

        private string GetOrCreateImagePart(string imagePath, byte[] imageBytes)
        {
            PartTypeInfo imagePartType = GetImagePartType(imagePath, imageBytes);
            ImagePart imagePart = this.currentMainDocumentPart.AddImagePart(imagePartType);
            using (MemoryStream stream = new MemoryStream(imageBytes))
            {
                imagePart.FeedData(stream);
            }

            return this.currentMainDocumentPart.GetIdOfPart(imagePart);
        }

        private static PartTypeInfo GetImagePartType(string imagePath, byte[] imageBytes)
        {
            switch (Path.GetExtension(imagePath).ToLowerInvariant())
            {
                case ".jpg":
                case ".jpeg":
                    return ImagePartType.Jpeg;
                case ".gif":
                    return ImagePartType.Gif;
                case ".bmp":
                    return ImagePartType.Bmp;
                case ".tif":
                case ".tiff":
                    return ImagePartType.Tiff;
                case ".png":
                    return ImagePartType.Png;
            }

            // Unknown extension: detect the real format from the file header.
            using (MemoryStream stream = new MemoryStream(imageBytes))
            using (Image image = Image.FromStream(stream))
            {
                if (image.RawFormat.Guid == System.Drawing.Imaging.ImageFormat.Jpeg.Guid)
                {
                    return ImagePartType.Jpeg;
                }

                if (image.RawFormat.Guid == System.Drawing.Imaging.ImageFormat.Gif.Guid)
                {
                    return ImagePartType.Gif;
                }

                if (image.RawFormat.Guid == System.Drawing.Imaging.ImageFormat.Bmp.Guid)
                {
                    return ImagePartType.Bmp;
                }

                if (image.RawFormat.Guid == System.Drawing.Imaging.ImageFormat.Tiff.Guid)
                {
                    return ImagePartType.Tiff;
                }
            }

            return ImagePartType.Png;
        }

        private void CreateStylesPart(MainDocumentPart mainPart)
        {
            StyleDefinitionsPart stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles();

            stylesPart.Styles.AppendChild(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "Calibri", ComplexScript = "Calibri" },
                    new FontSize { Val = "22" },
                    new FontSizeComplexScript { Val = "22" })),
                new ParagraphPropertiesDefault(new ParagraphPropertiesBaseStyle(
                    new SpacingBetweenLines { After = "160", Line = "259", LineRule = LineSpacingRuleValues.Auto }))));

            stylesPart.Styles.AppendChild(new Style(new StyleName { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });

            string[] headingSizes = { "32", "26", "24", "22", "22", "22" };
            for (int index = 1; index <= 6; index++)
            {
                StyleRunProperties styleRunProperties = new StyleRunProperties(new WColor { Val = "2E74B5" }, new FontSize { Val = headingSizes[index - 1] });
                if (index >= 4)
                {
                    styleRunProperties.PrependChild(new Italic());
                }

                stylesPart.Styles.AppendChild(new Style(
                    new StyleName { Val = "heading " + index },
                    new BasedOn { Val = "Normal" },
                    new NextParagraphStyle { Val = "Normal" },
                    new PrimaryStyle(),
                    new ParagraphProperties(
                        new KeepNext(),
                        new KeepLines(),
                        new SpacingBetweenLines { Before = "240", After = "0" },
                        new OutlineLevel { Val = index - 1 }),
                    styleRunProperties)
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Heading" + index,
                });
            }

            stylesPart.Styles.Save();
        }

        private void CreateNumberingPart(MainDocumentPart mainPart)
        {
            NumberingDefinitionsPart numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            Numbering numbering = new Numbering();

            numbering.AppendChild(CreateListAbstractNum(1, false));
            numbering.AppendChild(CreateListAbstractNum(2, true));
            numbering.AppendChild(new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });

            numberingPart.Numbering = numbering;
        }

        private static AbstractNum CreateListAbstractNum(int abstractNumId, bool ordered)
        {
            string[] bulletCharacters = { "•", "◦", "▪" };

            AbstractNum abstractNum = new AbstractNum { AbstractNumberId = abstractNumId };
            for (int level = 0; level < 3; level++)
            {
                Level numberingLevel = new Level { LevelIndex = level };
                numberingLevel.AppendChild(new StartNumberingValue { Val = 1 });

                if (ordered)
                {
                    numberingLevel.AppendChild(new NumberingFormat { Val = NumberFormatValues.Decimal });
                    numberingLevel.AppendChild(new LevelText { Val = "%" + (level + 1) + "." });
                }
                else
                {
                    numberingLevel.AppendChild(new NumberingFormat { Val = NumberFormatValues.Bullet });
                    numberingLevel.AppendChild(new LevelText { Val = bulletCharacters[level] });
                }

                numberingLevel.AppendChild(new LevelJustification { Val = LevelJustificationValues.Left });
                numberingLevel.AppendChild(new PreviousParagraphProperties(new Indentation { Left = ((level + 1) * 720).ToString(), Hanging = "360" }));

                abstractNum.AppendChild(numberingLevel);
            }

            return abstractNum;
        }

        private int CreateOrderedListNumberingInstance()
        {
            this.nextNumberingInstanceId++;

            Numbering numbering = this.currentMainDocumentPart.NumberingDefinitionsPart.Numbering;
            numbering.AppendChild(new NumberingInstance(
                new AbstractNumId { Val = 2 },
                new LevelOverride(new StartOverrideNumberingValue { Val = 1 }) { LevelIndex = 0 })
            {
                NumberID = this.nextNumberingInstanceId,
            });

            return this.nextNumberingInstanceId;
        }
    }
}
