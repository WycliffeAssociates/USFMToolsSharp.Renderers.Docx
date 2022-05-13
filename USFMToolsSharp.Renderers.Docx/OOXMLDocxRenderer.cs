using System.Collections.Generic;
using USFMToolsSharp.Models.Markers;
using USFMToolsSharp.Renderers.Docx.Extensions;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using System;

namespace USFMToolsSharp.Renderers.Docx
{

    public class OOXMLDocxRenderer
    {
        public List<string> UnrenderableMarkers;
        public Dictionary<string, Marker> CrossRefMarkers;
        private DocxConfig configDocx;
        private Body body;
        private Footnotes footnotes;
        private WordprocessingDocument newDoc;
        private int pageHeaderCount = 1;
        private string previousBookHeader = null;
        private const string chapterLabelDefault = "Chapter";
        private string chapterLabel = chapterLabelDefault;
        private string currentChapterLabel = "";
        private bool beforeFirstChapter = true;
        private int nextFootnoteNum = 1;
        private Marker thisMarker = null;
        private Marker previousMarker = null;
        private Paragraph lastAppendedParagraph;

        public OOXMLDocxRenderer()
        {
            configDocx = new DocxConfig();
        }
        public OOXMLDocxRenderer(DocxConfig config)
        {
            configDocx = config;
        }

        public USFMDocument FrontMatter { get; set; } = null;

        public Stream Render(USFMDocument input)
        {
            var outputStream = new MemoryStream();
            UnrenderableMarkers = new List<string>();
            CrossRefMarkers = new Dictionary<string, Marker>();
            using (newDoc = WordprocessingDocument.Create(outputStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var mainPart = newDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings();
                settingsPart.Settings.Append(new UpdateFieldsOnOpen() { Val = true });

                var footnotesPart = mainPart.AddNewPart<FootnotesPart>();
                footnotes = footnotesPart.Footnotes = new Footnotes();
                body = mainPart.Document.AppendChild(new Body());
                var stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                var styles = stylePart.Styles = new Styles();
                AddStyles(styles);

                if (configDocx.renderTableOfContents)
                {
                    RenderTOC();
                }

                if (FrontMatter != null)
                {
                    // TODO: Come back to this RenderFrontMatter(FrontMatter);
                }

                

                foreach (Marker marker in input.Contents)
                {
                    RenderMarker(marker, new StyleConfig());
                }

                // Add section header for final book
                if (previousBookHeader != null)
                {
                    //createBookHeaders(previousBookHeader);
                }

                // Make final document section continuous so that it doesn't
                // create an extra page at the end.  Final section is unique:
                // it's a direct child of the document, not a child of the last
                // paragraph.

                var sectionProperties = AppendToBody(new SectionProperties());
                var sectionType = sectionProperties.AppendChild(new SectionType());
                sectionType.Val = SectionMarkValues.Continuous;
                if (configDocx.showPageNumbers)
                {
                    var pageNumber = sectionProperties.AppendChild(new PageNumberType());
                    pageNumber.Format = NumberFormatValues.Decimal;
                }
                var columns = sectionProperties.AppendChild(new Columns());
                columns.ColumnCount = (Int16Value)configDocx.columnCount;
            }
            return outputStream;
        }
        private Paragraph AppendToBody(Paragraph item)
        {
            lastAppendedParagraph = item;
            return body.AppendChild(item);
        }
        private SectionProperties AppendToBody(SectionProperties item)
        {
            return body.AppendChild(item);
        }

        private void AddStyles(Styles styles)
        {
            Style normalStyle = new Style
            {
                StyleId = "Normal",
                Default = true,
                Type = StyleValues.Paragraph
            };

            normalStyle.AppendChild(new StyleName { Val = "Normal" });
            normalStyle.AppendChild(new PrimaryStyle());
            Style headingStyle = styles.AppendChild(new Style
            {
                StyleId = "BookHeading",
                Type = StyleValues.Paragraph
            });

            headingStyle.AppendChild(new StyleName { Val = "Book Heading" });
            headingStyle.AppendChild(new BasedOn() { Val = "Normal" });

            headingStyle.AppendChild(new NextParagraphStyle { Val = "Normal" });

            headingStyle.AppendChild(new LinkedStyle { Val = "Heading1Char" });

            headingStyle.AppendChild(new UIPriority { Val = 9 });

            headingStyle.AppendChild(new PrimaryStyle());


            StyleParagraphProperties headingStyleParagraphProperties = headingStyle.AppendChild(new StyleParagraphProperties());
            headingStyleParagraphProperties.AppendChild(new KeepNext());
            headingStyleParagraphProperties.AppendChild(new KeepLines());

            headingStyleParagraphProperties.AppendChild(new SpacingBetweenLines
            {
                Before = "240",
                After = "0"
            });

            headingStyleParagraphProperties.AppendChild(new OutlineLevel { Val = 0 });

            StyleRunProperties styleRunProperties = headingStyle.AppendChild(new StyleRunProperties());

            styleRunProperties.AppendChild(new RunFonts
            {
                AsciiTheme = ThemeFontValues.MajorHighAnsi,
                HighAnsiTheme = ThemeFontValues.MajorHighAnsi,
                EastAsiaTheme = ThemeFontValues.MajorEastAsia,
                ComplexScriptTheme = ThemeFontValues.MajorBidi
            });

            styleRunProperties.AppendChild(new FontSize { Val = "32" });
            styleRunProperties.AppendChild(new FontSizeComplexScript { Val = "32" });
        }

        StyleDefinitionsPart BuildStyles()
        {
            var stylePart = newDoc.AddNewPart<StyleDefinitionsPart>();
            var root = new Styles();
            root.Save(stylePart);
            return stylePart;
        }

        T EnsureExists<T>(OpenXmlElement input) where T: OpenXmlElement, new()
        {
            var tmp = input.Descendants<T>().FirstOrDefault();
            if (tmp != null)
            {
                return tmp;
            }
            return input.AppendChild(new T());
        }
        
        void SetColumnCount(Document doc, int columnCount, bool equalWidth = true)
        {
            var columns = doc.AppendChild(new Columns());
            columns.ColumnCount = (Int16Value)columnCount;
            columns.EqualWidth = equalWidth;
        }
        Paragraph CreateParagraph(DocxConfig configDocx, StyleConfig styleConfig, int indentation = 0, int spaceAfter = 0, string paragraphStyleId = null)
        {
            var paragraph = new Paragraph();
            var paragraphProperties = paragraph.AppendChild(new ParagraphProperties());

            var bidi = paragraphProperties.AppendChild(new BiDi());
            bidi.Val = new OnOffValue(configDocx.rightToLeft);

            var spacing = paragraphProperties.AppendChild(new SpacingBetweenLines());
            spacing.Line = (configDocx.lineSpacing * 240).ToString();
            spacing.After = (spaceAfter != 0 ? spaceAfter : 200).ToString();

            int marginLeft = configDocx.marginLeft * 567;
            int marginRight = configDocx.marginRight * 567;
            if (indentation != 0)
            {
                if (configDocx.rightToLeft)
                {
                    marginRight += indentation;
                }
                else
                {
                    marginLeft += indentation;
                }
            }
            var indentationElement = paragraphProperties.AppendChild(new Indentation());
            indentationElement.Left = marginLeft.ToString();
            indentationElement.Right = marginRight.ToString();

            if (paragraphStyleId != null)
            {
                paragraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = paragraphStyleId };
            }
            paragraphProperties.AppendChild(new Justification() { Val = (JustificationValues)configDocx.textAlign });

            return paragraph;
        }
        Run CreateRun(StyleConfig styleConfig, bool isSuperScript = false, int? runSpacing = null)
        {
            var run = new Run();
            var runProperties = run.AppendChild(new RunProperties());
            if (styleConfig.isBold)
            {
                var bold = runProperties.AppendChild(new Bold());
                bold.Val = styleConfig.isBold;
            }
            if (styleConfig.isItalics)
            {
                var italic = runProperties.AppendChild(new Italic());
                italic.Val = styleConfig.isItalics;
            }
            if (runSpacing.HasValue)
            {
               runProperties.AppendChild(new Spacing()).Val = runSpacing;
            }
            var fontSize = runProperties.AppendChild(new FontSize());
            fontSize.Val = (styleConfig.fontSize *2).ToString();
            if (isSuperScript)
            {
                var verticalAlignment = runProperties.AppendChild(new VerticalTextAlignment());
                verticalAlignment.Val = VerticalPositionValues.Superscript;
            }
            return run;
        }
        Run CreateBreakRun(BreakValues type)
        {
            var run = new Run();
            var breakElement = run.AppendChild(new Break());
            breakElement.Type = type;
            return run;
        }


        private void RenderMarker(Marker input, StyleConfig styles, Paragraph parentParagraph = null)
        {
            // Keep track of the previous marker
            previousMarker = thisMarker;
            thisMarker = input;

            StyleConfig markerStyle = (StyleConfig)styles.Clone();
            switch (input)
            {
                case PMarker _:
                    Paragraph paragraph = parentParagraph;
                    // If the previous marker was a chapter marker, don't create a new paragraph.
                    if (!(previousMarker is CMarker _))
                    {
                        paragraph = AppendToBody(CreateParagraph(configDocx, styles));
                    }

                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, paragraph);
                        // TODO: This is horribly innefficient needs to be rethought
                        if (lastAppendedParagraph != paragraph)
                        {
                            if (!lastAppendedParagraph.Descendants<Run>().Any())
                            {
                                paragraph = lastAppendedParagraph;
                                continue;
                            }
                            paragraph = AppendToBody(CreateParagraph(configDocx, styles));
                        }
                    }
                    break;
                case CLMarker clMarker:
                    if (beforeFirstChapter)
                    {
                        // A CL before the first chapter means that we should use
                        // this string instead of the word "Chapter".
                        chapterLabel = clMarker.Label;
                    }
                    break;
                case CMarker cMarker:

                    if (beforeFirstChapter)
                    {
                        // We found the first chapter, so set the flag to false.
                        beforeFirstChapter = false;
                    }
                    else
                    {
                        if (configDocx.separateChapters)
                        {
                            var newParagraph = AppendToBody(CreateParagraph(configDocx, styles));
                            var run =newParagraph.AppendChild(new Run());
                            var breakType = run.AppendChild(new Break());
                            breakType.Type = BreakValues.Page;
                        }
                    }

                    createBookHeaders(previousBookHeader);

                    var newChapter = AppendToBody(CreateParagraph(configDocx,styles));
                    var chapterMarker = newChapter.AppendChild(new Run());
                    string simpleNumber = cMarker.Number.ToString();
                    if (cMarker.CustomChapterLabel != simpleNumber)
                    {
                        // Use the custom label for this section, e.g. "Psalm One" instead of "Chapter 1"
                        currentChapterLabel = cMarker.CustomChapterLabel;
                    }
                    else
                    {
                        // Use the default chapter text for this section, e.g. "Chapter 1"
                        currentChapterLabel = chapterLabel + " " + simpleNumber;
                    }
                    var runProperties = chapterMarker.AppendChild(new RunProperties());
                    var fontSize = runProperties.AppendChild(new FontSize());
                    fontSize.Val = ((int)(configDocx.fontSize * 3)).ToString();
                    chapterMarker.AppendChild(new Text(currentChapterLabel));

                    var chapterVerses = AppendToBody(CreateParagraph(configDocx, markerStyle));
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, chapterVerses);
                    }

                    RenderCrossReferences(markerStyle);

                    break;
                case VMarker vMarker:

                    // If there is no parent paragraph, then we're maybe
                    // missing a chapter marker prior to this verse.  Let's
                    // create a stub parent paragraph so we can keep rendering.
                    if (parentParagraph == null)
                    {
                        parentParagraph = AppendToBody(CreateParagraph(configDocx, markerStyle));
                    }

                    if (configDocx.separateVerses)
                    {
                        var newLine = parentParagraph.AppendChild(new Run());
                        var breakElement = newLine.AppendChild(new Break());
                        breakElement.Type = BreakValues.TextWrapping;
                    }

                    markerStyle.fontSize = configDocx.fontSize;
                    var verseMarker = parentParagraph.AppendChild(CreateRun(markerStyle, isSuperScript: true));
                    verseMarker.AppendChild(new Text(vMarker.VerseCharacter));
                    verseMarker.AppendChild(new Text("\u00A0"));

                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, parentParagraph);
                    }
                    List<Text> textBlocks = parentParagraph.Descendants<Text>().ToList();

                    if (textBlocks.Count() != 0 && !textBlocks[textBlocks.Count()-1].Text.EndsWith(" "))
                    {
                        AppendSpace(parentParagraph);
                    }
                    break;
                case QMarker qMarker:

                    markerStyle.fontSize = configDocx.fontSize;
                    var poetryParagraph = CreateParagraph(configDocx, markerStyle, spaceAfter: 200, indentation: qMarker.Depth * 500);

                    if (!parentParagraph.Descendants<Run>().Any() && parentParagraph.Parent != null)
                    {
                            body.RemoveChild(parentParagraph);
                            AppendToBody(poetryParagraph);
                    }
                    else
                    {
                        AppendToBody(poetryParagraph);
                    }

                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, poetryParagraph);
                    }
                    break;
                case MMarker mMarker:
                    break;
                case TextBlock textBlock:
                    markerStyle.fontSize = configDocx.fontSize;
                    var blockText = parentParagraph.AppendChild(CreateRun(markerStyle));
                    blockText.AppendChild(new Text(textBlock.Text)).Space = SpaceProcessingModeValues.Preserve;
                    break;
                case BDMarker bdMarker:
                    markerStyle.isBold = true;
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, parentParagraph);
                    }
                    break;
                case HMarker hMarker:
                    // Add section header for previous book, if any
                    // (section page headers are set at the final paragraph of the section)
                    if (previousBookHeader != null)
                    {
                        // Create new section and page header
                        createBookHeaders(previousBookHeader);
                        // Print page break
                        var sectionParagraph = AppendToBody(CreateParagraph(configDocx,markerStyle));
                        sectionParagraph.AppendChild(CreateBreakRun(BreakValues.Page));
                    }
                    previousBookHeader = hMarker.HeaderText;

                    // Write body header text
                    markerStyle.fontSize = configDocx.fontSize;
                    var newHeader = AppendToBody(CreateParagraph(configDocx, markerStyle, paragraphStyleId:"BookHeading", spaceAfter:200));
                    var headerTitle = newHeader.AppendChild(CreateRun(markerStyle));

                    headerTitle.AppendChild(new Text(hMarker.HeaderText));

                    break;
                case FMarker fMarker:
                    var footnote = footnotes.AppendChild(new Footnote());
                    footnote.Id = nextFootnoteNum;
                    footnote.Type = FootnoteEndnoteValues.Normal;
                    StyleConfig footnoteMarkerStyle = (StyleConfig)styles.Clone();
                    footnoteMarkerStyle.fontSize = 12;
                    var footnoteParagraph = footnote.AppendChild(new Paragraph());
                    var footnoteRun = footnoteParagraph.AppendChild(CreateRun(footnoteMarkerStyle));
                    footnoteRun.AppendChild(new Text($"F{nextFootnoteNum} ")).Space = SpaceProcessingModeValues.Preserve;

                    foreach (Marker marker in fMarker.Contents)
                    {
                        RenderMarker(marker, footnoteMarkerStyle, footnoteParagraph);
                    }

                    var referenceRun = parentParagraph.AppendChild(new Run());
                    var referenceRunProperties = referenceRun.AppendChild(new RunProperties());
                    referenceRunProperties.AppendChild(new Underline());
                    referenceRunProperties.AppendChild(new VerticalTextAlignment()).Val = VerticalPositionValues.Superscript;
                    referenceRun.AppendChild(new Text($"F"));

                    var footnoteReference = new FootnoteReference();
                    footnoteReference.Id = nextFootnoteNum;
                    referenceRun.AppendChild(footnoteReference);

                    nextFootnoteNum++;
                    break;
                case FPMarker fPMarker:
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, parentParagraph);
                    }
                    break;
                case FTMarker fTMarker:

                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, parentParagraph);
                    }
                    break;
                case FRMarker fRMarker:
                    markerStyle.isBold = true;
                    parentParagraph.AppendChild(CreateRun(markerStyle));
                    parentParagraph.AppendChild(new Text(fRMarker.VerseReference));
                    break;
                case FKMarker fKMarker:
                    var FootNoteKeyword = parentParagraph.AppendChild(CreateRun(markerStyle));
                    FootNoteKeyword.AppendChild(new Text($" {fKMarker.FootNoteKeyword.ToUpper()}: "));
                    break;
                case FQMarker fQMarker:
                case FQAMarker fQAMarker:

                    markerStyle.isItalics = true;
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, parentParagraph);
                    }
                    break;
                // Cross References
                case XMarker xMarker:
                    string crossId;
                    switch (xMarker.CrossRefCaller)
                    {
                        case "-":
                            crossId = "";
                            break;
                        case "+":
                            crossId = $"{CrossRefMarkers.Count + 1}";
                            break;
                        default:
                            crossId = xMarker.CrossRefCaller;
                            break;
                    }
                    var crossRefMarker = parentParagraph.AppendChild(CreateRun(markerStyle, isSuperScript: true));
                    crossRefMarker.AppendChild(new Text(crossId));

                    CrossRefMarkers[crossId] = xMarker;
                    break;
                case XOMarker xOMarker:
                    markerStyle.isBold = true;
                    var CrossVerseReference = parentParagraph.AppendChild(CreateRun(markerStyle));
                    CrossVerseReference.AppendChild(new Text($" {xOMarker.OriginRef} "));
                    break;
                case XTMarker xTMarker:
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, parentParagraph);
                    }
                    break;
                case XQMarker xQMarker:
                    markerStyle.isItalics = true;
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, parentParagraph);
                    }
                    break;
                // Table Markers
                case TableBlock table:
                    var tableContainer = new Table();
                    tableContainer.AppendChild(new TableProperties(new TableBorders(
                        new TopBorder
                        {
                            Val = BorderValues.None
                        },
                        new BottomBorder
                        {
                            Val = BorderValues.None
                        },
                        new LeftBorder
                        {
                            Val = BorderValues.None
                        },
                        new RightBorder
                        {
                            Val = BorderValues.None
                        },
                        new InsideHorizontalBorder
                        {
                            Val = BorderValues.None
                        },
                        new InsideVerticalBorder
                        {
                            Val= BorderValues.None
                        }
                        )));

                    foreach (Marker marker in input.Contents)
                    {
                        getRenderedRows(marker, markerStyle, tableContainer);
                    }
                    break;
                case BMarker bMarker:
                    var newLineBreak = parentParagraph.AppendChild(CreateRun(markerStyle));
                    var breakObject = newLineBreak.AppendChild(new Break());
                    breakObject.Type = BreakValues.TextWrapping;
                    break;
                case IDMarker _:
                    // This is the start of a new book.
                    beforeFirstChapter = true;
                    chapterLabel = chapterLabelDefault;
                    currentChapterLabel = "";
                    break;
                case IPMarker _:
                    Paragraph introParagraph = parentParagraph;
                    // If the previous marker was a chapter marker, don't create a new paragraph.
                    if (!(previousMarker is CMarker _))
                    {
                        var newParagraph = AppendToBody(CreateParagraph(configDocx, markerStyle, spaceAfter:200));
                        introParagraph = newParagraph;
                    }

                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, introParagraph);
                    }
                    break;
                case XEndMarker _:
                case FEndMarker _:
                case IDEMarker _:
                case VPMarker _:
                case VPEndMarker _:
                    break;
                default:
                    UnrenderableMarkers.Add(input.Identifier);
                    break;
            }
        }

        /// <summary>
        /// Appends a text run containing a single space.  The run is
        /// space-preserved so that the space will be visible.
        /// </summary>
        /// <param name="paragraph">The parent paragraph to add the run to.</param>
        private void AppendSpace(Paragraph paragraph)
        {
            var run = paragraph.AppendChild(new Run());
            run.AppendChild(new Text(" ")).Space = SpaceProcessingModeValues.Preserve;
        }

        private void RenderCrossReferences(StyleConfig config)
        {

            if (CrossRefMarkers.Count > 0)
            {
                StyleConfig markerStyle = (StyleConfig)config.Clone();
                markerStyle.fontSize = 12;

                foreach (var crossRefKVP in CrossRefMarkers)
                {
                    var renderCrossRef = AppendToBody(new Paragraph(new ParagraphProperties(
                        new ParagraphBorders(new TopBorder() { Val = BorderValues.Single})),
                        new BiDi() { Val = new OnOffValue(configDocx.rightToLeft)}));
                    var crossRefMarker = renderCrossRef.AppendChild(CreateRun(markerStyle, isSuperScript:true));
                    crossRefMarker.AppendChild(new Text(crossRefKVP.Key));

                    foreach (Marker input in crossRefKVP.Value.Contents)
                    {
                        RenderMarker(input, markerStyle, renderCrossRef);
                    }

                }
                CrossRefMarkers.Clear();
            }
        }
        public void setStartPageNumber()
        {
            var properties = AppendToBody(new SectionProperties());
            var numberType = properties.AppendChild(new PageNumberType());
            numberType.Format = NumberFormatValues.Decimal;
            numberType.Start = 1;
        }

        /// <summary>
        /// Creates a new section with the given page header.  Must be
        /// called *after* the final paragraph of the section.  In DOCX, a
        /// section definition is a child of the final paragraph of the
        /// section, except for the final section of the document, which
        /// is a direct child of the body.
        /// </summary>
        /// <param name="bookname"> The name of the book to display, usually from the \h marker </param>
        public void createBookHeaders(string bookname)
        {
            // Create page heading content for book
            var header = new Header();
            var headerParagraph = header.AppendChild(new Paragraph());
            var ppr = headerParagraph.AppendChild(new ParagraphProperties());
            var align = ppr.AppendChild(new Justification());
            align.Val = JustificationValues.Center;
            var run = headerParagraph.AppendChild(new Run());

            // Show page numbers if requested
            if (configDocx.showPageNumbers)
            {
                // Page number
                run.AppendChild(new FieldChar()).FieldCharType = FieldCharValues.Begin;
                run.AppendChild(new FieldCode(" PAGE "));
                run.AppendChild(new FieldChar()).FieldCharType = FieldCharValues.Separate;

                run.AppendChild(new FieldCode("1"));
                run.AppendChild(new FieldChar()).FieldCharType = FieldCharValues.End;
                run.AppendChild(new Text(" - "));
            }

            // Book name
            run.AppendChild(new Text(bookname == null ? "" : bookname));
            // Chapter name
            if (currentChapterLabel.Length > 0)
            {
                run.AppendChild(new Text("  -  "));
                run.AppendChild(new Text(currentChapterLabel));
            }

            var headerId = $"rId{pageHeaderCount}";

            var headerPart = newDoc.MainDocumentPart.AddNewPart<HeaderPart>(headerId);
            headerPart.Header = header;


            // Create page header
            var sectionProperties = AppendToBody(new Paragraph()).AppendChild(new ParagraphProperties()).AppendChild(new SectionProperties());
            var headerReference = new HeaderReference();
            headerReference.Id = headerId;
            headerReference.Type = HeaderFooterValues.Default;
            sectionProperties.Append(headerReference);
            var sectionType = sectionProperties.AppendChild(new SectionType());
            sectionType.Val = SectionMarkValues.Continuous;

            var pageNumberType = sectionProperties.AppendChild(new PageNumberType());
            pageNumberType.Format = NumberFormatValues.Decimal;
            pageNumberType.ChapterSeparator = ChapterSeparatorValues.Hyphen;

            sectionProperties.AppendChild(new Columns()).ColumnCount = (Int16Value)configDocx.columnCount;

            // Increment page header count so each one gets a unique ID
            pageHeaderCount++;
        }

        /// <summary>
        /// Creates an empty header for front pages.
        /// The returned paragraph should be inserted in front of document
        /// </summary>
        /// <example>xwpfDoc.Document.body.Items.Insert(1, CreateFrontHeader());</example>
        /// <returns>CT_P paragraph that contains a blank header</returns>
        private Paragraph CreateFrontHeader()
        {


            var header = new Header();
            header.AppendChild(new Paragraph(new ParagraphProperties()));

            var headerId = $"rId{pageHeaderCount}";

            var headerPart = newDoc.MainDocumentPart.AddNewPart<HeaderPart>(headerId);
            headerPart.Header = header;

            // Create new section and set its header
            var p = new Paragraph();
            var paragraphProperties = p.AppendChild(new ParagraphProperties());
            var newSection = paragraphProperties.AppendChild(new SectionProperties());
            newSection.AppendChild(new SectionType() { Val = SectionMarkValues.Continuous});
            newSection.AppendChild(new HeaderReference() { Type = HeaderFooterValues.Default, Id = headerId });


            // Increment page header count so each one gets a unique ID
            pageHeaderCount++;

            return p;
        }

        public void getRenderedRows(Marker input, StyleConfig config, Table parentTable)
        {
            var tableRowContainer = parentTable.AppendChild(new TableRow());
            foreach (Marker marker in input.Contents)
            {
                getRenderedCell(marker, config, tableRowContainer);
            }
        }
        public void getRenderedCell(Marker input, StyleConfig config, TableRow parentRow)
        {
            StyleConfig markerStyle = (StyleConfig)config.Clone();
            var tableCellContainer = parentRow.AppendChild(new TableCell());
            Paragraph cellContents;
            switch (input)
            {
                case THMarker tHMarker:
                    markerStyle.isBold = true;
                    cellContents = tableCellContainer.AppendChild(CreateParagraph(configDocx, markerStyle));
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, cellContents);
                    }
                    break;
                case THRMarker tHRMarker:
                    markerStyle.isAlignRight = true;
                    markerStyle.isBold = true;
                    cellContents = tableCellContainer.AppendChild(CreateParagraph(configDocx, markerStyle));
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, cellContents);
                    }
                    break;
                case TCMarker tCMarker:
                    cellContents = tableCellContainer.AppendChild(CreateParagraph(configDocx, markerStyle));
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, cellContents);
                    }
                    break;
                case TCRMarker tCRMarker:
                    markerStyle.isAlignRight = true;
                    cellContents = tableCellContainer.AppendChild(CreateParagraph(configDocx, markerStyle));
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, cellContents);
                    }
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// Renders a Table of Contents (TOC) in front of the document.
        /// 
        /// Please set the paragraphs style to "Heading{#}". 
        /// Otherwise, it renders an empty TOC.
        /// </summary>
        private void RenderTOC()
        {

            var sdtBlock = body.AppendChild(new SdtBlock());
            sdtBlock.AppendChild(new SdtProperties(
                new SdtContentDocPartObject(
                    new DocPartGallery() { Val = "Table of Contents"},
                    new DocPartUnique() { Val = true}
                    )
                ));
            var sdtContent = sdtBlock.AppendChild(new SdtContentBlock());
            sdtContent.AppendChild(new Paragraph(
                new Run( new FieldChar() { FieldCharType = FieldCharValues.Begin}),
                new Run(new FieldCode() { Space = SpaceProcessingModeValues.Preserve, Text =" TOC \\f \\o \"1-9\" \\h" }),
                //new Run(new FieldCode() { Space = SpaceProcessingModeValues.Preserve, Text =" TOC " }),
                new Run( new FieldChar() { FieldCharType = FieldCharValues.Separate}),
                new Run( new FieldChar() { FieldCharType = FieldCharValues.End})
                ));
            AppendToBody(CreateFrontHeader());
        }

        private void RenderFrontMatter(USFMDocument frontMatter)
        {
            // reset default format before rendering front matters
            DocxConfig currentConfig = configDocx;
            configDocx = new DocxConfig(); 

            foreach (var marker in frontMatter.Contents)
            {
                RenderMarker(marker, new StyleConfig());
            }

            // revert to user config format
            configDocx = currentConfig;

            AppendToBody(CreateFrontHeader());
            AppendToBody(new Paragraph(CreateBreakRun(BreakValues.Page)));
        }
        private Paragraph GetLastParagraph()
        {
            for(var i = body.ChildElements.Count - 1; i >= 0; i--)
            {
                if (body.ChildElements[i] is Paragraph)
                {
                    return (Paragraph)body.ChildElements[i];
                }

            }
            return null;
        }
    }
}
