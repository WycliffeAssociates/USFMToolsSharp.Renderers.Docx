using System.Collections.Generic;
using System.Text;
using USFMToolsSharp.Models.Markers;
using NPOI.XWPF.UserModel;
using NPOI.XWPF.Model;
using NPOI.OpenXmlFormats.Wordprocessing;
using USFMToolsSharp.Renderers.Docx.Extensions;
using System;
using USFMToolsSharp.Renderers.Docx.Utils;

namespace USFMToolsSharp.Renderers.Docx
{

    public class DocxRenderer
    {
        public List<string> UnrenderableMarkers;
        public Dictionary<string, Marker> CrossRefMarkers;
        private Dictionary<string, string> TOCEntries;
        private DocxConfig configDocx;
        private XWPFDocument newDoc;
        private int pageHeaderCount = 1;
        private string previousBookHeader = null;
        private const string chapterLabelDefault = "Chapter";
        private string chapterLabel = chapterLabelDefault;
        private string currentChapterLabel = "";
        private bool beforeFirstChapter = true;
        private int nextFootnoteNum = 1;
        private Marker thisMarker = null;
        private Marker previousMarker = null;

        public DocxRenderer()
        {
            configDocx = new DocxConfig();
        }
        public DocxRenderer(DocxConfig config)
        {
            configDocx = config;
        }

        public XWPFDocument Render(USFMDocument input)
        {
            UnrenderableMarkers = new List<string>();
            CrossRefMarkers = new Dictionary<string, Marker>();
            TOCEntries = new Dictionary<string, string>();
            newDoc = new XWPFDocument();
            newDoc.CreateFootnotes();

            setStartPageNumber();

            newDoc.ColumnCount = configDocx.columnCount;

            foreach (Marker marker in input.Contents)
            {
                RenderMarker(marker, new StyleConfig());
            }

            // Add section header for final book
            if (previousBookHeader != null)
            {
                createBookHeaders(previousBookHeader);
            }

            // Make final document section continuous so that it doesn't
            // create an extra page at the end.  Final section is unique:
            // it's a direct child of the document, not a child of the last
            // paragraph.
            CT_SectPr finalSection = new CT_SectPr();
            finalSection.type = new CT_SectType();
            finalSection.type.val = ST_SectionMark.continuous;
            newDoc.Document.body.sectPr = finalSection;
            finalSection.cols.num = configDocx.columnCount.ToString();
            CT_PageNumber pageNum = new CT_PageNumber
            {
                fmt = ST_NumberFormat.@decimal
            };
            finalSection.pgNumType = pageNum;

            RenderTOC();
            return newDoc;

        }
        private void RenderMarker(Marker input, StyleConfig styles, XWPFParagraph parentParagraph = null)
        {
            // Keep track of the previous marker
            previousMarker = thisMarker;
            thisMarker = input;

            StyleConfig markerStyle = (StyleConfig)styles.Clone();
            switch (input)
            {
                case PMarker _:

                    XWPFParagraph paragraph = parentParagraph;
                    // If the previous marker was a chapter marker, don't create a new paragraph.
                    if (!(previousMarker is CMarker _))
                    {
                        XWPFParagraph newParagraph = newDoc.CreateParagraph(markerStyle);
                        newParagraph.SetBidi(configDocx.rightToLeft);
                        newParagraph.Alignment = configDocx.textAlign;
                        newParagraph.SpacingBetween = configDocx.lineSpacing;
                        newParagraph.SpacingAfter = 200;
                        paragraph = newParagraph;
                    }

                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, paragraph);
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
                            // Add page break between chapters.
                            newDoc.CreateParagraph().CreateRun().AddBreak(BreakType.PAGE);
                        }
                    }

                    createBookHeaders(previousBookHeader);

                    XWPFParagraph newChapter = newDoc.CreateParagraph(markerStyle);
                    newChapter.SetBidi(configDocx.rightToLeft);
                    newChapter.Alignment = configDocx.textAlign;
                    newChapter.SpacingBetween = configDocx.lineSpacing;
                    newChapter.SpacingBefore = 200;
                    newChapter.SpacingAfter = 200;
                    XWPFRun chapterMarker = newChapter.CreateRun(markerStyle);
                    setRTL(chapterMarker);
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
                    chapterMarker.SetText(currentChapterLabel);
                    chapterMarker.FontSize = 20;

                    XWPFParagraph chapterVerses = newDoc.CreateParagraph(markerStyle);
                    chapterVerses.SetBidi(configDocx.rightToLeft);
                    chapterVerses.Alignment = configDocx.textAlign;
                    chapterVerses.SpacingBetween = configDocx.lineSpacing;
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
                        parentParagraph = newDoc.CreateParagraph(markerStyle);
                        parentParagraph.SetBidi(configDocx.rightToLeft);
                        parentParagraph.Alignment = configDocx.textAlign;
                        parentParagraph.SpacingBetween = configDocx.lineSpacing;
                        parentParagraph.SpacingAfter = 200;
                    }

                    if (configDocx.separateVerses)
                    {
                        XWPFRun newLine = parentParagraph.CreateRun();
                        newLine.AddBreak(BreakType.TEXTWRAPPING);
                    }

                    XWPFRun verseMarker = parentParagraph.CreateRun(markerStyle);
                    setRTL(verseMarker);
                    verseMarker.SetText(vMarker.VerseCharacter);
                    verseMarker.Subscript = VerticalAlign.SUPERSCRIPT;
                    AppendSpace(parentParagraph);

                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, parentParagraph);
                    }
                    if (parentParagraph.Text.EndsWith(" ") == false)
                    {
                        AppendSpace(parentParagraph);
                    }
                    break;
                case QMarker qMarker:
                    XWPFParagraph poetryParagraph = newDoc.CreateParagraph(markerStyle);
                    poetryParagraph.SetBidi(configDocx.rightToLeft);
                    poetryParagraph.Alignment = configDocx.textAlign;
                    poetryParagraph.SpacingBetween = configDocx.lineSpacing;
                    poetryParagraph.IndentationLeft = qMarker.Depth * 500;
                    poetryParagraph.SpacingAfter = 200;

                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, poetryParagraph);
                    }
                    break;
                case MMarker mMarker:
                    break;
                case TextBlock textBlock:
                    XWPFRun blockText = parentParagraph.CreateRun(markerStyle);
                    setRTL(blockText);
                    blockText.SetText(textBlock.Text);
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
                        XWPFParagraph sectionParagraph = newDoc.CreateParagraph();
                        sectionParagraph.SetBidi(configDocx.rightToLeft);
                        sectionParagraph.Alignment = configDocx.textAlign;
                        sectionParagraph.CreateRun().AddBreak(BreakType.PAGE);
                    }
                    previousBookHeader = hMarker.HeaderText;

                    // Write body header text
                    markerStyle.fontSize = 24;
                    XWPFParagraph newHeader = newDoc.CreateParagraph(markerStyle);
                    newHeader.SetBidi(configDocx.rightToLeft);
                    newHeader.SpacingAfter = 200;
                    XWPFRun headerTitle = newHeader.CreateRun(markerStyle);
                    setRTL(headerTitle);
                    headerTitle.SetText(hMarker.HeaderText);

                    break;
                case FMarker fMarker:
                    string footnoteId;
                    footnoteId = nextFootnoteNum.ToString();
                    nextFootnoteNum++;

                    CT_FtnEdn footnote = new CT_FtnEdn();
                    footnote.id = footnoteId;
                    footnote.type = ST_FtnEdn.normal;
                    StyleConfig footnoteMarkerStyle = (StyleConfig)styles.Clone();
                    footnoteMarkerStyle.fontSize = 12;
                    CT_P footnoteParagraph = footnote.AddNewP();
                    XWPFParagraph xFootnoteParagraph = new XWPFParagraph(footnoteParagraph, parentParagraph.Body);
                    xFootnoteParagraph.SetBidi(configDocx.rightToLeft);
                    footnoteParagraph.AddNewR().AddNewT().Value = "F" + footnoteId.ToString() + " ";
                    foreach (Marker marker in fMarker.Contents)
                    {
                        RenderMarker(marker, footnoteMarkerStyle, xFootnoteParagraph);
                    }
                    parentParagraph.Document.AddFootnote(footnote);

                    XWPFRun footnoteReferenceRun = parentParagraph.CreateRun();
                    setRTL(footnoteReferenceRun);
                    CT_RPr rpr = footnoteReferenceRun.GetCTR().AddNewRPr();
                    rpr.rStyle = new CT_String();
                    rpr.rStyle.val = "FootnoteReference";
                    CT_FtnEdnRef footnoteReference = new CT_FtnEdnRef();
                    footnoteReference.id = footnoteId;
                    footnoteReference.isEndnote = false;
                    footnoteReferenceRun.SetUnderline(UnderlinePatterns.Single);
                    footnoteReferenceRun.AppendText("F");
                    footnoteReferenceRun.GetCTR().Items.Add(footnoteReference);
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
                    XWPFRun VerseReference = parentParagraph.CreateRun(markerStyle);
                    setRTL(VerseReference);
                    VerseReference.SetText(fRMarker.VerseReference);
                    break;
                case FKMarker fKMarker:
                    XWPFRun FootNoteKeyword = parentParagraph.CreateRun(markerStyle);
                    setRTL(FootNoteKeyword);
                    FootNoteKeyword.SetText($" {fKMarker.FootNoteKeyword.ToUpper()}: ");
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
                    XWPFRun crossRefMarker = parentParagraph.CreateRun(markerStyle);
                    setRTL(crossRefMarker);
                    crossRefMarker.SetText(crossId);
                    crossRefMarker.Subscript = VerticalAlign.SUPERSCRIPT;

                    CrossRefMarkers[crossId] = xMarker;
                    break;
                case XOMarker xOMarker:
                    markerStyle.isBold = true;
                    XWPFRun CrossVerseReference = parentParagraph.CreateRun(markerStyle);
                    setRTL(CrossVerseReference);
                    CrossVerseReference.SetText($" {xOMarker.OriginRef} ");
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
                    XWPFTable tableContainer = newDoc.CreateTable();

                    // Clear Borders
                    tableContainer.SetBottomBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "#FFFFFFF");
                    tableContainer.SetLeftBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "#FFFFFFF");
                    tableContainer.SetRightBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "#FFFFFFF");
                    tableContainer.SetTopBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "#FFFFFFF");
                    // Clear Inside Borders
                    tableContainer.SetInsideHBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "#FFFFFFF");
                    tableContainer.SetInsideVBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "#FFFFFFF");

                    foreach (Marker marker in input.Contents)
                    {
                        getRenderedRows(marker, markerStyle, tableContainer);
                    }
                    break;
                case BMarker bMarker:
                    XWPFRun newLineBreak = parentParagraph.CreateRun();
                    setRTL(newLineBreak);
                    newLineBreak.AddBreak(BreakType.TEXTWRAPPING);
                    break;
                case IDMarker _:
                    // This is the start of a new book.
                    beforeFirstChapter = true;
                    chapterLabel = chapterLabelDefault;
                    currentChapterLabel = "";
                    break;
                case IPMarker _:
                    XWPFParagraph introParagraph = parentParagraph;
                    // If the previous marker was a chapter marker, don't create a new paragraph.
                    if (!(previousMarker is CMarker _))
                    {
                        XWPFParagraph newParagraph = newDoc.CreateParagraph(markerStyle);
                        newParagraph.SetBidi(configDocx.rightToLeft);
                        newParagraph.Alignment = configDocx.textAlign;
                        newParagraph.SpacingBetween = configDocx.lineSpacing;
                        newParagraph.SpacingAfter = 200;
                        introParagraph = newParagraph;
                    }

                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, introParagraph);
                    }
                    break;
                case TOC2Marker toc2Marker:
                    string text = toc2Marker.ShortTableOfContentsText;
                    string bookMarkRef = AddTOCBookMark(text);
                    TOCEntries.Add(text, bookMarkRef);
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
        private void AppendSpace(XWPFParagraph paragraph)
        {
            XWPFRun run = paragraph.CreateRun();
            setRTL(run);
            CT_R ctr = run.GetCTR();
            CT_Text text = ctr.AddNewT();
            text.Value = " ";
            text.space = "preserve";
        }

        private void setRTL(XWPFRun run)
        {
            if (configDocx.rightToLeft)
            {
                CT_RPr rpr = run.GetCTR().AddNewRPr();
                rpr.rtl = new CT_OnOff();
                rpr.rtl.val = configDocx.rightToLeft;
            }
        }

        private void RenderCrossReferences(StyleConfig config)
        {

            if (CrossRefMarkers.Count > 0)
            {
                XWPFParagraph renderCrossRefStart = newDoc.CreateParagraph();
                renderCrossRefStart.BorderTop = Borders.Single;

                StyleConfig markerStyle = (StyleConfig)config.Clone();
                markerStyle.fontSize = 12;

                foreach (KeyValuePair<string, Marker> crossRefKVP in CrossRefMarkers)
                {
                    XWPFParagraph renderCrossRef = newDoc.CreateParagraph();
                    renderCrossRef.SetBidi(configDocx.rightToLeft);
                    XWPFRun crossRefMarker = renderCrossRef.CreateRun(markerStyle);
                    setRTL(crossRefMarker);
                    crossRefMarker.SetText(crossRefKVP.Key);
                    crossRefMarker.Subscript = VerticalAlign.SUPERSCRIPT;

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
            newDoc.Document.body.sectPr.pgNumType.fmt = ST_NumberFormat.@decimal;
            newDoc.Document.body.sectPr.pgNumType.start = "1";
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
            CT_Hdr header = new CT_Hdr();
            CT_P headerParagraph = header.AddNewP();
            CT_PPr ppr = headerParagraph.AddNewPPr();
            CT_Jc align = ppr.AddNewJc();
            align.val = ST_Jc.center;
            CT_R run = headerParagraph.AddNewR();

            // Show page numbers if requested
            if (configDocx.showPageNumbers)
            {
                // Page number
                run.AddNewFldChar().fldCharType = ST_FldCharType.begin;
                run.AddNewInstrText().Value = " PAGE ";
                run.AddNewFldChar().fldCharType = ST_FldCharType.separate;
                run.AddNewInstrText().Value = "1";
                run.AddNewFldChar().fldCharType = ST_FldCharType.end;
                run.AddNewT().Value = "  -  ";
            }

            // Book name
            run.AddNewT().Value = bookname;
            // Chapter name
            if (currentChapterLabel.Length > 0)
            {
                run.AddNewT().Value = "  -  ";
                run.AddNewT().Value = currentChapterLabel;
            }


            // Create page header
            XWPFHeader documentHeader = (XWPFHeader)newDoc.CreateRelationship(XWPFRelation.HEADER, XWPFFactory.GetInstance(), pageHeaderCount);
            documentHeader.SetHeaderFooter(header);

            // Create new section and set its header
            CT_SectPr newSection = newDoc.Document.body.AddNewP().AddNewPPr().createSectPr();
            newSection.type = new CT_SectType();
            newSection.type.val = ST_SectionMark.continuous;
            CT_HdrFtrRef headerRef = newSection.AddNewHeaderReference();
            headerRef.type = ST_HdrFtr.@default;
            headerRef.id = documentHeader.GetPackageRelationship().Id;

            // Set number of columns
            newSection.cols.num = configDocx.columnCount.ToString();

            // Set page numbers
            CT_PageNumber pageNum = new CT_PageNumber
            {
                fmt = ST_NumberFormat.@decimal
            };
            newSection.pgNumType = pageNum;

            // Increment page header count so each one gets a unique ID
            pageHeaderCount++;
        }

        /// <summary>
        /// Creates an empty header for front pages. Append the returned
        /// paragraph to the document body at the end of the front page,
        /// after the page break paragraph.
        /// 
        /// e.g. newDoc.Document.body.Items.Insert(1, frontHeader);
        /// where Items[0] is a page break
        /// </summary>
        /// <returns>A CT_P paragraph that contains the header</returns>
        private CT_P CreateFrontHeader()
        {
            CT_Hdr header = new CT_Hdr();
            CT_P headerParagraph = header.AddNewP();
            headerParagraph.AddNewPPr();

            XWPFHeader documentHeader = (XWPFHeader)newDoc.CreateRelationship(XWPFRelation.HEADER, XWPFFactory.GetInstance(), pageHeaderCount);
            documentHeader.SetHeaderFooter(header);

            // Create new section and set its header
            CT_P p = new CT_P();
            CT_SectPr newSection = p.AddNewPPr().createSectPr();
            newSection.type = new CT_SectType();
            newSection.type.val = ST_SectionMark.continuous;
            CT_HdrFtrRef headerRef = newSection.AddNewHeaderReference();
            headerRef.type = ST_HdrFtr.@default;
            headerRef.id = documentHeader.GetPackageRelationship().Id;

            // Increment page header count so each one gets a unique ID
            pageHeaderCount++;

            return p;
        }

        public void getRenderedRows(Marker input, StyleConfig config, XWPFTable parentTable)
        {
            XWPFTableRow tableRowContainer = parentTable.CreateRow();
            foreach (Marker marker in input.Contents)
            {
                getRenderedCell(marker, config, tableRowContainer);
            }
        }
        public void getRenderedCell(Marker input, StyleConfig config, XWPFTableRow parentRow)
        {
            StyleConfig markerStyle = (StyleConfig)config.Clone();
            XWPFTableCell tableCellContainer = parentRow.CreateCell();
            XWPFParagraph cellContents;
            switch (input)
            {
                case THMarker tHMarker:
                    markerStyle.isBold = true;
                    cellContents = tableCellContainer.AddParagraph(markerStyle);
                    cellContents.SetBidi(configDocx.rightToLeft);
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, cellContents);
                    }
                    break;
                case THRMarker tHRMarker:
                    markerStyle.isAlignRight = true;
                    markerStyle.isBold = true;
                    cellContents = tableCellContainer.AddParagraph(markerStyle);
                    cellContents.SetBidi(configDocx.rightToLeft);
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, cellContents);
                    }
                    break;
                case TCMarker tCMarker:
                    cellContents = tableCellContainer.AddParagraph(markerStyle);
                    cellContents.SetBidi(configDocx.rightToLeft);
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, cellContents);
                    }
                    break;
                case TCRMarker tCRMarker:
                    markerStyle.isAlignRight = true;
                    cellContents = tableCellContainer.AddParagraph(markerStyle);
                    cellContents.SetBidi(configDocx.rightToLeft);
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
        /// Inserts a hidden Table of Contents bookmark to the document body.
        /// </summary>
        /// <param name="name">The identifier name for a bookmark</param>
        /// <returns>The bookmark reference value</returns>
        private string AddTOCBookMark(string name)
        {
            var bookmarkName = name.Replace(" ", ""); // remove spaces
            //Bookmark start
            CT_P para = newDoc.Document.body.AddNewP();
            CT_Bookmark bookmark = new CT_Bookmark();
            bookmark.name = string.Format(TOCBuilder.TOC_BOOKMARK, bookmarkName);
            string bookmarkId = TOCEntries.Count.ToString();
            bookmark.id = bookmarkId;
            para.Items.Add(bookmark);
            para.ItemsElementName.Add(ParagraphItemsChoiceType.bookmarkStart);
            CT_R run = para.AddNewR();
            run.AddNewRPr().vanish = new CT_OnOff() { val = true };
            var t = run.AddNewT();
            t.Value = "This is the bookmark of " + name;

            //Bookmark end
            bookmark = new CT_Bookmark();
            bookmark.id = bookmarkId;
            para.Items.Add(bookmark);
            para.ItemsElementName.Add(ParagraphItemsChoiceType.bookmarkEnd);

            var pPr = para.AddNewPPr();
            pPr.AddNewRPr().vanish = new CT_OnOff() { val = true };

            return bookmarkName;    
        }

        /// <summary>
        /// Renders a Table of Contents (TOC) in front of the document
        /// based on the bookmarks in the document body.
        /// 
        /// Only invoke this method after parsing the markers content.
        /// Otherwise, it renders an empty TOC.
        /// </summary>
        private void RenderTOC()
        {
            TOCBuilder tocBuilder = new TOCBuilder();

            foreach (var entry in TOCEntries)
            {
                tocBuilder.AddRow(1, entry.Key, entry.Value);
            }

            CT_SdtBlock toc = tocBuilder.Build();

            // add page break after TOC
            CT_P pBreak = new CT_P();
            pBreak.AddNewR().AddNewBr().type = ST_BrType.page;
            
            CT_P pHeader = CreateFrontHeader();

            newDoc.Document.body.Items.Insert(0, toc);
            newDoc.Document.body.Items.Insert(1, pBreak);
            newDoc.Document.body.Items.Insert(2, pHeader);

            newDoc.EnforceUpdateFields();
        }

    }
}
