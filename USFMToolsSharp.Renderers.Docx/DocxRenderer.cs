using System.Collections.Generic;
using System.Text;
using USFMToolsSharp.Models.Markers;
using NPOI.XWPF.UserModel;
using NPOI.XWPF.Model;
using NPOI.OpenXmlFormats.Wordprocessing;
using USFMToolsSharp.Renderers.Docx.Extensions;

namespace USFMToolsSharp.Renderers.Docx
{
    
    public class DocxRenderer
    {
        public List<string> UnrenderableMarkers;
        //public Dictionary<string,Marker> FootnoteMarkers;
        public Dictionary<string, Marker> CrossRefMarkers;
        private DocxConfig configDocx;
        private XWPFDocument newDoc;
        private int pageHeaderCount = 1;
        private string previousBookHeader = null;
        private const string chapterLabelDefault = "Chapter";
        private string chapterLabel = chapterLabelDefault;
        private bool beforeFirstChapter = true;
        private int nextFootnoteNum = 1;

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
            //FootnoteMarkers = new Dictionary<string, Marker>();
            CrossRefMarkers = new Dictionary<string, Marker>();
            newDoc = new XWPFDocument();
            newDoc.CreateFootnotes();

            setStartPageNumber();

            newDoc.ColumnCount = configDocx.columnCount;
            newDoc.TextDirection= configDocx.textDirection;

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

            return newDoc;

        }
        private void RenderMarker(Marker input, StyleConfig styles, XWPFParagraph parentParagraph = null)
        {
            StyleConfig markerStyle = (StyleConfig)styles.Clone();
            switch (input)
            {
                case PMarker _:
                    XWPFParagraph newParagraph = newDoc.CreateParagraph(markerStyle);

                    newParagraph.Alignment = configDocx.textAlign;
                    newParagraph.SpacingBetween = configDocx.lineSpacing;

                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, newParagraph);
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

                    XWPFParagraph newChapter = newDoc.CreateParagraph(markerStyle);
                    XWPFRun chapterMarker = newChapter.CreateRun(markerStyle);
                    string simpleNumber = cMarker.Number.ToString();
                    if (cMarker.CustomChapterLabel != simpleNumber)
                    {
                        // Use the custom label for this section, e.g. "Psalm One" instead of "Chapter 1"
                        chapterMarker.SetText(cMarker.CustomChapterLabel);
                    }
                    else
                    {
                        // Use the default chapter text for this section, e.g. "Chapter 1"
                        chapterMarker.SetText(chapterLabel + " " + simpleNumber);
                    }
                    chapterMarker.FontSize = 20;

                    XWPFParagraph chapterVerses = newDoc.CreateParagraph(markerStyle);
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle ,chapterVerses);
                    }

                    //RenderFootnotes(markerStyle);
                    RenderCrossReferences(markerStyle);

                    break;
                case VMarker vMarker:

                    if (configDocx.separateVerses)
                    {
                        XWPFRun newLine = parentParagraph.CreateRun();
                        newLine.AddBreak(BreakType.TEXTWRAPPING);
                    }
                    else 
                    { 
                        AppendSpace(parentParagraph);
                    }

                    XWPFRun verseMarker = parentParagraph.CreateRun(markerStyle);
                    verseMarker.SetText(vMarker.VerseCharacter);
                    verseMarker.Subscript = VerticalAlign.SUPERSCRIPT;

                    foreach (Marker marker in input.Contents)
                    {
                        AppendSpace(parentParagraph);
                        RenderMarker(marker, markerStyle, parentParagraph);
                    }
                    break;
                case QMarker qMarker:
                    XWPFParagraph poetryParagraph = newDoc.CreateParagraph(markerStyle);
                    poetryParagraph.IndentationLeft = qMarker.Depth;

                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker,markerStyle, poetryParagraph);
                    }
                    break;
                case MMarker mMarker:
                    break;
                case TextBlock textBlock:
                    XWPFRun blockText = parentParagraph.CreateRun(markerStyle);
                    blockText.SetText(textBlock.Text);
                    break;
                case BDMarker bdMarker:
                    markerStyle.isBold = true;
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker,markerStyle,parentParagraph);
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
                        newDoc.CreateParagraph().CreateRun().AddBreak(BreakType.PAGE);
                    }
                    previousBookHeader = hMarker.HeaderText;

                    // Write body header text
                    markerStyle.fontSize = 24;
                    XWPFParagraph newHeader = newDoc.CreateParagraph(markerStyle);
                    XWPFRun headerTitle = newHeader.CreateRun(markerStyle);
                    headerTitle.SetText(hMarker.HeaderText);

                    break;
                case FMarker fMarker:
                    string footnoteId;
                    switch (fMarker.FootNoteCaller)
                    {
                        case "-":
                            footnoteId = "";
                            break;
                        case "+":
                            footnoteId = nextFootnoteNum.ToString();
                            nextFootnoteNum++;
                            break;
                        default:
                            footnoteId = fMarker.FootNoteCaller;
                            break;
                    }
                    CT_FtnEdn footnote = new CT_FtnEdn();
                    footnote.id = footnoteId;
                    footnote.type = ST_FtnEdn.normal;
                    CT_P footnoteParagraph = footnote.AddNewP();
                    footnoteParagraph.AddNewR().AddNewT().Value = footnoteId + " Placeholder Text";
                    parentParagraph.Document.AddFootnote(footnote);

                    XWPFRun footnoteReferenceRun = parentParagraph.CreateRun();
                    CT_RPr rpr = footnoteReferenceRun.GetCTR().AddNewRPr();
                    rpr.rStyle = new CT_String();
                    rpr.rStyle.val = "FootnoteReference";
                    CT_FtnEdnRef footnoteReference = new CT_FtnEdnRef();
                    footnoteReference.id = footnoteId;
                    footnoteReference.isEndnote = false;
                    footnoteReferenceRun.GetCTR().Items.Add(footnoteReference);

                    //FootnoteMarkers[footnoteId] = fMarker;

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
                    VerseReference.SetText(fRMarker.VerseReference);
                    break;
                case FKMarker fKMarker:
                    XWPFRun FootNoteKeyword = parentParagraph.CreateRun(markerStyle);
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

                    crossRefMarker.SetText(crossId);
                    crossRefMarker.Subscript = VerticalAlign.SUPERSCRIPT;

                    CrossRefMarkers[crossId] = xMarker;
                    break;
                case XOMarker xOMarker:
                    markerStyle.isBold = true;
                    XWPFRun CrossVerseReference = parentParagraph.CreateRun(markerStyle);
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
                    tableContainer.SetBottomBorder(XWPFTable.XWPFBorderType.NONE,0,0,"#FFFFFFF");
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
                    newLineBreak.AddBreak(BreakType.TEXTWRAPPING);
                    break;
                case IDMarker _:
                    // This is the start of a new book.
                    beforeFirstChapter = true;
                    chapterLabel = chapterLabelDefault;
                    nextFootnoteNum = 1;
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
            CT_R ctr = run.GetCTR();
            CT_Text text = ctr.AddNewT();
            text.Value = " ";
            text.space = "preserve";
        }

        //private void RenderFootnotes(StyleConfig styles)
        //{
        //    if (FootnoteMarkers.Count > 0)
        //    {
        //        XWPFParagraph renderFootnoteStart = newDoc.CreateParagraph();
        //        renderFootnoteStart.BorderTop = Borders.Single;

        //        StyleConfig markerStyle = (StyleConfig)styles.Clone();
        //        markerStyle.fontSize = 12;

        //        foreach (KeyValuePair<string,Marker> footnoteKVP in FootnoteMarkers)
        //        {
        //            XWPFParagraph renderFootnote = newDoc.CreateParagraph(markerStyle);
        //            XWPFRun footnoteMarker = renderFootnote.CreateRun(markerStyle);
        //            footnoteMarker.SetText(footnoteKVP.Key);
        //            footnoteMarker.Subscript = VerticalAlign.SUPERSCRIPT;
        //            foreach(Marker input in footnoteKVP.Value.Contents)
        //            {
        //                RenderMarker(input, markerStyle, renderFootnote);
        //            }
        //        }
        //        FootnoteMarkers.Clear();
        //    }
        //}

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
                    XWPFRun crossRefMarker = renderCrossRef.CreateRun(markerStyle);
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
        public void createFooter()
        {
            // Footer Object
            CT_Ftr footer = new CT_Ftr();
            CT_P footerParagraph = footer.AddNewP();
            CT_PPr ppr = footerParagraph.AddNewPPr();
            CT_Jc align = ppr.AddNewJc();
            align.val = ST_Jc.center;

            // Page Number Format OOXML
            footerParagraph.AddNewR().AddNewFldChar().fldCharType = ST_FldCharType.begin;
            CT_Text pageNumber = footerParagraph.AddNewR().AddNewInstrText();
            pageNumber.Value = "PAGE   \\* MERGEFORMAT";
            pageNumber.space = "preserve";
            footerParagraph.AddNewR().AddNewFldChar().fldCharType = ST_FldCharType.separate;

            CT_R centerRun = footerParagraph.AddNewR();
            centerRun.AddNewRPr().AddNewNoProof();
            centerRun.AddNewT().Value = "2";

            CT_R endRun= footerParagraph.AddNewR();
            endRun.AddNewRPr().AddNewNoProof();
            endRun.AddNewFldChar().fldCharType = ST_FldCharType.end;


            // Linking to Footer Style Object to Document
            XWPFRelation footerRelation = XWPFRelation.FOOTER;
            XWPFFooter documentFooter = (XWPFFooter)newDoc.CreateRelationship(footerRelation, XWPFFactory.GetInstance(), newDoc.FooterList.Count + 1);
            documentFooter.SetHeaderFooter(footer);
            CT_HdrFtrRef footerRef = newDoc.Document.body.sectPr.AddNewFooterReference();
            footerRef.type = ST_HdrFtr.@default;
            footerRef.id = documentFooter.GetPackageRelationship().Id;

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
            align.val = ST_Jc.left;
            headerParagraph.AddNewR().AddNewT().Value = bookname;

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

            // Increment page header count so each one gets a unique ID
            pageHeaderCount++;
        }
        public void getRenderedRows(Marker input, StyleConfig config,XWPFTable parentTable)
        {
            XWPFTableRow tableRowContainer = parentTable.CreateRow();
            foreach (Marker marker in input.Contents)
            {
                getRenderedCell(marker, config, tableRowContainer);
            }
        }
        public void getRenderedCell(Marker input,StyleConfig config, XWPFTableRow parentRow)
        {
            StyleConfig markerStyle = (StyleConfig)config.Clone();
            XWPFTableCell tableCellContainer = parentRow.CreateCell();
            XWPFParagraph cellContents;
            switch (input)
            {
                case THMarker tHMarker:
                    markerStyle.isBold = true;
                    cellContents = tableCellContainer.AddParagraph(markerStyle);
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, cellContents);
                    }
                    break;
                case THRMarker tHRMarker:
                    markerStyle.isAlignRight = true;
                    markerStyle.isBold = true;
                    cellContents = tableCellContainer.AddParagraph(markerStyle);
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, cellContents);
                    }
                    break;
                case TCMarker tCMarker:
                    cellContents = tableCellContainer.AddParagraph(markerStyle);
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, cellContents);
                    }
                    break;
                case TCRMarker tCRMarker:
                    markerStyle.isAlignRight = true;
                    cellContents = tableCellContainer.AddParagraph(markerStyle);
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, markerStyle, cellContents);
                    }
                    break;
                default:
                    break;
            }
        }

    }
}
