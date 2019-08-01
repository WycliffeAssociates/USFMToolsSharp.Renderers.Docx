using System.Collections.Generic;
using System.Text;
using USFMToolsSharp.Models.Markers;
using NPOI.XWPF.UserModel;
using NPOI.XWPF.Model;
using NPOI.OpenXmlFormats.Wordprocessing;

namespace USFMToolsSharp.Renderers.Docx
{
    public class DocxRenderer
    {
        public List<string> UnrenderableMarkers;
        public Dictionary<string,Marker> FootnoteTextTags;
        private DocxConfig configDocx;
        private XWPFDocument newDoc;
        private int bookNameCount=0;
        public List<string> BookNames;

        public DocxRenderer()
        {
            configDocx = new DocxConfig();

            UnrenderableMarkers = new List<string>();
            BookNames = new List<string>();
            FootnoteTextTags = new Dictionary<string,Marker>();
            newDoc = new XWPFDocument();
        }
        public DocxRenderer(DocxConfig config)
        {
            configDocx = config;

            UnrenderableMarkers = new List<string>();
            BookNames = new List<string>();
            FootnoteTextTags = new Dictionary<string,Marker>();
            newDoc = new XWPFDocument();

        }
        public XWPFDocument Render(USFMDocument input)
        {
            setStartPageNumber();

            newDoc.ColumnCount = configDocx.columnCount;
            newDoc.TextDirection= configDocx.textDirection;

            foreach (Marker marker in input.Contents)
                {

                    RenderMarker(marker);

                }
            createBookHeaders(BookNames[BookNames.Count - 1]);
            createFooter();
            return newDoc;

        }
        private void RenderMarker(Marker input, XWPFParagraph parentParagraph = null, bool isBold = false, bool isItalics = false)
        {
            switch (input)
            {
                case PMarker _:
                    XWPFParagraph newParagraph = newDoc.CreateParagraph();

                    newParagraph.Alignment = configDocx.textAlign;
                    newParagraph.SpacingBetween = configDocx.lineSpacing;
                        
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, newParagraph);
                    }
                    break;
                case CMarker cMarker:
                    XWPFParagraph newChapter = newDoc.CreateParagraph();
                    XWPFRun chapterMarker = newChapter.CreateRun();
                    chapterMarker.SetText(cMarker.Number.ToString());
                    chapterMarker.FontSize = 24;

                    XWPFParagraph chapterVerses = newDoc.CreateParagraph();
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, chapterVerses);
                    }

                    RenderFootnotes();
                    if (configDocx.separateChapters)
                    {
                        newDoc.CreateParagraph().CreateRun().AddBreak(BreakType.PAGE);
                    }

                    break;
                case VMarker vMarker:

                    if (configDocx.separateVerses)
                    {
                        XWPFRun newLine = parentParagraph.CreateRun();
                        newLine.AddBreak(BreakType.TEXTWRAPPING);
                    }
                    XWPFRun verseMarker = parentParagraph.CreateRun();

                    verseMarker.SetText(vMarker.VerseCharacter);
                    verseMarker.Subscript = VerticalAlign.SUPERSCRIPT;

                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, parentParagraph);
                    }
                    break;
                case QMarker qMarker:
                    XWPFParagraph poetryParagraph = newDoc.CreateParagraph();
                    poetryParagraph.IndentationLeft = qMarker.Depth;

                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker,poetryParagraph);
                    }
                    break;
                case MMarker mMarker:
                    break;
                case TextBlock textBlock:
                    XWPFRun blockText = parentParagraph.CreateRun();
                    blockText.SetText(textBlock.Text);
                    blockText.FontSize = 16;

                    if (isBold)
                    {
                        blockText.IsBold = true;
                    }
                    if (isItalics)
                    {
                        blockText.IsItalic = true;
                    }
                    break;
                case BDMarker bdMarker:
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker,parentParagraph,isBold:true);
                    }
                    break;
                case HMarker hMarker:
                    XWPFParagraph newHeader = newDoc.CreateParagraph();
                    XWPFRun headerTitle = newHeader.CreateRun();
                    headerTitle.SetText(hMarker.HeaderText);
                    headerTitle.FontSize = 24;
                    break;
                case MTMarker mTMarker:
                    createBookHeaders(mTMarker.Title);
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker);
                    }
                    break;
                case FMarker fMarker:
                    StringBuilder footnote = new StringBuilder();

                    string footnoteId;
                    switch (fMarker.FootNoteCaller)
                    {
                        case "-":
                            footnoteId = "";
                            break;
                        case "+":
                            footnoteId = $"{FootnoteTextTags.Count + 1}";
                            break;
                        default:
                            footnoteId = fMarker.FootNoteCaller;
                            break;
                    }
                    XWPFRun footnoteMarker = parentParagraph.CreateRun();

                    footnoteMarker.SetText(footnoteId);
                    footnoteMarker.Subscript = VerticalAlign.SUPERSCRIPT;

                    FootnoteTextTags[footnoteId] = fMarker;

                    break;
                case FTMarker fTMarker:

                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker,parentParagraph);
                    }

                    break;
                case FQAMarker fQAMarker:
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, parentParagraph,isItalics:true);
                    }
                    break;
                case FQAEndMarker fQAEndMarker:
                case FEndMarker _:
                case IDEMarker _:
                case IDMarker _:
                case VPMarker _:
                case VPEndMarker _:
                    break;
                default:
                    UnrenderableMarkers.Add(input.Identifier);
                    break;
            }
        }
        private void RenderFootnotes()
        {

            if (FootnoteTextTags.Count > 0)
            {
                XWPFParagraph renderFootnoteHeader = newDoc.CreateParagraph();
                XWPFRun FootnoteHeader = renderFootnoteHeader.CreateRun();
                FootnoteHeader.SetText("Footnotes");
                FootnoteHeader.FontSize = 24;

                foreach(KeyValuePair<string,Marker> footnoteKVP in FootnoteTextTags)
                {
                    XWPFParagraph renderFootnote = newDoc.CreateParagraph();
                    XWPFRun footnoteMarker = renderFootnote.CreateRun();
                    footnoteMarker.SetText(footnoteKVP.Key);
                    footnoteMarker.Subscript = VerticalAlign.SUPERSCRIPT;

                    foreach(Marker marker in footnoteKVP.Value.Contents)
                    {
                        RenderMarker(marker, renderFootnote);
                    }
  
                }
                FootnoteTextTags.Clear();
            }
        }
        public void setStartPageNumber()
        {
            newDoc.Document.body.sectPr.pgNumType.fmt = ST_NumberFormat.@decimal;
            newDoc.Document.body.sectPr.pgNumType.start = "0";
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

        public void createBookHeaders(string bookname)
        {
            BookNames.Add(bookname);
            if (BookNames.Count > 1)
            {
                CT_Hdr header = new CT_Hdr();
                CT_P headerParagraph = header.AddNewP();
                CT_PPr ppr = headerParagraph.AddNewPPr();
                CT_Jc align = ppr.AddNewJc();
                align.val = ST_Jc.left;

                headerParagraph.AddNewR().AddNewT().Value = BookNames[bookNameCount-1];

                XWPFRelation headerRelation = XWPFRelation.HEADER;

                // newDoc.HeaderList doesn't update with header additions
                XWPFHeader documentHeader = (XWPFHeader)newDoc.CreateRelationship(headerRelation, XWPFFactory.GetInstance(), bookNameCount);
                documentHeader.SetHeaderFooter(header);
                CT_SectPr diffHeader = newDoc.Document.body.AddNewP().AddNewPPr().createSectPr();
                CT_HdrFtrRef headerRef = diffHeader.AddNewHeaderReference();
                headerRef.type = ST_HdrFtr.@default;
                headerRef.id = documentHeader.GetPackageRelationship().Id;
            
                
            }
            bookNameCount++;
        }


    }
}
