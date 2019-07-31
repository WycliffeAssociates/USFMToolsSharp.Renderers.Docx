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
        public List<string> UnrenderableTags;
        public Dictionary<string,Marker> FootnoteTextTags;
        private DocxConfig configDocx;
        private XWPFDocument newDoc;
        private int bookNameCount=1;

        public DocxRenderer()
        {
            configDocx = new DocxConfig();

            UnrenderableTags = new List<string>();
            FootnoteTextTags = new Dictionary<string,Marker>();
            newDoc = new XWPFDocument();
        }
        public DocxRenderer(DocxConfig config)
        {
            configDocx = config;

            UnrenderableTags = new List<string>();
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
            return newDoc;

        }
        private void RenderMarker(Marker input, XWPFParagraph parentParagraph = null, bool isBold = false, bool isItalics = false,int fontSize = 16)
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
                    blockText.FontSize = fontSize;

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

                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker);
                    }
                    if (!configDocx.separateChapters)   // No double page breaks before books
                    {
                        newDoc.CreateParagraph().CreateRun().AddBreak(BreakType.PAGE);
                    }
                    createBookHeaders(mTMarker.Title);
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
                    footnoteMarker.Subscript = VerticalAlign.SUBSCRIPT;

                    FootnoteTextTags[footnoteId] = fMarker;

                    break;
                case FPMarker fPMarker:
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, parentParagraph, fontSize: 12);
                    }
                    break;
                case FTMarker fTMarker:

                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker,parentParagraph,fontSize:12);
                    }
                    break;
                case FRMarker fRMarker:
                    XWPFRun VerseReference = parentParagraph.CreateRun();
                    VerseReference.SetText(fRMarker.VerseReference);
                    VerseReference.IsBold = true;
                    VerseReference.FontSize = fontSize;
                    break;
                case FKMarker fKMarker:
                    XWPFRun FootNoteKeyword = parentParagraph.CreateRun();
                    FootNoteKeyword.SetText($" {fKMarker.FootNoteKeyword.ToUpper()}: ");
                    FootNoteKeyword.FontSize = fontSize;
                    break;
                case FQMarker fQMarker:
                case FQAMarker fQAMarker:
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, parentParagraph,isItalics:true,fontSize: 12);
                    }
                    break;
                case FVMarker fVMarker:
                    XWPFRun footnoteVerseMarker = parentParagraph.CreateRun();
                    footnoteVerseMarker.SetText(fVMarker.VerseCharacter);
                    footnoteVerseMarker.Subscript = VerticalAlign.SUPERSCRIPT;
                    break;
                case FQEndMarker _:
                case FQAEndMarker _:
                case FEndMarker _:
                case IDEMarker _:
                case IDMarker _:
                case VPMarker _:
                case VPEndMarker _:
                    break;
                default:
                    UnrenderableTags.Add(input.Identifier);
                    break;
            }
        }
        private void RenderFootnotes()
        {

            if (FootnoteTextTags.Count > 0)
            {
                XWPFParagraph renderFootnoteStart = newDoc.CreateParagraph();
                renderFootnoteStart.BorderTop = Borders.Single;

                foreach(KeyValuePair<string,Marker> footnoteKVP in FootnoteTextTags)
                {
                    XWPFParagraph renderFootnote = newDoc.CreateParagraph();
                    XWPFRun footnoteMarker = renderFootnote.CreateRun();
                    footnoteMarker.SetText(footnoteKVP.Key);
                    footnoteMarker.Subscript = VerticalAlign.SUPERSCRIPT;

                    foreach(Marker marker in footnoteKVP.Value.Contents)
                    {
                        RenderMarker(marker, renderFootnote,fontSize:12);
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

            CT_Hdr header = new CT_Hdr();
            CT_P headerParagraph = header.AddNewP();
            CT_PPr ppr = headerParagraph.AddNewPPr();
            CT_Jc align = ppr.AddNewJc();
            align.val = ST_Jc.left;

            headerParagraph.AddNewR().AddNewT().Value = bookname;

            XWPFRelation headerRelation = XWPFRelation.HEADER;

            // newDoc.HeaderList doesn't update with header additions
            XWPFHeader documentHeader = (XWPFHeader)newDoc.CreateRelationship(headerRelation, XWPFFactory.GetInstance(), bookNameCount);
            documentHeader.SetHeaderFooter(header);
            CT_SectPr diffHeader = newDoc.Document.body.AddNewP().AddNewPPr().createSectPr();
            CT_HdrFtrRef headerRef = diffHeader.AddNewHeaderReference();
            headerRef.type = ST_HdrFtr.@default;
            headerRef.id = documentHeader.GetPackageRelationship().Id;

            bookNameCount++;
        }


    }
}
