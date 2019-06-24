using System;
using System.Collections.Generic;
using System.Text;
using USFMToolsSharp.Models.Markers;
using NPOI.XWPF.UserModel;
using System.IO;

namespace USFMToolsSharp.Renderers.Docx
{
    public class DocxRenderer
    {
        public List<string> UnrenderableTags;
        public Dictionary<string,List<Marker>> FootnoteTextTags;
        
        private XWPFDocument newDoc;

        private bool isSingleSpaced = true;
        private bool hasOneColumn = true;
        private bool isL2RDirection = true;
        private bool isTextJustified = false;
        private bool separateChapters = true;

        public string FrontMatterDocx { get; set; }
        public string InsertedFooter { get; set; }
        public string InsertedHead { get; set; }

        public DocxRenderer()
        {
            UnrenderableTags = new List<string>();
            FootnoteTextTags = new Dictionary<string,List<Marker>>();
            newDoc = new XWPFDocument();
        }
        public DocxRenderer(bool isSingleSpaced = true, bool hasOneColumn = true, bool isL2RDirection = true, bool isTextJustified = false, bool separateChapters = false)
        {
            this.isSingleSpaced = isSingleSpaced;
            this.hasOneColumn = hasOneColumn;
            this.isL2RDirection = isL2RDirection;
            this.isTextJustified = isTextJustified;
            this.separateChapters = separateChapters;

            UnrenderableTags = new List<string>();
            FootnoteTextTags = new Dictionary<string, List<Marker>>();
            newDoc = new XWPFDocument();

        }


        public XWPFDocument Render(USFMDocument input)
        {

                foreach (Marker marker in input.Contents)
                {

                    RenderMarker(marker);

                }
            return newDoc;

        }
        private void RenderMarker(Marker input, XWPFParagraph parentParagraph = null, bool isBold = false, bool isItalics = false)
        {
            
            switch (input)
            {
                case PMarker _:
                    
                        XWPFParagraph newParagraph = newDoc.CreateParagraph();
                        foreach (Marker marker in input.Contents)
                        {
                            RenderMarker(marker, newParagraph);
                        }
                    break;
                case CMarker cMarker:
                    XWPFParagraph newChapter = newDoc.CreateParagraph();
                    XWPFRun chapterMarker = newChapter.CreateRun();
                    chapterMarker.SetText($"{cMarker.Number}");

                    chapterMarker.FontSize = 24;

                    XWPFParagraph chapterVerses = newDoc.CreateParagraph();
                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, chapterVerses);
                    }
                    

                    RenderFootnotes();

                    // Page breaks after each chapter
                    if (separateChapters)
                    {
                        newDoc.CreateParagraph().CreateRun().AddBreak(BreakType.PAGE);
                    }

                    break;
                case VMarker vMarker:
                    

                    XWPFRun verseMarker = parentParagraph.CreateRun();
                    verseMarker.SetText($"  {vMarker.VerseCharacter}  ");
                    verseMarker.Subscript = VerticalAlign.SUPERSCRIPT;


                    foreach (Marker marker in input.Contents)
                    {
                        RenderMarker(marker, parentParagraph);
                    }
                    break;
                case QMarker qMarker:
                    XWPFParagraph poetryParagraph = newDoc.CreateParagraph();

                    // Not sure if indentation works
                    poetryParagraph.IndentationLeft =qMarker.Depth;

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
                        RenderMarker(marker,parentParagraph,true);
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
                    if (!separateChapters)   // No double page breaks before books
                    {
                        newDoc.CreateParagraph().CreateRun().AddBreak(BreakType.PAGE);
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

                    footnoteMarker.SetText($"{footnoteId}");
                    footnoteMarker.Subscript = VerticalAlign.SUBSCRIPT;

                    FootnoteTextTags[footnoteId] = new List<Marker>();

                    foreach (Marker marker in input.Contents)
                    {
                        FootnoteTextTags[footnoteId].Add(marker);

                    }
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
                        RenderMarker(marker, parentParagraph,false,true);
                    }
                    break;
                case FQAEndMarker fQAEndMarker:
                    break;
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
                XWPFParagraph renderFootnoteHeader = newDoc.CreateParagraph();
                XWPFRun FootnoteHeader = renderFootnoteHeader.CreateRun();
                FootnoteHeader.SetText("Footnotes");
                FootnoteHeader.FontSize = 24;

                foreach(KeyValuePair<string,List<Marker>> footnoteKVP in FootnoteTextTags)
                {
                    XWPFParagraph renderFootnote = newDoc.CreateParagraph();
                    XWPFRun footnoteMarker = renderFootnote.CreateRun();
                    footnoteMarker.SetText(footnoteKVP.Key);
                    footnoteMarker.Subscript = VerticalAlign.SUBSCRIPT;
                    

                    foreach(Marker input in footnoteKVP.Value)
                    {
                        RenderMarker(input, renderFootnote);
                    }
                    
                }
                FootnoteTextTags.Clear();
            }
        }
        

    }
}
