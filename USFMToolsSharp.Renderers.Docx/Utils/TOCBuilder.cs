using NPOI.OpenXmlFormats.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Text;

namespace USFMToolsSharp.Renderers.Docx.Utils
{
    class TOCBuilder
    {
        public static readonly string TOC_BOOKMARK = "_Toc{0}";

        public int HeaderSize = 24;
        public int RowSize = 12;

        private CT_SdtBlock block;

        public TOCBuilder()
        {
            block = new CT_SdtBlock();
            init();
        }

        public TOCBuilder(CT_SdtBlock block)
        {
            this.block = block;
            init();
        }

        private void init()
        {
            CT_SdtPr sdtPr = block.AddNewSdtPr();
            CT_DecimalNumber id = sdtPr.AddNewId();
            id.val = ("4844945");
            sdtPr.AddNewDocPartObj().AddNewDocPartGallery().val = ("Table of Contents");
            CT_SdtEndPr sdtEndPr = block.AddNewSdtEndPr();
            CT_RPr rPr = sdtEndPr.AddNewRPr();
            CT_Fonts fonts = rPr.AddNewRFonts();
            fonts.asciiTheme = (ST_Theme.minorHAnsi);
            fonts.eastAsiaTheme = (ST_Theme.minorHAnsi);
            fonts.hAnsiTheme = (ST_Theme.minorHAnsi);
            fonts.cstheme = (ST_Theme.minorBidi);
            CT_SdtContentBlock content = block.AddNewSdtContent();
            CT_P p = content.AddNewP();
            byte[] b = Encoding.Unicode.GetBytes("00EF7E24");
            p.rsidR = b;
            p.rsidRDefault = b;
            CT_PPr pPr = p.AddNewPPr();
            pPr.AddNewPStyle().val = ("TOCHeading");
            pPr.AddNewJc().val = ST_Jc.center;
            CT_R run = p.AddNewR();
            run.AddNewRPr().AddNewSz().val = (ulong)HeaderSize * 2;
            run.AddNewT().Value = ("Table of Contents");
            run.AddNewBr().type = ST_BrType.textWrapping; // line break

            // TOC Field
            p = content.AddNewP();
            pPr = p.AddNewPPr();
            pPr.AddNewPStyle().val = "TOC1";
            pPr.AddNewRPr().AddNewNoProof();

            run = p.AddNewR();
            run.AddNewFldChar().fldCharType = ST_FldCharType.begin;

            run = p.AddNewR();
            CT_Text text = run.AddNewInstrText();
            text.space = "preserve";
            text.Value = (" TOC \\h \\z ");

            p.AddNewR().AddNewFldChar().fldCharType = ST_FldCharType.separate;
        }

        /*
         * (Optional) Adds a custom row to TOC.
         * Will be wiped off when the user updates TOC.
         */
        public void AddRow(int level, string title, string bookmarkRef, int page = 1)
        {
            CT_SdtContentBlock contentBlock = block.sdtContent;
            CT_P p = contentBlock.AddNewP();
            byte[] b = Encoding.Unicode.GetBytes("00EF7E24");
            p.rsidR = b;
            p.rsidRDefault = b;
            CT_PPr pPr = p.AddNewPPr();
            pPr.AddNewPStyle().val = ("TOC" + level);
            CT_Tabs tabs = pPr.AddNewTabs();
            CT_TabStop tab = tabs.AddNewTab();
            tab.val = (ST_TabJc.right);
            tab.leader = (ST_TabTlc.dot);
            tab.pos = "8100";
            pPr.AddNewRPr().AddNewNoProof();
            CT_R Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
            Run.rPr.AddNewSz().val = (ulong)RowSize * 2;
            Run.AddNewT().Value = title;
            Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
            Run.AddNewTab();
            Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
            Run.AddNewFldChar().fldCharType = (ST_FldCharType.begin);
            // pageref run
            Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
            Run.rPr.AddNewSz().val = (ulong)RowSize * 2;
            Run.rPr.AddNewSzCs().val = (ulong)RowSize * 2;
            CT_Text text = Run.AddNewInstrText();
            text.space = "preserve";
            // bookmark reference
            text.Value = ($" PAGEREF {string.Format(TOC_BOOKMARK, bookmarkRef)} \\h \\z ");
            p.AddNewR().AddNewRPr().AddNewNoProof();
            Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
            Run.AddNewFldChar().fldCharType = (ST_FldCharType.separate);
            // page number run
            Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
            Run.rPr.AddNewSz().val = (ulong)RowSize * 2;
            Run.AddNewT().Value = page.ToString();
            Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
            Run.AddNewFldChar().fldCharType = (ST_FldCharType.end);
        }

        public CT_SdtBlock Build()
        {
            // append "end" field char for TOC
            CT_SdtContentBlock contentBlock = block.sdtContent;
            CT_P p = contentBlock.AddNewP();
            CT_R run = p.AddNewR();
            run.AddNewRPr().AddNewNoProof();
            run.AddNewFldChar().fldCharType = ST_FldCharType.end;

            return block;
        }
    }
}
