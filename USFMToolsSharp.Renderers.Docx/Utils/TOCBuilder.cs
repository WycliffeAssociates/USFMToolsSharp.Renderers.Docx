using NPOI.OpenXmlFormats.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Text;

namespace USFMToolsSharp.Renderers.Docx.Utils
{
    class TOCBuilder
    {
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
            sdtPr.AddNewDocPartObj().AddNewDocPartGallery().val = ("Table of contents");
            CT_SdtEndPr sdtEndPr = block.AddNewSdtEndPr();
            CT_RPr rPr = sdtEndPr.AddNewRPr();
            CT_Fonts fonts = rPr.AddNewRFonts();
            fonts.asciiTheme = (ST_Theme.minorHAnsi);
            fonts.eastAsiaTheme = (ST_Theme.minorHAnsi);
            fonts.hAnsiTheme = (ST_Theme.minorHAnsi);
            fonts.cstheme = (ST_Theme.minorBidi);
            rPr.AddNewB().val = false;
            rPr.AddNewBCs().val = false;
            rPr.AddNewColor().val = ("auto");
            rPr.AddNewSz().val = 24;
            rPr.AddNewSzCs().val = 24;
            CT_SdtContentBlock content = block.AddNewSdtContent();
            CT_P p = content.AddNewP();
            byte[] b = Encoding.Unicode.GetBytes("00EF7E24");
            p.rsidR = b;
            p.rsidRDefault = b;
            p.AddNewPPr().AddNewPStyle().val = ("TOCHeading");
            p.AddNewR().AddNewT().Value = ("Table of Contents");

            // TOC Field
            p = content.AddNewP();
            CT_PPr pPr = p.AddNewPPr();
            pPr.AddNewPStyle().val = "TOCHeading";
            pPr.AddNewRPr().AddNewNoProof();

            CT_R run = p.AddNewR();
            run.AddNewFldChar().fldCharType = ST_FldCharType.begin;

            run = p.AddNewR();
            CT_Text text = run.AddNewInstrText();
            text.space = "preserve";
            text.Value = (" TOC \\h \\z");

            p.AddNewR().AddNewFldChar().fldCharType = ST_FldCharType.separate;
        }

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
            tab.pos = "8290";
            pPr.AddNewRPr().AddNewNoProof();
            CT_R Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
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
            CT_Text text = Run.AddNewInstrText();
            text.space = "preserve";
            // bookmark reference
            text.Value = (" PAGEREF _Toc" + bookmarkRef + " \\h \\z");
            p.AddNewR().AddNewRPr().AddNewNoProof();
            Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
            Run.AddNewFldChar().fldCharType = (ST_FldCharType.separate);
            // page number run
            Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
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
