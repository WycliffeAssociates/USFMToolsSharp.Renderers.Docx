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

        }

        public TOCBuilder(CT_SdtBlock block)
        {
            this.block = block;
        }

        public TOCBuilder init()
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
            pPr.AddNewPStyle().val = "TOC1";
            var tab = pPr.AddNewTabs().AddNewTab();
            tab.leader = ST_TabTlc.dot;
            tab.pos = "9350";
            tab.val = ST_TabJc.right;

            pPr.AddNewRPr().AddNewNoProof();

            CT_R run = p.AddNewR();
            run.AddNewFldChar().fldCharType = ST_FldCharType.begin;

            run = p.AddNewR();
            CT_Text text = run.AddNewInstrText();
            text.space = "preserve";
            text.Value = " TOC \\h \\z";

            p.AddNewR().AddNewFldChar().fldCharType = ST_FldCharType.separate;

            return this;
        }

        public void AddRowTOC(int level, String title, int page, String bookmarkRef)
        {
            CT_SdtContentBlock contentBlock = this.block.sdtContent;
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
            tab.pos = "8290"; //(new BigInteger("8290"));
            pPr.AddNewRPr().AddNewNoProof();
            CT_R Run = p.AddNewR();
            Run.AddNewRPr().AddNewNoProof();
            Run.AddNewT().Value = (title);
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
            text.space = "preserve";// (Space.PRESERVE);
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

        public void BuildLowLevelTOC(CT_SdtBlock block, IDictionary<string, string> entries)
        {
            var sdtPr = block.AddNewSdtPr();
            var docPartObj = sdtPr.AddNewDocPartObj();
            docPartObj.AddNewDocPartGallery().val = "Table of Contents";
            
            var sdtContent = block.AddNewSdtContent();
            CT_P p = sdtContent.AddNewP();
            
            p.AddNewPPr().AddNewPStyle().val = "TOCHeading";
            CT_R run = p.AddNewR();
            run.AddNewT().Value = "Table of Contents";
            run.AddNewRPr().AddNewSz().val = (ulong) 24 * 2;

            // TOC Field
            p = sdtContent.AddNewP();
            CT_PPr pPr = p.AddNewPPr();
            pPr.AddNewPStyle().val = "TOC1";
            var tab = pPr.AddNewTabs().AddNewTab();
            tab.leader = ST_TabTlc.dot;
            tab.pos = "9350";
            tab.val = ST_TabJc.right;

            pPr.AddNewRPr().AddNewNoProof();

            run = p.AddNewR();
            run.AddNewFldChar().fldCharType = ST_FldCharType.begin;

            run = p.AddNewR();
            CT_Text text = run.AddNewInstrText();
            text.space = "preserve";
            text.Value = " TOC \\h";

            p.AddNewR().AddNewFldChar().fldCharType = ST_FldCharType.separate;

            // TOC rows...
            AddLowLevelRows(sdtContent, entries);


            // end toc
            p = sdtContent.AddNewP();
            run = p.AddNewR();
            run.AddNewRPr().AddNewNoProof();
            run.AddNewFldChar().fldCharType = ST_FldCharType.end;
        }

        private void AddLowLevelRows(CT_SdtContentBlock sdtContent, IDictionary<string, string> entries)
        {
            CT_P p;
            CT_PPr pPr;
            CT_R run;
            CT_RPr rPr;
            CT_Text text;
            CT_FldChar fieldChar;

            foreach (var entry in entries)
            {
                p = sdtContent.AddNewP();
                pPr = p.AddNewPPr();
                pPr.AddNewPStyle().val = "TOC1";
                var tab = pPr.AddNewTabs().AddNewTab();
                tab.leader = ST_TabTlc.dot;
                tab.pos = "8100";
                tab.val = ST_TabJc.right;
                pPr.AddNewRPr().AddNewNoProof();
                
                // TOC entry name
                run = p.AddNewR();
                rPr = run.AddNewRPr();
                rPr.AddNewNoProof();
                rPr.AddNewSz().val = (ulong) 12 * 2;
                text = run.AddNewT();
                text.Value = entry.Key;

                // add tab
                run = p.AddNewR();
                run.AddNewRPr().AddNewNoProof();
                run.AddNewTab();

                // add field code - reference
                run = p.AddNewR();
                run.AddNewRPr().AddNewNoProof();
                fieldChar = run.AddNewFldChar();
                fieldChar.fldCharType = ST_FldCharType.begin;

                run = p.AddNewR();
                run.AddNewRPr().AddNewNoProof();
                text = run.AddNewInstrText();
                text.space = "preserve";
                text.Value = $" PAGEREF _Toc{entry.Value} \\h ";

                run = p.AddNewR();
                run.AddNewRPr().AddNewNoProof();
                fieldChar = run.AddNewFldChar();
                fieldChar.fldCharType = ST_FldCharType.separate;

                // display page #
                run = p.AddNewR();
                rPr = run.AddNewRPr();
                rPr.AddNewNoProof();
                rPr.AddNewSz().val = (ulong) 12 * 2;
                run.AddNewT().Value = "1";

                run = p.AddNewR();
                run.AddNewRPr().AddNewNoProof();
                run.AddNewFldChar().fldCharType = ST_FldCharType.end;
            }
        }
    }
}
