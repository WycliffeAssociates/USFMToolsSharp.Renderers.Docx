﻿using NPOI.XWPF.UserModel;
using System;
using NPOI.OpenXmlFormats.Wordprocessing;

using System.Collections.Generic;
using System.Text;

namespace USFMToolsSharp.Renderers.Docx.Utils
{
    class StylesBuilder
    {
        private XWPFStyles documentStyles;
        private CT_Styles ctStyles;

        public StylesBuilder(XWPFDocument docxDocument)
        {
            documentStyles = docxDocument.CreateStyles();
            ctStyles = new CT_Styles();
        }

        /// <summary>
        /// Sets default font, size for text in the document.
        /// </summary>
        public void AddDefaultStyle()
        {
            CT_DocDefaults docDefaults = ctStyles.AddNewDocDefaults();
            CT_RPrDefault rprDefault = docDefaults.AddNewRPrDefault();
            CT_RPr rpr = rprDefault.AddNewRPr();
            rpr.AddNewSz().val = 24;
            rpr.AddNewSzCs().val = 24;
            var font = rpr.AddNewRFonts();
            font.asciiTheme = ST_Theme.minorAscii;
            font.cstheme = ST_Theme.minorBidi;
            font.eastAsiaTheme = ST_Theme.minorHAnsi;
            font.hAnsiTheme = ST_Theme.minorHAnsi;
        }

        public void AddCustomHeadingStyle(string name, int headingLevel, int outlineLevel, int ptSize = 12)
        {

            CT_Style ctStyle = ctStyles.AddNewStyle();
            ctStyle.styleId = (name);

            CT_String styleName = new CT_String();
            styleName.val = (name);
            ctStyle.name = (styleName);

            CT_DecimalNumber indentNumber = new CT_DecimalNumber();
            indentNumber.val = headingLevel.ToString();

            // lower number > style is more prominent in the formats bar
            ctStyle.uiPriority = (indentNumber);

            CT_OnOff onoffnull = new CT_OnOff();
            ctStyle.unhideWhenUsed = (onoffnull);

            // style shows up in the formats bar
            ctStyle.qFormat = (onoffnull);

            // style defines a heading of the given level
            CT_PPr ppr = new CT_PPr();
            ppr.outlineLvl = new CT_DecimalNumber() { val = outlineLevel.ToString() };
            ctStyle.pPr = (ppr);

            CT_RPr rpr = new CT_RPr();
            rpr.AddNewSz().val = (ulong)ptSize * 2;
            ctStyle.rPr = rpr;

            XWPFStyle style = new XWPFStyle(ctStyle);
            style.StyleType = (ST_StyleType.paragraph);

        }

        public void Build()
        {
            documentStyles.SetStyles(ctStyles);
        }

        /// <summary>
        /// Builds the styles required for rendering & handling
        /// Table of Contents (TOC).
        /// 
        /// Please set paragraph style to "Heading{#}" in order to
        /// have it rendered in the TOC.
        /// </summary>
        /// <param name="doc"></param>
        public static void BuildStylesForTOC(XWPFDocument doc)
        {
            var styleBuilder = new StylesBuilder(doc);
            styleBuilder.AddDefaultStyle();
            styleBuilder.AddCustomHeadingStyle("TOCHeading", 1, 9);
            styleBuilder.AddCustomHeadingStyle("TOC1", 2, 0);
            styleBuilder.AddCustomHeadingStyle("TOC2", 3, 0);
            styleBuilder.AddCustomHeadingStyle("Heading1", 4, 0);
            styleBuilder.AddCustomHeadingStyle("Heading2", 5, 1);
            styleBuilder.Build();
        }
    }
}
