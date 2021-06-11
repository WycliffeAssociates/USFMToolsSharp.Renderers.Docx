using NPOI.XWPF.UserModel;
using System;
using NPOI.OpenXmlFormats.Wordprocessing;

using System.Collections.Generic;
using System.Text;

namespace USFMToolsSharp.Renderers.Docx.Utils
{
    class StylesBuilder
    {

        public static void AddCustomHeadingStyle(XWPFDocument docxDocument, string name, int headingLevel, int outlineLevel, int pointSize = 12)
        {

            CT_Style ctStyle = new CT_Style();
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

            //CT_RPr rpr = new CT_RPr();
            //rpr.AddNewSz().val = (ulong)pointSize * 2;
            //ctStyle.rPr = rpr;

            XWPFStyle style = new XWPFStyle(ctStyle);

            // is a null op if already defined
            XWPFStyles styles = docxDocument.CreateStyles();

            style.StyleType = (ST_StyleType.paragraph);
            styles.AddStyle(style);

        }

    }
}
