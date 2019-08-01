using System;
using System.Collections.Generic;
using System.Text;
using NPOI.XWPF.UserModel;


namespace USFMToolsSharp.Renderers.Docx.Extensions
{
    public static class XWPFParagraphExtensions
    {
        public static void StyleParagraph(this XWPFParagraph para, StyleConfig styles)
        {
            para.Alignment = (styles.isAlignRight ? ParagraphAlignment.RIGHT : ParagraphAlignment.LEFT);
        }
    }
}
