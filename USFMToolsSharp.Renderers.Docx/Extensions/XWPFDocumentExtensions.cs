using System;
using System.Collections.Generic;
using System.Text;
using NPOI.XWPF.UserModel;


namespace USFMToolsSharp.Renderers.Docx.Extensions
{
    public static class XWPFDocumentExtensions
    {
        public static XWPFParagraph CreateParagraph(this XWPFDocument doc, StyleConfig styles)
        {
            XWPFParagraph para = doc.CreateParagraph();
            para.Alignment = (styles.isAlignRight ? ParagraphAlignment.RIGHT : ParagraphAlignment.LEFT);
            return para;
        }
    }
}
