using System;
using System.Collections.Generic;
using System.Text;
using NPOI.XWPF.UserModel;


namespace USFMToolsSharp.Renderers.Docx.Extensions
{
    public static class XWPFTableCellExtensions
    {
        public static XWPFParagraph AddParagraph(this XWPFTableCell cell, StyleConfig styles)
        {
            XWPFParagraph para = cell.AddParagraph();
            para.Alignment = styles.Alignment;
            return para;

        }
    }
}
