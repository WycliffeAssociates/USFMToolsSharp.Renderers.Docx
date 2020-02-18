using System;
using System.Collections.Generic;
using System.Text;
using NPOI.XWPF.UserModel;


namespace USFMToolsSharp.Renderers.Docx.Extensions
{
    public static class XWPFParagraphExtensions
    {
        public static XWPFRun CreateRun(this XWPFParagraph para, StyleConfig styles)
        {
            XWPFRun run = para.CreateRun();
            run.IsBold = styles.isBold;
            run.IsItalic = styles.isItalics;
            run.FontSize = styles.fontSize;
            run.IsSmallCaps = styles.isSmallCaps;
            run.Subscript = (styles.isSuperscript ? VerticalAlign.SUPERSCRIPT: VerticalAlign.BASELINE);
            return run;

        }
    }
}
