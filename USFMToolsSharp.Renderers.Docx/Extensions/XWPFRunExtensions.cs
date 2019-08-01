using System;
using System.Collections.Generic;
using System.Text;
using NPOI.XWPF.UserModel;


namespace USFMToolsSharp.Renderers.Docx.Extensions
{
    public static class XWPFRunExtensions
    {
        public static void StyleRun(this XWPFRun run, StyleConfig styles)
        {
            run.IsBold = styles.isBold;
            run.IsItalic = styles.isItalics;
            run.FontSize = styles.fontSize;
            run.IsSmallCaps = styles.isSmallCaps;

        }
    }
}
