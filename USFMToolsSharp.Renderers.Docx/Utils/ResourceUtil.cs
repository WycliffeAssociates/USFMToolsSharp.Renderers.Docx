using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace USFMToolsSharp.Renderers.Docx.Utils
{
    class ResourceUtil
    {
        public static Stream GetResourceStream(string filename)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "USFMToolsSharp.Renderers.Docx." + filename;
            return assembly.GetManifestResourceStream(resourceName);
        }

        public static Stream GetResourceStream(string filename, string namespacename)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = $"{namespacename}.{filename}";
            return assembly.GetManifestResourceStream(resourceName);
        }
    }
}
