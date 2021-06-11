using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace USFMToolsSharp.Renderers.Docx.Utils
{
    class ResourceUtil
    {
        public static Stream GetResourceStream(string filename, string namespacename = "")
        {
            var assembly = Assembly.GetExecutingAssembly();
            
            if (namespacename == "")
            {
                namespacename = assembly.GetName().Name;
            }
            
            var resourceName = $"{namespacename}.{filename}";
            return assembly.GetManifestResourceStream(resourceName);
        }
    }
}
