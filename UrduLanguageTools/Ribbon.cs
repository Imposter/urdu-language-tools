using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace UrduLanguageTools
{
    [ComVisible(true)]
    public partial class Ribbon : IRibbonExtensibility
    {
        private Application App => Globals.Addin.Application;

        private IRibbonUI ribbon;
        private List<Style> styles = new List<Style>();

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("UrduLanguageTools.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(IRibbonUI ribbon)
        {
            this.ribbon = ribbon;

            // Set callbacks
            App.DocumentOpen += doc => RefreshUI(doc);
            App.DocumentChange += () => { if (App.Documents.Count > 0) RefreshUI(App.ActiveDocument); };
        }

        #endregion

        #region Helpers

        private void RefreshUI(Document document)
        {
            styles.Clear();
            styles.AddRange(document.Styles.Cast<Style>().Where(s => s.Type == WdStyleType.wdStyleTypeParagraph));
            ribbon.Invalidate();
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
