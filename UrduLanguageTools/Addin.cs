using Microsoft.Office.Tools.Ribbon;

namespace UrduLanguageTools
{
    public partial class Addin
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {            
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override IRibbonExtension[] CreateRibbonObjects()
        {
            return new[] { new LanguageToolsRibbon() };
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
