using Microsoft.Office.Core;
using UrduLanguageTools.Extensions;

namespace UrduLanguageTools
{
    public partial class Ribbon
    {
        #region Ribbon Callbacks

        public void Refresh_Clicked(IRibbonControl control)
        {
            RefreshUI(App.ActiveDocument);
        }
        
        public void RemoveMultipleSpaces_Clicked(IRibbonControl control)
        {
            var modifiedText = App.Selection.Text.RemoveMultipleSpaces();
            App.Selection.TypeText(modifiedText);
        }

        #endregion
    }
}
