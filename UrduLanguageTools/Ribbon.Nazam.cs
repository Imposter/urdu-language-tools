using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using UrduLanguageTools.Extensions;

namespace UrduLanguageTools
{
    public partial class Ribbon
    {
        #region Ribbon Callbacks

        public void NazamPaste_Clicked(IRibbonControl control)
        {
            var options = App.ActiveDocument.GetSetting<AppSettings, NazamOptions>(s => s.NazamOptions);
            if (!App.ActiveDocument.TryGetStyle(options.ParagraphStyle, out var paragraphStyle))
            {
                MessageBox.Show("The specified style does not exist in the document.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (paragraphStyle.Type != WdStyleType.wdStyleTypeParagraph)
            {
                MessageBox.Show("The specified style is not a character type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var lines = Clipboard.GetText()
                .RemoveMultipleSpaces()
                .GetLines(CharCode.BraillePatternBlank);
            if (lines.Count == 0)
            {
                MessageBox.Show("The clipboard does not contain any text.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            App.Selection.InsertNazam(lines, options);
        }

        public void NazamFormat_Clicked(IRibbonControl control)
        {
            var options = App.ActiveDocument.GetSetting<AppSettings, NazamOptions>(s => s.NazamOptions);
            if (!App.ActiveDocument.TryGetStyle(options.ParagraphStyle, out var paragraphStyle))
            {
                MessageBox.Show("The specified style does not exist in the document.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (paragraphStyle.Type != WdStyleType.wdStyleTypeParagraph)
            {
                MessageBox.Show("The specified style is not a character type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var lines = App.Selection.Range.GetText()
                .RemoveMultipleSpaces()
                .GetLines(CharCode.BraillePatternBlank);
            if (lines.Count == 0)
            {
                MessageBox.Show("Please highlight the text you want to update.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            App.Selection.InsertNazam(lines, options);
        }
        
        public void NazamStyle_Changed(IRibbonControl control, string selectedId, int selectedIndex)
        {
            var style = styles[selectedIndex];
            App.ActiveDocument.SetSetting<AppSettings, string>((s, v) => s.NazamOptions.ParagraphStyle = v, style.NameLocal);
        }
        
        public int NazamStyle_ItemSource_Count(IRibbonControl control)
        {
            return App.Documents.Count == 0 ? 0 : styles.Count(s => s.Type == WdStyleType.wdStyleTypeParagraph);
        }
        
        public string NazamStyle_ItemSource_Label(IRibbonControl control, int index)
        {
            return styles.Where(s => s.Type == WdStyleType.wdStyleTypeParagraph).Skip(index).First().NameLocal;
        }
        
        public int NazamStyle_ItemSource_GetSelectedItemIndex(IRibbonControl control)
        {
            var styleName = App.ActiveDocument.GetSetting<AppSettings, string>(s => s.NazamOptions.ParagraphStyle);
            return styles.ToList().FindIndex(s => s.NameLocal == styleName);
        }
        
        public void NazamAddToTableOfContents_Checked(IRibbonControl control, bool isChecked)
        {
            App.ActiveDocument.SetSetting<AppSettings, bool>((s, v) => s.NazamOptions.AddToTableOfContents = v, isChecked);
        }
        
        public bool NazamAddToTableOfContents_GetPressed(IRibbonControl control)
        {
            return App.ActiveDocument.GetSetting<AppSettings, bool>(s => s.NazamOptions.AddToTableOfContents);
        }
        
        #endregion
    }
}
