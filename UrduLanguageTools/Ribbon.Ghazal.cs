using System;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using UrduLanguageTools.Extensions;

namespace UrduLanguageTools
{
    public partial class Ribbon
    {
        private static readonly int[] LinesPerVerses = { 2, 3, 4 };

        #region Ribbon Callbacks
        
        public void GhazalPaste_Clicked(IRibbonControl control)
        {
            var options = App.ActiveDocument.GetSetting<AppSettings, GhazalOptions>(s => s.GhazalOptions);
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

            App.Selection.InsertGhazal(lines, options);
        }
        
        public void GhazalFormat_Clicked(IRibbonControl control)
        {
            var options = App.ActiveDocument.GetSetting<AppSettings, GhazalOptions>(s => s.GhazalOptions);
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

            App.Selection.InsertGhazal(lines, options);
        }

        public void GhazalStyle_Changed(IRibbonControl control, string selectedId, int selectedIndex)
        {
            var style = styles[selectedIndex];
            App.ActiveDocument.SetSetting<AppSettings, string>((s, v) => s.GhazalOptions.ParagraphStyle = v, style.NameLocal);
        }

        public void GhazalLinesPerVerse_Changed(IRibbonControl control, string selectedId, int selectedIndex)
        {
            var linesPerVerse = LinesPerVerses[selectedIndex];
            App.ActiveDocument.SetSetting<AppSettings, int>((s, v) => s.GhazalOptions.LinesPerVerse = v, linesPerVerse);
        }

        public int GhazalStyle_ItemSource_Count(IRibbonControl control)
        {
            return App.Documents.Count == 0 ? 0 : styles.Count(s => s.Type == WdStyleType.wdStyleTypeParagraph);
        }

        public string GhazalStyle_ItemSource_Label(IRibbonControl control, int index)
        {
            return styles.Where(s => s.Type == WdStyleType.wdStyleTypeParagraph).ToList()[index].NameLocal;
        }
        
        public int GhazalStyle_ItemSource_GetSelectedItemIndex(IRibbonControl control)
        {
            var styleName = App.ActiveDocument.GetSetting<AppSettings, string>(s => s.GhazalOptions.ParagraphStyle);
            return styles.ToList().FindIndex(s => s.NameLocal == styleName);
        }

        public int GhazalLinesPerVerse_ItemSource_Count(IRibbonControl control)
        {
            return App.Documents.Count == 0 ? 0 : LinesPerVerses.Length;
        }
        
        public string GhazalLinesPerVerse_ItemSource_Label(IRibbonControl control, int index)
        {
            return LinesPerVerses[index].ToString();
        }
        
        public int GhazalLinesPerVerse_ItemSource_GetSelectedItemIndex(IRibbonControl control)
        {
            var linesPerVerse = App.ActiveDocument.GetSetting<AppSettings, int>(s => s.GhazalOptions.LinesPerVerse);
            return Array.IndexOf(LinesPerVerses, linesPerVerse);
        }
        
        public void GhazalAddToTableOfContents_Checked(IRibbonControl control, bool isChecked)
        {
            App.ActiveDocument.SetSetting<AppSettings, bool>((s, v) => s.GhazalOptions.AddToTableOfContents = v, isChecked);
        }
        
        public bool GhazalAddToTableOfContents_GetPressed(IRibbonControl control)
        {
            return App.ActiveDocument.GetSetting<AppSettings, bool>(s => s.GhazalOptions.AddToTableOfContents);
        }

        #endregion
    }
}
