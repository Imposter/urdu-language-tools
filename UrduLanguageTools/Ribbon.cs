using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using UrduLanguageTools.Extensions;
using Application = Microsoft.Office.Interop.Word.Application;
using Office = Microsoft.Office.Core;

namespace UrduLanguageTools
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private static readonly int[] LinesPerVerses = { 2, 3, 4 };

        private List<Style> styles = new List<Style>();
        
        private Application App => Globals.Addin.Application;
        
        private Office.IRibbonUI ribbon;

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("UrduLanguageTools.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        
        public void Ribbon_Load(Office.IRibbonUI ribbon)
        {
            this.ribbon = ribbon;
            
            // Set callbacks
            Globals.Addin.Application.DocumentOpen += doc => RefreshUI();
            Globals.Addin.Application.DocumentChange += () => RefreshUI();
        }

        public void Refresh_Clicked(Office.IRibbonControl control)
        {
            RefreshUI();
        }
        
        public void RemoveMultipleSpaces_Clicked(Office.IRibbonControl control)
        {
            var selectedText = App.Selection.Text;
            string modifiedText;

            do
            {
                modifiedText = selectedText.Replace("  ", " ");
                if (modifiedText == selectedText)
                    break;
                selectedText = modifiedText;
            } while (true);

            App.Selection.TypeText(modifiedText);
        }
        
        public void GhazalPaste_Clicked(Office.IRibbonControl control)
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

            var lines = Clipboard.GetText().GetLines(CharCode.BraillePatternBlank);
            if (lines.Count == 0)
            {
                MessageBox.Show("The clipboard does not contain any text.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            App.Selection.InsertGhazal(lines, options);
        }
        
        public void GhazalFormat_Clicked(Office.IRibbonControl control)
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

            var lines = App.Selection.Range.GetLines(CharCode.BraillePatternBlank);
            if (lines.Count == 0)
            {
                MessageBox.Show("Please highlight the text you want to update.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            App.Selection.InsertGhazal(lines, options);
        }

        public void GhazalStyle_Changed(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            var style = styles[selectedIndex];
            App.ActiveDocument.SetSetting<AppSettings, string>((s, v) => s.GhazalOptions.ParagraphStyle = v, style.NameLocal);
        }

        public void GhazalLinesPerVerse_Changed(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            var linesPerVerse = LinesPerVerses[selectedIndex];
            App.ActiveDocument.SetSetting<AppSettings, int>((s, v) => s.GhazalOptions.LinesPerVerse = v, linesPerVerse);
        }

        public int GhazalStyle_ItemSource_Count(Office.IRibbonControl control)
        {
            return App.Documents.Count == 0 ? 0 : styles.Count(s => s.Type == WdStyleType.wdStyleTypeParagraph);
        }

        public string GhazalStyle_ItemSource_Label(Office.IRibbonControl control, int index)
        {
            return styles.Where(s => s.Type == WdStyleType.wdStyleTypeParagraph).ToList()[index].NameLocal;
        }
        
        public int GhazalStyle_ItemSource_GetSelectedItemIndex(Office.IRibbonControl control)
        {
            var styleName = App.ActiveDocument.GetSetting<AppSettings, string>(s => s.GhazalOptions.ParagraphStyle);
            return styles.ToList().FindIndex(s => s.NameLocal == styleName);
        }

        public int GhazalLinesPerVerse_ItemSource_Count(Office.IRibbonControl control)
        {
            return App.Documents.Count == 0 ? 0 : LinesPerVerses.Length;
        }
        
        public string GhazalLinesPerVerse_ItemSource_Label(Office.IRibbonControl control, int index)
        {
            return LinesPerVerses[index].ToString();
        }
        
        public int GhazalLinesPerVerse_ItemSource_GetSelectedItemIndex(Office.IRibbonControl control)
        {
            var linesPerVerse = App.ActiveDocument.GetSetting<AppSettings, int>(s => s.GhazalOptions.LinesPerVerse);
            return Array.IndexOf(LinesPerVerses, linesPerVerse);
        }
        
        public void GhazalAddToTableOfContents_Checked(Office.IRibbonControl control, bool isChecked)
        {
            App.ActiveDocument.SetSetting<AppSettings, bool>((s, v) => s.GhazalOptions.AddToTableOfContents = v, isChecked);
        }
        
        public bool GhazalAddToTableOfContents_GetPressed(Office.IRibbonControl control)
        {
            return App.ActiveDocument.GetSetting<AppSettings, bool>(s => s.GhazalOptions.AddToTableOfContents);
        }

        #endregion

        #region Helpers

        private void RefreshUI()
        {
            if (App.Documents.Count == 0)
                return;
            
            styles.Clear();
            styles.AddRange(App.ActiveDocument.Styles.Cast<Style>().Where(s => s.Type == WdStyleType.wdStyleTypeParagraph));
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
