using System;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using UrduLanguageTools.Extensions;

namespace UrduLanguageTools
{
    public partial class Ribbon
    {
        private static readonly int[] LinesPerVerses = { 2, 3, 4 };

        #region Ribbon Callbacks

        public void GhazalStyle_Changed(IRibbonControl control, string selectedId, int selectedIndex)
        {
            var style = styles[selectedIndex];
            App.ActiveDocument.SetSetting<AppSettings, string>((s, v) => s.GhazalParagraphStyle = v, style.NameLocal);
        }

        public int GhazalStyle_ItemSource_Count(IRibbonControl control)
        {
            return App.Documents.Count == 0 ? 0 : styles.Count(s => s.Type == WdStyleType.wdStyleTypeParagraph);
        }

        public string GhazalStyle_ItemSource_Label(IRibbonControl control, int index)
        {
            return styles.Where(s => s.Type == WdStyleType.wdStyleTypeParagraph).Skip(index).First().NameLocal;
        }
        
        public int GhazalStyle_ItemSource_GetSelectedItemIndex(IRibbonControl control)
        {
            var styleName = App.ActiveDocument.GetSetting<AppSettings, string>(s => s.GhazalParagraphStyle);
            return styles.ToList().FindIndex(s => s.NameLocal == styleName);
        }
        
        public void NazamStyle_Changed(IRibbonControl control, string selectedId, int selectedIndex)
        {
            var style = styles[selectedIndex];
            App.ActiveDocument.SetSetting<AppSettings, string>((s, v) => s.NazamParagraphStyle = v, style.NameLocal);
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
            var styleName = App.ActiveDocument.GetSetting<AppSettings, string>(s => s.NazamParagraphStyle);
            return styles.ToList().FindIndex(s => s.NameLocal == styleName);
        }
        
        public void NasarStyle_Changed(IRibbonControl control, string selectedId, int selectedIndex)
        {
            var style = styles[selectedIndex];
            App.ActiveDocument.SetSetting<AppSettings, string>((s, v) => s.NasarParagraphStyle = v, style.NameLocal);
        }
        
        public int NasarStyle_ItemSource_Count(IRibbonControl control)
        {
            return App.Documents.Count == 0 ? 0 : styles.Count(s => s.Type == WdStyleType.wdStyleTypeParagraph);
        }
        
        public string NasarStyle_ItemSource_Label(IRibbonControl control, int index)
        {
            return styles.Where(s => s.Type == WdStyleType.wdStyleTypeParagraph).Skip(index).First().NameLocal;
        }
        
        public int NasarStyle_ItemSource_GetSelectedItemIndex(IRibbonControl control)
        {
            var styleName = App.ActiveDocument.GetSetting<AppSettings, string>(s => s.NasarParagraphStyle);
            return styles.ToList().FindIndex(s => s.NameLocal == styleName);
        }
        
        public void AddToTableOfContents_Checked(IRibbonControl control, bool isChecked)
        {
            App.ActiveDocument.SetSetting<AppSettings, bool>((s, v) => s.AddToTableOfContents = v, isChecked);
        }
        
        public bool AddToTableOfContents_GetPressed(IRibbonControl control)
        {
            return App.ActiveDocument.GetSetting<AppSettings, bool>(s => s.AddToTableOfContents);
        }

        public void LinesPerVerse_Changed(IRibbonControl control, string selectedId, int selectedIndex)
        {
            var linesPerVerse = LinesPerVerses[selectedIndex];
            App.ActiveDocument.SetSetting<AppSettings, int>((s, v) => s.LinesPerVerse = v, linesPerVerse);
        }

        public int LinesPerVerse_ItemSource_Count(IRibbonControl control)
        {
            return App.Documents.Count == 0 ? 0 : LinesPerVerses.Length;
        }
        
        public string LinesPerVerse_ItemSource_Label(IRibbonControl control, int index)
        {
            return LinesPerVerses[index].ToString();
        }
        
        public int LinesPerVerse_ItemSource_GetSelectedItemIndex(IRibbonControl control)
        {
            var linesPerVerse = App.ActiveDocument.GetSetting<AppSettings, int>(s => s.LinesPerVerse);
            return Array.IndexOf(LinesPerVerses, linesPerVerse);
        }

        #endregion
    }
}
