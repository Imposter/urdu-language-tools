using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
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
        
        public void GhazalPaste_Clicked(IRibbonControl control)
        {
            var settings = App.ActiveDocument.GetSettings<AppSettings>();
            if (!App.ActiveDocument.TryGetStyle(settings.GhazalParagraphStyle, out var paragraphStyle))
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

            App.Selection.InsertGhazal(lines, new GhazalOptions
            {
                ParagraphStyle = paragraphStyle,
                AddToTableOfContents = settings.AddToTableOfContents,
                LinesPerVerse = settings.LinesPerVerse,
                ParagraphEnding = settings.ParagraphEnding
            });
        }
        
        public void GhazalFormat_Clicked(IRibbonControl control)
        {
            var settings = App.ActiveDocument.GetSettings<AppSettings>();
            if (!App.ActiveDocument.TryGetStyle(settings.GhazalParagraphStyle, out var paragraphStyle))
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
            
            App.Selection.InsertGhazal(lines, new GhazalOptions
            {
                ParagraphStyle = paragraphStyle,
                AddToTableOfContents = settings.AddToTableOfContents,
                LinesPerVerse = settings.LinesPerVerse,
                ParagraphEnding = settings.ParagraphEnding
            });
        }

        public void NazamPaste_Clicked(IRibbonControl control)
        {
            var settings = App.ActiveDocument.GetSettings<AppSettings>();
            if (!App.ActiveDocument.TryGetStyle(settings.NazamParagraphStyle, out var paragraphStyle))
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
            
            App.Selection.InsertNazam(lines, new NazamOptions
            {
                ParagraphStyle = paragraphStyle,
                AddToTableOfContents = settings.AddToTableOfContents,
                ParagraphEnding = settings.ParagraphEnding
            });
        }

        public void NazamFormat_Clicked(IRibbonControl control)
        {
            var settings = App.ActiveDocument.GetSettings<AppSettings>();
            if (!App.ActiveDocument.TryGetStyle(settings.NazamParagraphStyle, out var paragraphStyle))
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
            
            App.Selection.InsertNazam(lines, new NazamOptions
            {
                ParagraphStyle = paragraphStyle,
                AddToTableOfContents = settings.AddToTableOfContents,
                ParagraphEnding = settings.ParagraphEnding
            });
        }
        
        public void NasarPaste_Clicked(IRibbonControl control)
        {
            var settings = App.ActiveDocument.GetSettings<AppSettings>();
            if (!App.ActiveDocument.TryGetStyle(settings.NasarParagraphStyle, out var paragraphStyle))
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
            
            App.Selection.InsertNasar(lines, new NasarOptions
            {
                ParagraphStyle = paragraphStyle,
                AddToTableOfContents = settings.AddToTableOfContents,
                ParagraphEnding = settings.ParagraphEnding
            });
        }
        
        public void NasarFormat_Clicked(IRibbonControl control)
        {
            var settings = App.ActiveDocument.GetSettings<AppSettings>();
            if (!App.ActiveDocument.TryGetStyle(settings.NasarParagraphStyle, out var paragraphStyle))
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
            
            App.Selection.InsertNasar(lines, new NasarOptions
            {
                ParagraphStyle = paragraphStyle,
                AddToTableOfContents = settings.AddToTableOfContents,
                ParagraphEnding = settings.ParagraphEnding
            });
        }

        #endregion
    }
}
