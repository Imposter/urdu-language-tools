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

        public void Paste_Clicked(IRibbonControl control)
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

            switch (control.Id)
            {
                case "PasteGhazal":
                    App.Selection.InsertGhazal(lines, new GhazalOptions
                    {
                        ParagraphStyle = paragraphStyle,
                        AddToTableOfContents = settings.AddToTableOfContents,
                        LinesPerVerse = settings.LinesPerVerse,
                        ParagraphEnding = settings.ParagraphEnding
                    });
                    break;
                case "PasteNazam":
                    App.Selection.InsertNazam(lines, new NazamOptions
                    {
                        ParagraphStyle = paragraphStyle,
                        AddToTableOfContents = settings.AddToTableOfContents,
                        ParagraphEnding = settings.ParagraphEnding
                    });
                    break;
                case "PasteNasar":
                    App.Selection.InsertNasar(lines, new NasarOptions
                    {
                        ParagraphStyle = paragraphStyle,
                        AddToTableOfContents = settings.AddToTableOfContents,
                        ParagraphEnding = settings.ParagraphEnding
                    });
                    break;
                default:
                    MessageBox.Show("Unknown paste operation.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }
        }
        
        public void Format_Clicked(IRibbonControl control)
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

            switch (control.Id)
            {
                case "FormatGhazal":
                    App.Selection.InsertGhazal(lines, new GhazalOptions
                    {
                        ParagraphStyle = paragraphStyle,
                        AddToTableOfContents = settings.AddToTableOfContents,
                        LinesPerVerse = settings.LinesPerVerse,
                        ParagraphEnding = settings.ParagraphEnding
                    });
                    break;
                case "FormatNazam":
                    App.Selection.InsertNazam(lines, new NazamOptions
                    {
                        ParagraphStyle = paragraphStyle,
                        AddToTableOfContents = settings.AddToTableOfContents,
                        ParagraphEnding = settings.ParagraphEnding
                    });
                    break;
                case "FormatNasar":
                    App.Selection.InsertNasar(lines, new NasarOptions
                    {
                        ParagraphStyle = paragraphStyle,
                        AddToTableOfContents = settings.AddToTableOfContents,
                        ParagraphEnding = settings.ParagraphEnding
                    });
                    break;
                default:
                    MessageBox.Show("Unknown format operation.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }
        }

        #endregion
    }
}
