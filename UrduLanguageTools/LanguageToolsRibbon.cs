using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Linq;
using System.Windows.Forms;
using UrduLanguageTools.Extensions;
using Application = Microsoft.Office.Interop.Word.Application;

namespace UrduLanguageTools
{
    public partial class LanguageToolsRibbon
    {
        private Application App => Globals.Addin.Application;

        private void LanguageToolsRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            App.DocumentOpen += App_DocumentOpen;
            App.DocumentChange += App_DocumentChange;
            App.DocumentBeforeSave += App_DocumentBeforeSave;
        }

        private void App_DocumentOpen(Document document)
        {
            UpdateInterface(document);
        }

        private void App_DocumentChange()
        {
            if (App.Documents.Count == 0)
                return;

            UpdateInterface(App.ActiveDocument);
        }

        private void App_DocumentBeforeSave(Document document, ref bool saveAsUI, ref bool cancel)
        {
            UpdateInterface(document);
        }

        private void UpdateInterface(Document document)
        {
            try
            {
                var settings = document.GetSettings(new AppSettings());
                drpGhazalStyle.SetItems(Factory, document.Styles.Cast<Style>().Where(s => s.Type == WdStyleType.wdStyleTypeParagraph), s => s.NameLocal);;
                drpGhazalStyle.SelectedItem = drpGhazalStyle.Items
                    .Cast<RibbonDropDownItem>()
                    .FirstOrDefault(i => i.Label == settings.GhazalOptions.ParagraphStyle) ?? throw new InvalidOperationException("Ghazal paragraph style not found.");

                drpGhazalLinesPerVerse.SetItems(Factory, new int[] { 2, 3, 4 }, i => i.ToString());
                drpGhazalLinesPerVerse.SelectedItem = drpGhazalLinesPerVerse.Items
                    .Cast<RibbonDropDownItem>()
                    .FirstOrDefault(i => (int)i.Tag == settings.GhazalOptions.LinesPerVerse) ?? throw new InvalidOperationException("Ghazal lines per verse not found.");

                cbGhazalAddToTableOfContents.Checked = settings.GhazalOptions.AddToTableOfContents;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void drpGhazalParagraphStyle_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            var style = (Style)drpGhazalStyle.SelectedItem.Tag;
            App.ActiveDocument.SetSetting<AppSettings, string>((s, v) => s.GhazalOptions.ParagraphStyle = v, style.NameLocal);
        }

        private void btnGhazalPaste_Click(object sender, RibbonControlEventArgs e)
        {
            if (!App.ActiveDocument.TryGetStyle(drpGhazalStyle.SelectedItem.Label, out var paragraphStyle))
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

            App.Selection.InsertGhazal(lines, new GhazalOptions
            {
                ParagraphStyle = drpGhazalStyle.SelectedItem.Label,
                LinesPerVerse = (int)drpGhazalLinesPerVerse.SelectedItem.Tag,
                AddToTableOfContents = cbGhazalAddToTableOfContents.Checked
            });
        }

        private void btnGhazalFormat_Click(object sender, RibbonControlEventArgs e)
        {
            if (!App.ActiveDocument.TryGetStyle(drpGhazalStyle.SelectedItem.Label, out var paragraphStyle))
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

            App.Selection.InsertGhazal(lines, new GhazalOptions
            {
                ParagraphStyle = drpGhazalStyle.SelectedItem.Label,
                LinesPerVerse = (int)drpGhazalLinesPerVerse.SelectedItem.Tag,
                AddToTableOfContents = cbGhazalAddToTableOfContents.Checked
            });
        }

        private void btnRemoveMultipleSpaces_Click(object sender, RibbonControlEventArgs e)
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
    }
}