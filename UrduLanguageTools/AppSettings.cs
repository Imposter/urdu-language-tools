namespace UrduLanguageTools
{
    public enum ParagraphEnding
    {
        None,
        Page,
        Section
    }
    
    public sealed class AppSettings
    {
        public bool AddToTableOfContents { get; set; } = true;

        public ParagraphEnding ParagraphEnding { get; set; } = ParagraphEnding.Page;
        
        public int LinesPerVerse { get; set; } = 2;
        
        public string GhazalParagraphStyle { get; set; } = "Normal";
        
        public string NazamParagraphStyle { get; set; } = "Normal";
        
        public string NasarParagraphStyle { get; set; } = "Normal";
    }
}
