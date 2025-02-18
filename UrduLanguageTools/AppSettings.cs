namespace UrduLanguageTools
{
    public sealed class AppSettings
    {
        public bool AddToTableOfContents { get; set; } = true;

        public bool AddPageBreakAtEnd { get; set; } = true;
        
        public int LinesPerVerse { get; set; } = 2;
        
        public string GhazalParagraphStyle { get; set; } = "Normal";
        
        public string NazamParagraphStyle { get; set; } = "Normal";
        
        public string NasarParagraphStyle { get; set; } = "Normal";
    }
}
