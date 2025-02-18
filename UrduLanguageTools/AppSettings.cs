namespace UrduLanguageTools
{
    public sealed class GhazalOptions
    {
        public string ParagraphStyle { get; set; } = "Normal";

        public bool AddToTableOfContents { get; set; } = true;

        public int LinesPerVerse { get; set; } = 2;
    }
    
    public sealed class AppSettings
    {
        public GhazalOptions GhazalOptions { get; set; } = new GhazalOptions();
    }
}
