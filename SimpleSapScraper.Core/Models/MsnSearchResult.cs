namespace SimpleSapScraper.Core.Models
{
    public sealed class MsnSearchResult
    {
        public string Msn { get; set; } = string.Empty;

        public bool Found { get; set; }

        public string Status => Found ? "Found" : "Not Found";

        public string Remarks { get; set; } = string.Empty;
    }
}