namespace UzbekOrfoAddIn.Models
{
    /// <summary>
    /// Candidate word with precomputed edit distance.
    /// </summary>
    public class CandidateMatch
    {
        public string Word { get; set; }
        public int Distance { get; set; }
    }
}
