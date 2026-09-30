using System;

namespace UzbekOrfoAddIn.Models
{
    /// <summary>
    /// Local-only aggregate. It deliberately contains no document text,
    /// document name, user identity, or surrounding context.
    /// </summary>
    public sealed class UnknownWordStat
    {
        public string Word { get; set; }
        public int Count { get; set; }
        public DateTime FirstSeenUtc { get; set; }
        public DateTime LastSeenUtc { get; set; }
    }
}
