using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using UzbekOrfoAddIn.Helpers;

namespace UzbekOrfoAddIn.Prediction
{
    /// <summary>Opt-in aggregate counters only. Never accepts document text or identifiers.</summary>
    public sealed class PredictionMetrics
    {
        private readonly string _path;
        private readonly Dictionary<string, long> _counts = new Dictionary<string, long>();
        public bool Enabled { get; set; }
        public PredictionMetrics(string root) { _path = Path.Combine(root, "metrics-v1.tsv"); }
        public void Count(string name)
        {
            if (!Enabled) return;
            long value; _counts.TryGetValue(name, out value); _counts[name] = value + 1;
        }
        public void Latency(long milliseconds)
        {
            if (!Enabled) return;
            Count(milliseconds <= 50 ? "latency_le_50ms" : milliseconds <= 100 ? "latency_le_100ms" :
                milliseconds <= 150 ? "latency_le_150ms" : "latency_gt_150ms");
        }
        public void KeyToGhostLatency(long milliseconds)
        {
            if (!Enabled) return;
            Count(milliseconds <= 50 ? "key_to_ghost_le_50ms" : milliseconds <= 100 ? "key_to_ghost_le_100ms" :
                milliseconds <= 150 ? "key_to_ghost_le_150ms" : "key_to_ghost_gt_150ms");
        }
        public void Save()
        {
            if (!Enabled) return;
            Directory.CreateDirectory(Path.GetDirectoryName(_path));
            AtomicFile.WriteAllLines(_path, new[] { "matnai-metrics-v1-current-session" }.Concat(
                _counts.OrderBy(p => p.Key).Select(p => p.Key + "\t" + p.Value.ToString(CultureInfo.InvariantCulture))).ToArray());
        }
    }
}
