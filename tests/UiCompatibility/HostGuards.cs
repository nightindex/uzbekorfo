using System;

namespace UzbekOrfoAddIn.Helpers
{
    internal static class DocumentHelper
    {
        public static void GoToRange(object range) => throw new InvalidOperationException("Unexpected Word navigation");
    }
    // The layout harness must never invoke a host action or open a modal prompt.
    // All controls/forms are real application sources; only this host boundary is guarded.
    internal static class SafeExecutor
    {
        public static void ShowWarning(string message) => throw new InvalidOperationException(message);
        public static void ShowError(string message) => throw new InvalidOperationException(message);
        public static bool Confirm(string message) => throw new InvalidOperationException(message);
    }
}

// No COM object is created: test error entries have null document ranges.
namespace Microsoft.Office.Interop.Word
{
    public interface Range { }
}
