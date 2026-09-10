using System;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Encapsulates "add currently selected word to dictionary" workflow.
    /// </summary>
    public sealed class AddToDictionaryWorkflowService
    {
        private readonly Func<string, bool, AddWordResult> _addWordToDictionary;

        public AddToDictionaryWorkflowService(Func<string, bool, AddWordResult> addWordToDictionary)
        {
            _addWordToDictionary = addWordToDictionary;
        }

        public AddToDictionaryWorkflowResult Execute()
        {
            if (_addWordToDictionary == null)
                return AddToDictionaryWorkflowResult.ServiceUnavailable();

            string selectedText = DocumentHelper.GetSelectedWord();
            if (string.IsNullOrWhiteSpace(selectedText))
                return AddToDictionaryWorkflowResult.WordNotSelected();

            string word = selectedText.Trim();
            _addWordToDictionary(word, true);
            return AddToDictionaryWorkflowResult.Completed(word);
        }
    }

    public sealed class AddToDictionaryWorkflowResult
    {
        public AddToDictionaryWorkflowStatus Status { get; private set; }
        public string Word { get; private set; }

        public static AddToDictionaryWorkflowResult ServiceUnavailable() =>
            new AddToDictionaryWorkflowResult
            {
                Status = AddToDictionaryWorkflowStatus.ServiceUnavailable
            };

        public static AddToDictionaryWorkflowResult WordNotSelected() =>
            new AddToDictionaryWorkflowResult
            {
                Status = AddToDictionaryWorkflowStatus.WordNotSelected
            };

        public static AddToDictionaryWorkflowResult Completed(string word) =>
            new AddToDictionaryWorkflowResult
            {
                Status = AddToDictionaryWorkflowStatus.Completed,
                Word = word
            };
    }

    public enum AddToDictionaryWorkflowStatus
    {
        ServiceUnavailable = 0,
        WordNotSelected = 1,
        Completed = 2
    }
}
