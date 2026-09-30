using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Core
{
    /// <summary>
    /// Provides optional, reviewed lexical metadata for dictionary words.
    /// Missing metadata must remain unknown rather than being guessed.
    /// </summary>
    public interface ILexemeMetadataProvider
    {
        bool TryGetLexeme(string word, out LexemeMetadata lexeme);
    }
}
