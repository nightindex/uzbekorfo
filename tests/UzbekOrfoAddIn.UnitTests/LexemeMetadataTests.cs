using UzbekOrfoAddIn.Models;
using Xunit;

namespace UzbekOrfoAddIn.UnitTests
{
    public class LexemeMetadataTests
    {
        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void AliasOrderDoesNotEraseReviewedMetadata(bool reviewedFirst)
        {
            var reviewed = new LexemeMetadata { Lemma = "root", PartOfSpeech = "noun", HunspellFlags = "A" };
            var alias = new LexemeMetadata { Lemma = "alias", PartOfSpeech = "unknown", HunspellFlags = "aA" };
            var target = reviewedFirst ? reviewed : alias;
            target.MergeFrom(reviewedFirst ? alias : reviewed);
            Assert.Equal("root", target.Lemma);
            Assert.Equal("noun", target.PartOfSpeech);
            Assert.Equal("Aa", target.HunspellFlags);
        }
    }
}
