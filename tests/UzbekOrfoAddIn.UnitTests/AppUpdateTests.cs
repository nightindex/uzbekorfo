using System;
using System.IO;
using UzbekOrfoAddIn.Services;
using Xunit;

namespace UzbekOrfoAddIn.UnitTests;

public class AppUpdateTests
{
    [Theory]
    [InlineData("v1.2.3", "1.2.3.0")]
    [InlineData("1.2.3.4", "1.2.3.4")]
    public void StableVersionIsNormalized(string tag, string expected)
    {
        var release = AppUpdateService.Parse("{\"draft\":false,\"prerelease\":false,\"tag_name\":\"" + tag + "\",\"body\":\"notes\",\"html_url\":\"https://evil.example/\"}");
        Assert.Equal(Version.Parse(expected), release.Version);
        Assert.Equal(AppUpdateService.ReleasesUrl + "/tag/" + tag, release.PageUrl);
        Assert.Equal("notes", release.Notes);
    }

    [Theory]
    [InlineData("v1.2.3-beta")]
    [InlineData("v1.2")]
    [InlineData("../../evil")]
    [InlineData("9999999999999.2.3")]
    public void UnsupportedTagsAreRejected(string tag) => Assert.Throws<InvalidDataException>(() =>
        AppUpdateService.Parse("{\"draft\":false,\"prerelease\":false,\"tag_name\":\"" + tag + "\"}"));

    [Theory]
    [InlineData("true", "false")]
    [InlineData("false", "true")]
    [InlineData("null", "false")]
    public void NonStableMetadataIsRejected(string draft, string prerelease) => Assert.Throws<InvalidDataException>(() =>
        AppUpdateService.Parse("{\"draft\":" + draft + ",\"prerelease\":" + prerelease + ",\"tag_name\":\"v1.2.3\"}"));

    [Fact] public void EquivalentVersionsDoNotOfferUpdate() =>
        Assert.Equal(AppUpdateService.Normalize(new Version(1, 2, 3)), new Version(1, 2, 3, 0));
}
