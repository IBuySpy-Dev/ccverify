using CCVerify.Library;
using Xunit;

namespace CCVerify.Tests;

public class CardNetworkDetectorTests
{
    [Theory]
    [InlineData("4532015112830366", CardNetwork.Visa)]      // Standard Visa (16 digits)
    [InlineData("4111111111111111", CardNetwork.Visa)]      // Classic Visa test number
    [InlineData("4000000000000002", CardNetwork.Visa)]      // Visa (13 digits not tested)
    public void Detect_Visa_Prefix4(string number, CardNetwork expected) =>
        Assert.Equal(expected, CardNetworkDetector.Detect(number));

    [Theory]
    [InlineData("5425233430109903", CardNetwork.MasterCard)]  // MC 51-55 range
    [InlineData("5500005555555559", CardNetwork.MasterCard)]  // MC 55
    [InlineData("2221000000000009", CardNetwork.MasterCard)]  // MC 2221
    [InlineData("2720999999999996", CardNetwork.MasterCard)]  // MC 2720
    public void Detect_MasterCard_Prefixes(string number, CardNetwork expected) =>
        Assert.Equal(expected, CardNetworkDetector.Detect(number));

    [Theory]
    [InlineData("374251018720955", CardNetwork.AmericanExpress)]  // Amex 37
    [InlineData("346631337289576", CardNetwork.AmericanExpress)]  // Amex 34
    public void Detect_AmericanExpress_Prefixes(string number, CardNetwork expected) =>
        Assert.Equal(expected, CardNetworkDetector.Detect(number));

    [Theory]
    [InlineData("6011111111111117", CardNetwork.Discover)]  // Discover 6011
    [InlineData("6500000000000002", CardNetwork.Discover)]  // Discover 65
    [InlineData("6441111111111111", CardNetwork.Discover)]  // Discover 644
    public void Detect_Discover_Prefixes(string number, CardNetwork expected) =>
        Assert.Equal(expected, CardNetworkDetector.Detect(number));

    [Fact]
    public void Detect_Unknown_ReturnsUnknown() =>
        Assert.Equal(CardNetwork.Unknown, CardNetworkDetector.Detect("9999999999999999"));

    [Fact]
    public void Detect_NullInput_ReturnsUnknown() =>
        Assert.Equal(CardNetwork.Unknown, CardNetworkDetector.Detect(null!));

    [Fact]
    public void Detect_TooShort_ReturnsUnknown() =>
        Assert.Equal(CardNetwork.Unknown, CardNetworkDetector.Detect("12345"));
}
