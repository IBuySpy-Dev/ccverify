using CCVerify.Library;
using Xunit;

namespace CCVerify.Tests;

public class LuhnAlgorithmTests
{
    [Theory]
    [InlineData("4532015112830366", true)]   // Valid Visa
    [InlineData("5425233430109903", true)]   // Valid MasterCard
    [InlineData("374251018720955", true)]    // Valid Amex
    [InlineData("6011111111111117", true)]   // Valid Discover
    [InlineData("1234567890123456", false)]  // Invalid
    [InlineData("4532015112830367", false)]  // Visa with bad check digit
    public void IsValid_ReturnsExpectedResult(string number, bool expected)
    {
        Assert.Equal(expected, LuhnAlgorithm.IsValid(number));
    }

    [Fact]
    public void IsValid_NullInput_ReturnsFalse() =>
        Assert.False(LuhnAlgorithm.IsValid(null!));

    [Fact]
    public void IsValid_EmptyInput_ReturnsFalse() =>
        Assert.False(LuhnAlgorithm.IsValid(string.Empty));

    [Fact]
    public void IsValid_TooShort_ReturnsFalse() =>
        Assert.False(LuhnAlgorithm.IsValid("123456"));
}
