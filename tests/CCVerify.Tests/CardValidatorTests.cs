using CCVerify.Library;
using Xunit;

namespace CCVerify.Tests;

public class CardValidatorTests
{
    [Theory]
    [InlineData("4532015112830366", true,  CardNetwork.Visa)]           // Valid Visa
    [InlineData("5425233430109903", true,  CardNetwork.MasterCard)]     // Valid MasterCard
    [InlineData("378282246310005",  true,  CardNetwork.AmericanExpress)] // Valid Amex 37 (Stripe)
    [InlineData("6011111111111117", true,  CardNetwork.Discover)]       // Valid Discover
    public void Validate_KnownGoodCards_ReturnsValidWithNetwork(
        string number, bool expectedValid, CardNetwork expectedNetwork)
    {
        var result = CardValidator.Validate(number);
        Assert.Equal(expectedValid, result.IsValid);
        Assert.Equal(expectedNetwork, result.Network);
        Assert.True(result.IsLuhnValid);
    }

    [Fact]
    public void Validate_BadLuhn_IsNotValid()
    {
        // Visa prefix but bad check digit
        var result = CardValidator.Validate("4532015112830367");
        Assert.False(result.IsValid);
        Assert.False(result.IsLuhnValid);
        Assert.Equal(CardNetwork.Visa, result.Network); // Network detected but Luhn fails
    }

    [Fact]
    public void Validate_UnknownNetwork_IsNotValid()
    {
        // Luhn-valid but unknown network prefix
        var result = CardValidator.Validate("9999999999999995");
        Assert.False(result.IsValid);
        Assert.Equal(CardNetwork.Unknown, result.Network);
    }

    [Fact]
    public void Validate_NullInput_IsNotValid()
    {
        var result = CardValidator.Validate(null!);
        Assert.False(result.IsValid);
        Assert.Equal(CardNetwork.Unknown, result.Network);
    }

    [Fact]
    public void Validate_EmptyInput_IsNotValid()
    {
        var result = CardValidator.Validate(string.Empty);
        Assert.False(result.IsValid);
    }
}