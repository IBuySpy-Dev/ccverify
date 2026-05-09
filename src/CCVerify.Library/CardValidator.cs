namespace CCVerify.Library;

/// <summary>
/// Combines Luhn algorithm validation with card network detection.
/// A card is valid when both the Luhn check passes and the network is recognized.
/// </summary>
public static class CardValidator
{
    /// <summary>
    /// Validates a card number by running Luhn check and network detection.
    /// </summary>
    public static CardValidationResult Validate(string number)
    {
        if (string.IsNullOrWhiteSpace(number))
            return CardValidationResult.Invalid();

        var network = CardNetworkDetector.Detect(number);
        var luhnValid = LuhnAlgorithm.IsValid(number);

        return new CardValidationResult(
            IsLuhnValid: luhnValid,
            Network: network,
            IsValid: luhnValid && network != CardNetwork.Unknown);
    }
}
