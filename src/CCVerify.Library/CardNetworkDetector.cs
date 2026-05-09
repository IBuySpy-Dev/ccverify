namespace CCVerify.Library;

/// <summary>
/// Detects the card network from the card number BIN prefix.
/// Supports Visa, MasterCard, American Express, and Discover.
/// </summary>
public static class CardNetworkDetector
{
    /// <summary>Detects the card network from the card number.</summary>
    public static CardNetwork Detect(string number)
    {
        if (string.IsNullOrWhiteSpace(number))
            return CardNetwork.Unknown;

        var digits = new string(number.Where(char.IsDigit).ToArray());
        if (digits.Length < 6)
            return CardNetwork.Unknown;

        if (IsVisa(digits)) return CardNetwork.Visa;
        if (IsMasterCard(digits)) return CardNetwork.MasterCard;
        if (IsAmericanExpress(digits)) return CardNetwork.AmericanExpress;
        if (IsDiscover(digits)) return CardNetwork.Discover;

        return CardNetwork.Unknown;
    }

    // Visa: starts with 4
    private static bool IsVisa(string digits) =>
        digits[0] == '4';

    // MasterCard: 51-55 or 2221-2720
    private static bool IsMasterCard(string digits)
    {
        if (digits.Length < 4)
            return false;

        if (int.TryParse(digits[..2], out int prefix2) && prefix2 >= 51 && prefix2 <= 55)
            return true;

        if (int.TryParse(digits[..4], out int prefix4) && prefix4 >= 2221 && prefix4 <= 2720)
            return true;

        return false;
    }

    // American Express: 34 or 37
    private static bool IsAmericanExpress(string digits)
    {
        if (digits.Length < 2)
            return false;

        return int.TryParse(digits[..2], out int prefix) &&
               (prefix == 34 || prefix == 37);
    }

    // Discover: 6011, 622126-622925, 644-649, 65
    private static bool IsDiscover(string digits)
    {
        if (digits.Length < 4)
            return false;

        if (digits.StartsWith("6011"))
            return true;

        if (digits.StartsWith("65"))
            return true;

        if (digits.Length >= 3 && int.TryParse(digits[..3], out int prefix3) && prefix3 >= 644 && prefix3 <= 649)
            return true;

        if (digits.Length >= 6 && int.TryParse(digits[..6], out int prefix6) && prefix6 >= 622126 && prefix6 <= 622925)
            return true;

        return false;
    }
}
