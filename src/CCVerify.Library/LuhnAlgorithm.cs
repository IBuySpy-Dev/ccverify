namespace CCVerify.Library;

/// <summary>Validates a card number using the Luhn algorithm.</summary>
public static class LuhnAlgorithm
{
    /// <summary>Returns true if the number passes the Luhn check.</summary>
    public static bool IsValid(string number)
    {
        if (string.IsNullOrWhiteSpace(number))
            return false;

        var digits = number.Where(char.IsDigit).ToArray();
        if (digits.Length < 13)
            return false;

        int sum = 0;
        bool doubleNext = false;

        for (int i = digits.Length - 1; i >= 0; i--)
        {
            int digit = digits[i] - '0';
            if (doubleNext)
            {
                digit *= 2;
                if (digit > 9)
                    digit -= 9;
            }
            sum += digit;
            doubleNext = !doubleNext;
        }

        return sum % 10 == 0;
    }
}
