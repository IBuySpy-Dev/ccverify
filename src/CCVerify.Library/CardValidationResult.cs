namespace CCVerify.Library;

public record CardValidationResult(
    bool IsLuhnValid,
    CardNetwork Network,
    bool IsValid)
{
    public static CardValidationResult Invalid(CardNetwork network = CardNetwork.Unknown) =>
        new(false, network, false);
}
