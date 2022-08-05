namespace DoIt.ExcelWriter;

public enum Currency
{
    EUR,
    SEK,
    USD
}

public record MoneyAmount(decimal Amount, Currency Currency);
