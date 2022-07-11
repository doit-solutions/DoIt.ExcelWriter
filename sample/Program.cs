using DoIt.ExcelWriter;

await using (var writer = new ExcelWriter("test.xlsx"))
{
    await using (var sheet = await writer.AddSheetAsync<FirstType>("Sheet1"))
    {
        await sheet.WriteAsync(new FirstType { Id = 1, FirstName = "David", LastName = "Nordvall", Birthday = new DateTime(1980, 11, 12), Old = false, Rating = 1000000.0m, HomePage = new Uri("https://www.internet.com") });
    }
    await using (var sheet = await writer.AddSheetAsync<SecondType>("Sheet2"))
    {
        await sheet.WriteAsync(new SecondType { Id = 1, StreetAddress = "111 Main Street", ZipCode = "11111", City = "Umeå", Country = "Sweden" });
    }
}

record FirstType
{
    public int Id { get; init; }
    [ExcelColumn("First name")]
    public string FirstName { get; init; } = string.Empty;
    [ExcelColumn("Last name")]
    public string LastName { get; init; } = string.Empty;
    public DateTime Birthday { get; init; }
    public bool Old { get; init; }
    public decimal Rating { get; init; }
    public Uri? HomePage { get; init; }
}
record SecondType
{
    [ExcelColumn(Ignore = true)]
    public int Id { get; init; }
    [ExcelColumn(Title = "Street address", CustomWidth = 20.0)]
    public string StreetAddress { get; init; } = string.Empty;
    [ExcelColumn("Zip code")]
    public string ZipCode { get; init; } = string.Empty;
    public string City { get; init; } = string.Empty;
    public string Country { get; init; } = string.Empty;
}
