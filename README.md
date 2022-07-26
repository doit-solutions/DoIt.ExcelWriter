# DoIt.ExcelWriter
[![NuGet Badge](https://buildstats.info/nuget/DoIt.ExcelWriter)](https://www.nuget.org/packages/DoIt.ExcelWriter/)

A "forward only" Excel writer.

## Why should I use this?
If you need to create Excel files based on large data sets in a fast and memory efficient manner, this is for you! This library allows you to write Excel data and stream the resulting Excel file as each row is written. This basically means that an ASP.NET application can stream the results of a database query, for example, directly to a client only holding a single result row in memory at any time.

## Sound great, how do I use it?
First, add the library to you project.

```
dotnet add package DoIt.ExcelWriter
```

Then create an `ExcelWriter` instance, add one (or more) typed sheets to it and write rows to the sheet.

```c#
using DoIt.ExcelWriter;

// Create an ExcelWriter and either provide a filename or a Stream instance as destination.
await using (var writer = new DoIt.ExcelWriter.ExcelWriter("test.xlsx"))
// Add a sheet. Note that the sheet is typed and only accepts rows of the specified type!
await using (var sheet = await writer.AddSheetAsync<MyDataType>("Sheet1"))
{
    // Each call to WriteAsync will write all public properties as a single row.
    await sheet.WriteAsync(new MyDataType { ... });
}
```

You can control the apperance of the produced Excel file by using the `ExcelColumnAttribute` attribute on your data type's public properties. This attribute allows you to

 * change the property's column title from the default value (the property name),
 * exclude (i.e ignore) a property,
 * set a custom width of a property's column.

```c#
public record MyDataType
{
    [ExcelColumn(Ignore = true)] // Exclude/ignore the column when writing the Excel data.
    public int Id { get; init; }

    [ExcelColumn("First name")] // Change the default column title.
    public string FirstName { get; init; } = string.Empty;

    [ExcelColumn(CustomWidth = 64)] // Set a custom width of the column.
    public string? Comment { get; init; }
}
```

The library handles properties of the following types:

 * Integers (`byte`, `short`, `int` and `long`)
 * Floating points (`float` and `double`)
 * `decimal`
 * `System.DateTime` and `System.DateTimeOffset`
 * `System.Uri` and `DoIt.ExcelWriter.Hyperlink` (becomes clickable links)
 * `bool`
 * `string`

Values of properties of other types are ignored.

Note that the API only has async methods and accepts `CancellationToken`s whenever possible.

## Fantastic! So what's the catch?
Since the library streams Excel data as each row is written, it is not possible to make changes to data already written. Since column definitions (like the width of a column) comes before the actual data in an Excel file, it is, for example, not possible to change the column width based on the actual data. The library does, however, set sensible default column widths based on each column's title and data type.
