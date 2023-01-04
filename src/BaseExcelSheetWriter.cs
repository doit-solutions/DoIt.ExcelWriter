using System.Globalization;
using System.Xml;

namespace DoIt.ExcelWriter;

internal abstract class BaseExcelSheetWriter : IDisposable, IAsyncDisposable
{
    private const string SpreadsheetMlXmlNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    protected record ColumnSpec(int Index, string Name, Type Type, decimal? CustomWidth);

    private readonly XmlWriter _writer;

    private bool _introWritten = false;

    protected BaseExcelSheetWriter(Stream destination)
    {
        _writer = XmlWriterFactory.Create(destination);
    }

    public void Dispose()
    {
        WriteIntroAsync(Array.Empty<ColumnSpec>(), CancellationToken.None).ConfigureAwait(false).GetAwaiter().GetResult();
        WriteOutroAsync().ConfigureAwait(false).GetAwaiter().GetResult();
        _writer.Dispose();
    }

    public async ValueTask DisposeAsync()
    {
        await WriteIntroAsync(Array.Empty<ColumnSpec>(), CancellationToken.None);
        await WriteOutroAsync();
        #if NET6_0_OR_GREATER
        await _writer.DisposeAsync();
        #else
        _writer.Dispose();
        #endif
    }

    protected async Task WriteIntroAsync(IEnumerable<ColumnSpec> columns, CancellationToken cancellationToken)
    {
        if (!_introWritten)
        {
            await _writer.WriteStartDocumentAsync();
            await _writer.WriteStartElementAsync(null, "worksheet", SpreadsheetMlXmlNamespace);

            await _writer.WriteStartElementAsync(null, "cols", SpreadsheetMlXmlNamespace);
            foreach (var column in columns)
            {
                var minWidth = column.Type switch
                {
                    var type when type == typeof(byte) => 8.59m,
                    var type when type == typeof(short) => 8.59m,
                    var type when type == typeof(int) => 8.59m,
                    var type when type == typeof(long) => 8.59m,
                    var type when type == typeof(float) => 11.2m,
                    var type when type == typeof(double) => 11.2m,
                    var type when type == typeof(decimal) => 11.2m,
                    var type when type == typeof(bool) => 7.0m,
                    var type when type == typeof(DateTime) => 14.9m,
                    var type when type == typeof(DateTimeOffset) => 14.9m,
                    _ => 2.4m
                };
                var width = Math.Max(minWidth, column.CustomWidth.HasValue && column.CustomWidth.Value > 0.0m ? column.CustomWidth.Value : column.Name.Count() * 1.2m);
                await _writer.WriteStartElementAsync(null, "col", SpreadsheetMlXmlNamespace);
                await _writer.WriteAttributeStringAsync(null, "min", null, column.Index.ToString(CultureInfo.InvariantCulture));
                await _writer.WriteAttributeStringAsync(null, "max", null, column.Index.ToString(CultureInfo.InvariantCulture));
                await _writer.WriteAttributeStringAsync(null, "width", null, width.ToString(CultureInfo.InvariantCulture));
                await _writer.WriteAttributeStringAsync(null, "style", null, "0");
                await _writer.WriteAttributeStringAsync(null, "customWidth", null, "1");
                await _writer.WriteEndElementAsync(); // col
            }
            await _writer.WriteEndElementAsync(); // cols
            await _writer.WriteStartElementAsync(null, "sheetData", SpreadsheetMlXmlNamespace);
            await _writer.WriteStartElementAsync(null, "row", SpreadsheetMlXmlNamespace);
            await _writer.WriteAttributeStringAsync(null, "s", null, "1");
            await _writer.WriteAttributeStringAsync(null, "customFormat", null, "1");
            foreach (var column in columns)
            {
                await _writer.WriteStartElementAsync(null, "c", SpreadsheetMlXmlNamespace);
                await _writer.WriteAttributeStringAsync(null, "s", null, "1");
                await _writer.WriteAttributeStringAsync(null, "t", null, "inlineStr");
                await _writer.WriteStartElementAsync(null, "is", SpreadsheetMlXmlNamespace);
                await _writer.WriteElementStringAsync(null, "t", SpreadsheetMlXmlNamespace, column.Name);
                await _writer.WriteEndElementAsync(); // is
                await _writer.WriteEndElementAsync(); // c
            }
            await _writer.WriteEndElementAsync(); // row

            _introWritten = true;
        }
    }

    protected async Task WriteOutroAsync()
    {
        await _writer.WriteEndElementAsync(); // sheetData
        await _writer.WriteEndDocumentAsync();
    }

    protected async Task WriteRowAsync(IEnumerable<object?> row, CancellationToken cancellationToken)
    {
        await _writer.WriteStartElementAsync(null, "row", SpreadsheetMlXmlNamespace);
        foreach (var value in row)
        {
            await _writer.WriteStartElementAsync(null, "c", SpreadsheetMlXmlNamespace);
            switch (value)
            {
                case byte val:
                    await WriteIntegerValueAsync(val, cancellationToken);
                    break;
                case short val:
                    await WriteIntegerValueAsync(val, cancellationToken);
                    break;
                case int val:
                    await WriteIntegerValueAsync(val, cancellationToken);
                    break;
                case long val:
                    await WriteIntegerValueAsync(val, cancellationToken);
                    break;
                case float val:
                    await WriteDecimalValueAsync((decimal)val, cancellationToken);
                    break;
                case double val:
                    await WriteDecimalValueAsync((decimal)val, cancellationToken);
                    break;
                case decimal val:
                    await WriteDecimalValueAsync(val, cancellationToken);
                    break;
                case DateTimeOffset val:
                    await WriteDateTimeValueAsync(val.DateTime, cancellationToken);
                    break;
                case DateTime val:
                    await WriteDateTimeValueAsync(val, cancellationToken);
                    break;
                case Uri val:
                    await WriteUriValueAsync(val, cancellationToken);
                    break;
                case Hyperlink val:
                    await WriteHyperlinkValueAsync(val, cancellationToken);
                    break;
                case bool val:
                    await _writer.WriteAttributeStringAsync(null, "s", null, "0");
                    await _writer.WriteAttributeStringAsync(null, "t", null, "b");
                    await _writer.WriteElementStringAsync(null, "v", SpreadsheetMlXmlNamespace, val ? "1" : "0");
                    break;
                case string val:
                    await WriteStringValueAsync(val, cancellationToken);
                    break;
            }
            await _writer.WriteEndElementAsync(); // c
        }
        await _writer.WriteEndElementAsync(); // row
    }

    protected async Task WriteIntegerValueAsync(long val, CancellationToken cancellationToken)
    {
        await _writer.WriteAttributeStringAsync(null, "s", null, "5");
        await _writer.WriteAttributeStringAsync(null, "t", null, "n");
        await _writer.WriteElementStringAsync(null, "v", SpreadsheetMlXmlNamespace, val.ToString(CultureInfo.InvariantCulture));
    }

    protected async Task WriteDecimalValueAsync(decimal val, CancellationToken cancellationToken)
    {
        await _writer.WriteAttributeStringAsync(null, "s", null, "4");
        await _writer.WriteAttributeStringAsync(null, "t", null, "n");
        await _writer.WriteElementStringAsync(null, "v", SpreadsheetMlXmlNamespace, val.ToString(CultureInfo.InvariantCulture));
    }

    protected async Task WriteDateTimeValueAsync(DateTime val, CancellationToken cancellationToken)
    {
        var isDateTime = val.Hour != 0 || val.Minute != 0 || val.Second != 0 || val.Millisecond != 0;
        await _writer.WriteAttributeStringAsync(null, "s", null, isDateTime ? "3" : "2");
        await _writer.WriteAttributeStringAsync(null, "t", null, "d");
        await _writer.WriteElementStringAsync(null, "v", SpreadsheetMlXmlNamespace, val.ToString(isDateTime ? @"yyyy-MM-dd\THH:mm:ss.fff" : "yyyy-MM-dd", CultureInfo.InvariantCulture));
    }

    protected async Task WriteUriValueAsync(Uri val, CancellationToken cancellationToken)
    {
        await _writer.WriteAttributeStringAsync(null, "s", null, "6");
        await _writer.WriteAttributeStringAsync(null, "t", null, "str");
        await _writer.WriteElementStringAsync(null, "f", SpreadsheetMlXmlNamespace, $"HYPERLINK(\"{val}\")");
        await _writer.WriteElementStringAsync(null, "v", SpreadsheetMlXmlNamespace, val.ToString());
    }

    protected async Task WriteHyperlinkValueAsync(Hyperlink val, CancellationToken cancellationToken)
    {
        await _writer.WriteAttributeStringAsync(null, "s", null, "6");
        await _writer.WriteAttributeStringAsync(null, "t", null, "str");
        await _writer.WriteElementStringAsync(null, "f", SpreadsheetMlXmlNamespace, $"HYPERLINK(\"{val.Uri.ToString()}\"{(val.Title != null ? $",\"{val.Title.Replace("\"", "\"\"")}\"" : string.Empty)})");
        await _writer.WriteElementStringAsync(null, "v", SpreadsheetMlXmlNamespace, val.Title ?? val.Uri.ToString());
    }

    protected async Task WriteStringValueAsync(string val, CancellationToken cancellationToken)
    {
        await _writer.WriteAttributeStringAsync(null, "s", null, "0");
        await _writer.WriteAttributeStringAsync(null, "t", null, "inlineStr");
        await _writer.WriteStartElementAsync(null, "is", SpreadsheetMlXmlNamespace);
        await _writer.WriteElementStringAsync(null, "t", SpreadsheetMlXmlNamespace, val);
        await _writer.WriteEndElementAsync(); // is
    }
}
