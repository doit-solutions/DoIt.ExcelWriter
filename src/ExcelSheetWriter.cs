using System.Globalization;
using System.Reflection;
using System.Text;
using System.Xml;

namespace DoIt.ExcelWriter;

internal class ExcelSheetWriter<T> : IExcelSheetWriter<T>
{
    private const string SpreadsheetMlXmlNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public readonly XmlWriter _writer;

    private bool _introWritten = false;

    public ExcelSheetWriter(Stream desination)
    {
        _writer = XmlWriter.Create(desination, new XmlWriterSettings { Async = true, CloseOutput = true, Encoding = Encoding.UTF8 });
    }

    public async Task WriteAsync(T row, CancellationToken cancellationToken = default)
    {
        await WriteIntroAsync();

        await _writer.WriteStartElementAsync(null, "row", SpreadsheetMlXmlNamespace);
        foreach (var prop in typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance).Where(p => !p.GetCustomAttributes(typeof(ExcelColumnAttribute), inherit: true).OfType<ExcelColumnAttribute>().Any(a => a.Ignore)))
        {
            await _writer.WriteStartElementAsync(null, "c", SpreadsheetMlXmlNamespace);
            switch (prop.GetValue(row))
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

    public async Task WriteAsync(IEnumerable<T> rows, CancellationToken cancellationToken = default)
    {
        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await WriteAsync(row, cancellationToken);
        }
    }

    public async Task WriteAsync(IAsyncEnumerable<T> rows, CancellationToken cancellationToken = default)
    {
        await foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await WriteAsync(row, cancellationToken);
        }
    }

    public void Dispose()
    {
        WriteIntroAsync().RunSynchronously();
        WriteOutroAsync().RunSynchronously();
        _writer.Dispose();
    }

    public async ValueTask DisposeAsync()
    {
        await WriteIntroAsync();
        await WriteOutroAsync();
        await _writer.DisposeAsync();
    }

    private async Task WriteIntroAsync()
    {
        if (!_introWritten)
        {
            await _writer.WriteStartDocumentAsync();
            await _writer.WriteStartElementAsync(null, "worksheet", SpreadsheetMlXmlNamespace);

            var columns = typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => !p.GetCustomAttributes(typeof(ExcelColumnAttribute), inherit: true).OfType<ExcelColumnAttribute>().Any(a => a.Ignore))
                .Select((p, idx) =>
                {
                    var attr = p.GetCustomAttributes(typeof(ExcelColumnAttribute), inherit: true).OfType<ExcelColumnAttribute>().FirstOrDefault();
                    return
                    (
                        Index: idx + 1,
                        Name: attr?.Title ?? p.Name,
                        Type: p.PropertyType,
                        CustomWidth: attr?.CustomWidth,
                        MinWidth: p.PropertyType switch
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
                        }
                    );
                });
            await _writer.WriteStartElementAsync(null, "cols", SpreadsheetMlXmlNamespace);
            foreach (var column in columns)
            {
                var width = Math.Max(column.MinWidth, column.CustomWidth.HasValue && column.CustomWidth.Value > 0.0 ? (decimal)column.CustomWidth : column.Name.Count() * 1.2m);
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

    private async Task WriteOutroAsync()
    {
        await _writer.WriteEndElementAsync(); // sheetData
        await _writer.WriteEndDocumentAsync();
    }

    private async Task WriteIntegerValueAsync(long val, CancellationToken cancellationToken)
    {
        await _writer.WriteAttributeStringAsync(null, "s", null, "5");
        await _writer.WriteAttributeStringAsync(null, "t", null, "n");
        await _writer.WriteElementStringAsync(null, "v", SpreadsheetMlXmlNamespace, val.ToString(CultureInfo.InvariantCulture));
    }

    private async Task WriteDecimalValueAsync(decimal val, CancellationToken cancellationToken)
    {
        await _writer.WriteAttributeStringAsync(null, "s", null, "4");
        await _writer.WriteAttributeStringAsync(null, "t", null, "n");
        await _writer.WriteElementStringAsync(null, "v", SpreadsheetMlXmlNamespace, val.ToString(CultureInfo.InvariantCulture));
    }

    private async Task WriteDateTimeValueAsync(DateTime val, CancellationToken cancellationToken)
    {
        var isDateTime = val.Hour != 0 || val.Minute != 0 || val.Second != 0 || val.Millisecond != 0;
        await _writer.WriteAttributeStringAsync(null, "s", null, isDateTime ? "3" : "2");
        await _writer.WriteAttributeStringAsync(null, "t", null, "d");
        await _writer.WriteElementStringAsync(null, "v", SpreadsheetMlXmlNamespace, val.ToString(isDateTime ? @"yyyy-MM-dd\THH:mm:ss.fff" : "yyyy-MM-dd", CultureInfo.InvariantCulture));
    }

    private async Task WriteUriValueAsync(Uri val, CancellationToken cancellationToken)
    {
        await _writer.WriteAttributeStringAsync(null, "s", null, "0");
        await _writer.WriteAttributeStringAsync(null, "t", null, "str");
        await _writer.WriteElementStringAsync(null, "f", SpreadsheetMlXmlNamespace, $"HYPERLINK(\"{val}\")");
        await _writer.WriteElementStringAsync(null, "v", SpreadsheetMlXmlNamespace, val.ToString());
    }

    private async Task WriteHyperlinkValueAsync(Hyperlink val, CancellationToken cancellationToken)
    {
        await _writer.WriteAttributeStringAsync(null, "s", null, "0");
        await _writer.WriteAttributeStringAsync(null, "t", null, "str");
        await _writer.WriteElementStringAsync(null, "f", SpreadsheetMlXmlNamespace, $"HYPERLINK(\"{val.Uri.ToString()}\"{(val.Title != null ? $",\"{val.Title}\"" : string.Empty)})");
        await _writer.WriteElementStringAsync(null, "v", SpreadsheetMlXmlNamespace, val.Title ?? val.Uri.ToString());
    }

    private async Task WriteStringValueAsync(string val, CancellationToken cancellationToken)
    {
        await _writer.WriteAttributeStringAsync(null, "s", null, "0");
        await _writer.WriteAttributeStringAsync(null, "t", null, "inlineStr");
        await _writer.WriteStartElementAsync(null, "is", SpreadsheetMlXmlNamespace);
        await _writer.WriteElementStringAsync(null, "t", SpreadsheetMlXmlNamespace, val);
        await _writer.WriteEndElementAsync(); // is
    }
}
