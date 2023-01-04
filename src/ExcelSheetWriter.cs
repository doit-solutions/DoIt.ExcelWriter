using System.Reflection;

namespace DoIt.ExcelWriter;

internal class ExcelSheetWriter<T> : BaseExcelSheetWriter, IExcelSheetWriter<T>
{
    private const string SpreadsheetMlXmlNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public ExcelSheetWriter(Stream destination): base(destination)
    {
    }

    public async Task WriteAsync(T row, CancellationToken cancellationToken = default)
    {
        await WriteIntroAsync(GetColumns(), cancellationToken);
        await WriteRowAsync(GetRow(row), cancellationToken);
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

    private IEnumerable<ColumnSpec> GetColumns()
    {
        return typeof(T)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => !p.GetCustomAttributes(typeof(ExcelColumnAttribute), inherit: true).OfType<ExcelColumnAttribute>().Any(a => a.Ignore))
            .Select((p, idx) =>
            {
                var attr = p.GetCustomAttributes(typeof(ExcelColumnAttribute), inherit: true).OfType<ExcelColumnAttribute>().FirstOrDefault();
                return new ColumnSpec
                (
                    idx + 1,
                    attr?.Title ?? p.Name,
                    p.PropertyType,
                    attr?.CustomWidth != null ? (decimal)attr.CustomWidth : null
                );
            });
    }

    private IEnumerable<object?> GetRow(T row)
    {
        return typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => !p.GetCustomAttributes(typeof(ExcelColumnAttribute), inherit: true).OfType<ExcelColumnAttribute>().Any(a => a.Ignore))
            .Select(p => p.GetValue(row));
    }
}
