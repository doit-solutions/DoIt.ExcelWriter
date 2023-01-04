using System.Data;
using System.Data.Common;

namespace DoIt.ExcelWriter;

internal class DbDataReaderExcelSheetWriter : BaseExcelSheetWriter, IDbDataReaderExcelSheetWriter
{
    public DbDataReaderExcelSheetWriter(Stream destination): base(destination)
    {
    }

    public async Task WriteAllAsync(DbDataReader reader, CancellationToken cancellationToken = default)
    {
        await WriteIntroAsync(GetColumns(reader), cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            await WriteRowAsync(await GetRowAsync(reader, cancellationToken), cancellationToken);
        }
    }

    public async Task WriteAsync(DbDataReader reader, CancellationToken cancellationToken = default)
    {
        await WriteIntroAsync(GetColumns(reader), cancellationToken);
        await WriteRowAsync(await GetRowAsync(reader, cancellationToken), cancellationToken);
    }

    private IEnumerable<ColumnSpec> GetColumns(DbDataReader reader)
    {
        return Enumerable
            .Range(0, reader.FieldCount)
            .Select(idx => new ColumnSpec
            (
                idx + 1,
                reader.GetName(idx),
                reader.GetFieldType(idx),
                null
            ));
    }

    private async Task<IEnumerable<object?>> GetRowAsync(DbDataReader reader, CancellationToken cancellationToken)
    {
        return await Task.WhenAll
        (
            Enumerable
                .Range(0, reader.FieldCount)
                .Select(async idx => reader.GetFieldType(idx) switch
                {
                    var type when type == typeof(byte) => (object?)(await reader.IsDBNullAsync(idx, cancellationToken) ? null : await reader.GetFieldValueAsync<byte>(idx, cancellationToken)),
                    var type when type == typeof(sbyte) => (object?)(await reader.IsDBNullAsync(idx, cancellationToken) ? null : await reader.GetFieldValueAsync<sbyte>(idx, cancellationToken)),
                    var type when type == typeof(short) => (object?)(await reader.IsDBNullAsync(idx, cancellationToken) ? null : await reader.GetFieldValueAsync<short>(idx, cancellationToken)),
                    var type when type == typeof(int) => (object?)(await reader.IsDBNullAsync(idx, cancellationToken) ? null : await reader.GetFieldValueAsync<int>(idx, cancellationToken)),
                    var type when type == typeof(long) => (object?)(await reader.IsDBNullAsync(idx, cancellationToken) ? null : await reader.GetFieldValueAsync<long>(idx, cancellationToken)),
                    var type when type == typeof(float) => (object?)(await reader.IsDBNullAsync(idx, cancellationToken) ? null : await reader.GetFieldValueAsync<float>(idx, cancellationToken)),
                    var type when type == typeof(double) => (object?)(await reader.IsDBNullAsync(idx, cancellationToken) ? null : await reader.GetFieldValueAsync<double>(idx, cancellationToken)),
                    var type when type == typeof(decimal) => (object?)(await reader.IsDBNullAsync(idx, cancellationToken) ? null : await reader.GetFieldValueAsync<decimal>(idx, cancellationToken)),
                    var type when type == typeof(DateTimeOffset) => (object?)(await reader.IsDBNullAsync(idx, cancellationToken) ? null : await reader.GetFieldValueAsync<DateTimeOffset>(idx, cancellationToken)),
                    var type when type == typeof(DateTime) => (object?)(await reader.IsDBNullAsync(idx, cancellationToken) ? null : await reader.GetFieldValueAsync<DateTime>(idx, cancellationToken)),
                    var type when type == typeof(bool) => (object?)(await reader.IsDBNullAsync(idx, cancellationToken) ? null : await reader.GetFieldValueAsync<bool>(idx, cancellationToken)),
                    var type when type == typeof(string) => (object?)(await reader.IsDBNullAsync(idx, cancellationToken) ? null : await reader.GetFieldValueAsync<string>(idx, cancellationToken)),
                    _ => string.Empty
                })
        );
    }
}
