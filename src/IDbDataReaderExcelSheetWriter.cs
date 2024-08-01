using System.Data.Common;

namespace DoIt.ExcelWriter;

public interface IDbDataReaderExcelSheetWriter : IDisposable, IAsyncDisposable
{
    public Task WriteAsync(DbDataReader reader, CancellationToken cancellationToken = default);
    public Task WriteAllAsync(DbDataReader reader, CancellationToken cancellationToken = default);
}
