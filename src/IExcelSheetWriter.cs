namespace DoIt.ExcelWriter;

public interface IExcelSheetWriter<T> : IDisposable, IAsyncDisposable
{
    public Task WriteAsync(T row, CancellationToken cancellationToken = default(CancellationToken));
    public Task WriteAsync(IEnumerable<T> rows, CancellationToken cancellationToken = default(CancellationToken));
    public Task WriteAsync(IAsyncEnumerable<T> rows, CancellationToken cancellationToken = default(CancellationToken));
}
