namespace DoIt.ExcelWriter;

public interface IExcelSheetWriter<T> : IDisposable, IAsyncDisposable
{
    public Task WriteAsync(T row, CancellationToken cancellationToken = default);
    public Task WriteAsync(IEnumerable<T> rows, CancellationToken cancellationToken = default);
    public Task WriteAsync(IAsyncEnumerable<T> rows, CancellationToken cancellationToken = default);
}
