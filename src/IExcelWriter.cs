namespace DoIt.ExcelWriter;

public interface IExcelWriter : IAsyncDisposable, IDisposable
{
    /// <summary>
    ///     Adds a typed sheet to the Excel file. The added sheet can only contain rows of the provided generic type.
    ///     Note that any added sheet must bed disposed before another sheet may be added to an Excel writer instance.
    /// </summary>
    /// <param name="name">
    ///     The name of the Excel sheet. If no name is provided (<c>null</c>), the name of the sheet type (excluding the
    ///     namespace name) will be used as the sheet name.
    /// </param>
    /// <param name="cancellationToken">
    ///     A <c>CancellationToken</c> which will be passed on to all underlying asynchronous operations. Using the
    ///     provided cancellation token, it is possible to cancel the adding of the sheet from outside the method.
    /// </param>
    Task<IExcelSheetWriter<T>> AddSheetAsync<T>(string? name = null, CancellationToken cancellationToken = default(CancellationToken));

    /// <summary>
    ///     Adds a sheet to the Excel file. An active <c>DbDataReader</c> can be written to the sheet, wither one row at
    ///     a time or the entire data set at once. Even if the entire data set is written at once, only one row is held
    ///     in memory at any time. Note that any added sheet must bed disposed before another sheet may be added to an
    ///     Excel writer instance.
    /// </summary>
    /// <param name="name">
    ///     The name of the Excel sheet. If no name is provided (<c>null</c>), the sheet will be named "Sheet1" for the
    ///     first sheet in the Excel file, "Sheet2" for the second, etc.
    /// </param>
    /// <param name="cancellationToken">
    ///     A <c>CancellationToken</c> which will be passed on to all underlying asynchronous operations. Using the
    ///     provided cancellation token, it is possible to cancel the adding of the sheet from outside the method.
    /// </param>
    Task<IDbDataReaderExcelSheetWriter> AddDbDataReaderSheetAsync(string? name = null, CancellationToken cancellationToken = default(CancellationToken));
}
