namespace DoIt.ExcelWriter;

public static class ExcelWriterFactory
{
    /// <summary>
    ///     Creates an <c>IExcelWriter</c> instance writing to a file with the given file name/path.
    /// </summary>
    /// <param name="fileName">
    ///     The name/full path of the destination file. If a file with the given name already exists,
    ///     it will be truncated and overwritten.
    /// </param>
    /// <returns>
    ///     An <c>IExcelWriter</c> which will write to the provided file.
    /// </returns>
    public static IExcelWriter Create(string fileName)
    {
        return Create(File.Open(fileName, FileMode.Create, FileAccess.Write, FileShare.Read));
    }

    /// <summary>
    ///     Creates an <c>IExcelWriter</c> instance writing to the given stream.
    /// </summary>
    /// <param name="destination">
    ///     The destination stream. The stream must be writable.
    /// </param>
    /// <param name="leaveOpen">
    ///     A boolean which indicated if the <c>ExcelWriter</c> instance should close the provided stream
    ///     when the <c>ExcelWriter</c> is disposed. The default behavior is to close the provided stream.
    /// </param>
    /// <returns>
    ///     An <c>IExcelWriter</c> which will write to the provided stream.
    /// </returns>
    public static IExcelWriter Create(Stream stream, bool leaveOpen = false)
    {
#pragma warning disable CS0618
        return new ExcelWriter(stream, leaveOpen);
#pragma warning restore CS0618
    }
}
