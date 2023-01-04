using System.Globalization;
using SharpCompress.Writers.Zip;

namespace DoIt.ExcelWriter;

public class ExcelWriter : IExcelWriter
{
    private const string RelationshipsXmlNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string SpreadsheetMlXmlNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string ContentTypesXmlNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
    
    private readonly ZipWriter _zip;
    private readonly IDictionary<int, string> _sheets = new Dictionary<int, string>();

    private bool _introWritten = false;

    [Obsolete("This constructor will be removed in a future version. Please use ExcelWriterFactory.Create() instead.")]
    public ExcelWriter(string fileName): this(File.Open(fileName, FileMode.Create, FileAccess.Write, FileShare.Read))
    {
    }
    
    [Obsolete("This constructor will be removed in a future version. Please use ExcelWriterFactory.Create() instead.")]
    public ExcelWriter(Stream destination, bool leaveOpen = false)
    {
        _zip = new ZipWriter(destination, new ZipWriterOptions(SharpCompress.Common.CompressionType.Deflate) { LeaveStreamOpen = leaveOpen });
    }

    public async Task<IExcelSheetWriter<T>> AddSheetAsync<T>(string? name = null, CancellationToken cancellationToken = default(CancellationToken))
    {
        await WriteIntroAsync(cancellationToken);

        _sheets.Add(new KeyValuePair<int, string>(_sheets.Count() + 1, name ?? typeof(T).Name));
        return new ExcelSheetWriter<T>(_zip.WriteToStream($"/xl/worksheets/sheet{_sheets.Count()}.xml", new ZipWriterEntryOptions {}));
    }

    public async Task<IDbDataReaderExcelSheetWriter> AddDbDataReaderSheetAsync(string? name = null, CancellationToken cancellationToken = default(CancellationToken))
    {
        await WriteIntroAsync(cancellationToken);

        var idx = _sheets.Count() + 1;
        _sheets.Add(new KeyValuePair<int, string>(idx, name ?? $"Sheet{idx.ToString(CultureInfo.InvariantCulture)}"));
        return new DbDataReaderExcelSheetWriter(_zip.WriteToStream($"/xl/worksheets/sheet{_sheets.Count()}.xml", new ZipWriterEntryOptions {}));
    }

    public void Dispose()
    {
        WriteOutroAsync().ConfigureAwait(false).GetAwaiter().GetResult();
        _zip.Dispose();
    }

    public async ValueTask DisposeAsync()
    {
        await WriteOutroAsync();
        _zip.Dispose();
    }

    private async Task WriteIntroAsync(CancellationToken cancellationToken)
    {
        if (!_introWritten)
        {
            // Write the main relationships XML.
            using (var writer = XmlWriterFactory.Create(_zip.WriteToStream("/_rels/.rels", new ZipWriterEntryOptions { })))
            {
                await writer.WriteStartDocumentAsync();
                await writer.WriteStartElementAsync(null, "Relationships", RelationshipsXmlNamespace);
                await writer.WriteStartElementAsync(null, "Relationship", RelationshipsXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "Type", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
                await writer.WriteAttributeStringAsync(null, "Target", null, "/xl/workbook.xml");
                await writer.WriteAttributeStringAsync(null, "Id", null, "rIdWorkbook1");
                await writer.WriteEndElementAsync(); // Relationship
                await writer.WriteEndElementAsync(); // Relationships
                await writer.WriteEndDocumentAsync();
            }
            // Write the stylesheet XML.
            using (var writer = XmlWriterFactory.Create(_zip.WriteToStream("/xl/styles.xml", new ZipWriterEntryOptions {})))
            {
                await writer.WriteStartDocumentAsync();
                await writer.WriteStartElementAsync(null, "styleSheet", SpreadsheetMlXmlNamespace);

                await writer.WriteStartElementAsync(null, "numFmts", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "count", null, "1");
                await writer.WriteStartElementAsync(null, "numFmt", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "numFmtId", null, "0");
                await writer.WriteAttributeStringAsync(null, "formatCode", null, string.Empty);
                await writer.WriteEndElementAsync(); // numFmt
                await writer.WriteEndElementAsync(); // numFmts;

                await writer.WriteStartElementAsync(null, "fonts", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "count", null, "2");
                await writer.WriteStartElementAsync(null, "font", SpreadsheetMlXmlNamespace);
                await writer.WriteStartElementAsync(null, "vertAlign", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "val", null, "baseline");
                await writer.WriteEndElementAsync(); // vertAlign
                await writer.WriteStartElementAsync(null, "sz", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "val", null, "11");
                await writer.WriteEndElementAsync(); // sz
                await writer.WriteStartElementAsync(null, "color", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "theme", null, "1");
                await writer.WriteEndElementAsync(); // color
                await writer.WriteStartElementAsync(null, "name", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "val", null, "Calibri");
                await writer.WriteEndElementAsync(); // name
                await writer.WriteStartElementAsync(null, "family", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "val", null, "2");
                await writer.WriteEndElementAsync(); // family
                await writer.WriteEndElementAsync(); // font
                await writer.WriteStartElementAsync(null, "font", SpreadsheetMlXmlNamespace);
                await writer.WriteStartElementAsync(null, "b", SpreadsheetMlXmlNamespace);
                await writer.WriteEndElementAsync(); // b
                await writer.WriteStartElementAsync(null, "vertAlign", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "val", null, "baseline");
                await writer.WriteEndElementAsync(); // vertAlign
                await writer.WriteStartElementAsync(null, "sz", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "val", null, "11");
                await writer.WriteEndElementAsync(); // sz
                await writer.WriteStartElementAsync(null, "color", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "theme", null, "1");
                await writer.WriteEndElementAsync(); // color
                await writer.WriteStartElementAsync(null, "name", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "val", null, "Calibri");
                await writer.WriteEndElementAsync(); // name
                await writer.WriteStartElementAsync(null, "family", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "val", null, "2");
                await writer.WriteEndElementAsync(); // family
                await writer.WriteEndElementAsync(); // font
                await writer.WriteStartElementAsync(null, "font", SpreadsheetMlXmlNamespace);
                await writer.WriteStartElementAsync(null, "u", SpreadsheetMlXmlNamespace);
                await writer.WriteEndElementAsync(); // b
                await writer.WriteStartElementAsync(null, "vertAlign", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "val", null, "baseline");
                await writer.WriteEndElementAsync(); // vertAlign
                await writer.WriteStartElementAsync(null, "sz", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "val", null, "11");
                await writer.WriteEndElementAsync(); // sz
                await writer.WriteStartElementAsync(null, "color", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "theme", null, "10");
                await writer.WriteEndElementAsync(); // color
                await writer.WriteStartElementAsync(null, "name", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "val", null, "Calibri");
                await writer.WriteEndElementAsync(); // name
                await writer.WriteStartElementAsync(null, "family", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "val", null, "2");
                await writer.WriteEndElementAsync(); // family
                await writer.WriteEndElementAsync(); // font
                await writer.WriteEndElementAsync(); // fonts

                await writer.WriteStartElementAsync(null, "fills", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "count", null, "2");
                await writer.WriteStartElementAsync(null, "fill", SpreadsheetMlXmlNamespace);
                await writer.WriteStartElementAsync(null, "patternFill", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "patternType", null, "none");
                await writer.WriteEndElementAsync(); // patternFill
                await writer.WriteEndElementAsync(); // fill
                await writer.WriteStartElementAsync(null, "fill", SpreadsheetMlXmlNamespace);
                await writer.WriteStartElementAsync(null, "patternFill", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "patternType", null, "gray125");
                await writer.WriteEndElementAsync(); // patternFill
                await writer.WriteEndElementAsync(); // fill
                await writer.WriteEndElementAsync(); // fills
                
                await writer.WriteStartElementAsync(null, "borders", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "count", null, "2");
                await writer.WriteStartElementAsync(null, "border", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "diagonalUp", null, "0");
                await writer.WriteAttributeStringAsync(null, "diagonalDown", null, "0");
                await writer.WriteStartElementAsync(null, "left", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "style", null, "none");
                await writer.WriteStartElementAsync(null, "color", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
                await writer.WriteEndElementAsync(); // color
                await writer.WriteEndElementAsync(); // left
                await writer.WriteStartElementAsync(null, "right", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "style", null, "none");
                await writer.WriteStartElementAsync(null, "color", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
                await writer.WriteEndElementAsync(); // color
                await writer.WriteEndElementAsync(); // right
                await writer.WriteStartElementAsync(null, "top", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "style", null, "none");
                await writer.WriteStartElementAsync(null, "color", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
                await writer.WriteEndElementAsync(); // color
                await writer.WriteEndElementAsync(); // top
                await writer.WriteStartElementAsync(null, "bottom", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "style", null, "none");
                await writer.WriteStartElementAsync(null, "color", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
                await writer.WriteEndElementAsync(); // color
                await writer.WriteEndElementAsync(); // bottom
                await writer.WriteStartElementAsync(null, "diagonal", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "style", null, "none");
                await writer.WriteStartElementAsync(null, "color", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
                await writer.WriteEndElementAsync(); // color
                await writer.WriteEndElementAsync(); // diagonal
                await writer.WriteEndElementAsync(); // border
                await writer.WriteStartElementAsync(null, "border", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "diagonalUp", null, "0");
                await writer.WriteAttributeStringAsync(null, "diagonalDown", null, "0");
                await writer.WriteStartElementAsync(null, "left", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "style", null, "none");
                await writer.WriteStartElementAsync(null, "color", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
                await writer.WriteEndElementAsync(); // color
                await writer.WriteEndElementAsync(); // left
                await writer.WriteStartElementAsync(null, "right", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "style", null, "none");
                await writer.WriteStartElementAsync(null, "color", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
                await writer.WriteEndElementAsync(); // color
                await writer.WriteEndElementAsync(); // right
                await writer.WriteStartElementAsync(null, "top", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "style", null, "none");
                await writer.WriteStartElementAsync(null, "color", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
                await writer.WriteEndElementAsync(); // color
                await writer.WriteEndElementAsync(); // top
                await writer.WriteStartElementAsync(null, "bottom", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "style", null, "thin");
                await writer.WriteStartElementAsync(null, "color", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
                await writer.WriteEndElementAsync(); // color
                await writer.WriteEndElementAsync(); // bottom
                await writer.WriteStartElementAsync(null, "diagonal", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "style", null, "none");
                await writer.WriteStartElementAsync(null, "color", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
                await writer.WriteEndElementAsync(); // color
                await writer.WriteEndElementAsync(); // diagonal
                await writer.WriteEndElementAsync(); // border
                await writer.WriteEndElementAsync(); // borders

                await writer.WriteStartElementAsync(null, "cellStyleXfs", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "count", null, "7");
                await writer.WriteStartElementAsync(null, "xf", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "numFmtId", null, "0");
                await writer.WriteAttributeStringAsync(null, "fontId", null, "0");
                await writer.WriteAttributeStringAsync(null, "fillId", null, "0");
                await writer.WriteAttributeStringAsync(null, "borderId", null, "0");
                await writer.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyFill", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyBorder", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyProtection", null, "1");
                await writer.WriteStartElementAsync(null, "protection", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "locked", null, "1");
                await writer.WriteAttributeStringAsync(null, "hidden", null, "0");
                await writer.WriteEndElementAsync(); // protection
                await writer.WriteEndElementAsync(); // xf
                // Header style
                await writer.WriteStartElementAsync(null, "xf", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "numFmtId", null, "0");
                await writer.WriteAttributeStringAsync(null, "fontId", null, "1");
                await writer.WriteAttributeStringAsync(null, "fillId", null, "0");
                await writer.WriteAttributeStringAsync(null, "borderId", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyFill", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyBorder", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyProtection", null, "1");
                await writer.WriteStartElementAsync(null, "protection", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "locked", null, "1");
                await writer.WriteAttributeStringAsync(null, "hidden", null, "0");
                await writer.WriteEndElementAsync(); // protection
                await writer.WriteEndElementAsync(); // xf
                // Date style
                await writer.WriteStartElementAsync(null, "xf", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "numFmtId", null, "14");
                await writer.WriteAttributeStringAsync(null, "fontId", null, "0");
                await writer.WriteAttributeStringAsync(null, "fillId", null, "0");
                await writer.WriteAttributeStringAsync(null, "borderId", null, "0");
                await writer.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyFill", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyBorder", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyProtection", null, "1");
                await writer.WriteStartElementAsync(null, "protection", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "locked", null, "1");
                await writer.WriteAttributeStringAsync(null, "hidden", null, "0");
                await writer.WriteEndElementAsync(); // protection
                await writer.WriteEndElementAsync(); // xf
                // Date time style
                await writer.WriteStartElementAsync(null, "xf", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "numFmtId", null, "22");
                await writer.WriteAttributeStringAsync(null, "fontId", null, "0");
                await writer.WriteAttributeStringAsync(null, "fillId", null, "0");
                await writer.WriteAttributeStringAsync(null, "borderId", null, "0");
                await writer.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyFill", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyBorder", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyProtection", null, "1");
                await writer.WriteStartElementAsync(null, "protection", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "locked", null, "1");
                await writer.WriteAttributeStringAsync(null, "hidden", null, "0");
                await writer.WriteEndElementAsync(); // protection
                await writer.WriteEndElementAsync(); // xf
                // Decimal style
                await writer.WriteStartElementAsync(null, "xf", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "numFmtId", null, "4");
                await writer.WriteAttributeStringAsync(null, "fontId", null, "0");
                await writer.WriteAttributeStringAsync(null, "fillId", null, "0");
                await writer.WriteAttributeStringAsync(null, "borderId", null, "0");
                await writer.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyFill", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyBorder", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyProtection", null, "1");
                await writer.WriteStartElementAsync(null, "protection", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "locked", null, "1");
                await writer.WriteAttributeStringAsync(null, "hidden", null, "0");
                await writer.WriteEndElementAsync(); // protection
                await writer.WriteEndElementAsync(); // xf
                // Integer style
                await writer.WriteStartElementAsync(null, "xf", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "numFmtId", null, "3");
                await writer.WriteAttributeStringAsync(null, "fontId", null, "0");
                await writer.WriteAttributeStringAsync(null, "fillId", null, "0");
                await writer.WriteAttributeStringAsync(null, "borderId", null, "0");
                await writer.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyFill", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyBorder", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyProtection", null, "1");
                await writer.WriteStartElementAsync(null, "protection", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "locked", null, "1");
                await writer.WriteAttributeStringAsync(null, "hidden", null, "0");
                await writer.WriteEndElementAsync(); // protection
                await writer.WriteEndElementAsync(); // xf
                // Hyperlink style
                await writer.WriteStartElementAsync(null, "xf", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "numFmtId", null, "0");
                await writer.WriteAttributeStringAsync(null, "fontId", null, "2");
                await writer.WriteAttributeStringAsync(null, "fillId", null, "0");
                await writer.WriteAttributeStringAsync(null, "borderId", null, "0");
                await writer.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyFill", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyBorder", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyProtection", null, "1");
                await writer.WriteStartElementAsync(null, "protection", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "locked", null, "1");
                await writer.WriteAttributeStringAsync(null, "hidden", null, "0");
                await writer.WriteEndElementAsync(); // protection
                await writer.WriteEndElementAsync(); // xf
                await writer.WriteEndElementAsync(); // cellStyleXfs

                await writer.WriteStartElementAsync(null, "cellXfs", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "count", null, "7");
                await writer.WriteStartElementAsync(null, "xf", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "numFmtId", null, "0");
                await writer.WriteAttributeStringAsync(null, "fontId", null, "0");
                await writer.WriteAttributeStringAsync(null, "fillId", null, "0");
                await writer.WriteAttributeStringAsync(null, "borderId", null, "0");
                await writer.WriteAttributeStringAsync(null, "xfId", null, "0");
                await writer.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyFill", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyBorder", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyProtection", null, "1");
                await writer.WriteStartElementAsync(null, "alignment", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "horizontal", null, "general");
                await writer.WriteAttributeStringAsync(null, "vertical", null, "bottom");
                await writer.WriteAttributeStringAsync(null, "textRotation", null, "0");
                await writer.WriteAttributeStringAsync(null, "wrapText", null, "0");
                await writer.WriteAttributeStringAsync(null, "indent", null, "0");
                await writer.WriteAttributeStringAsync(null, "relativeIndent", null, "0");
                await writer.WriteAttributeStringAsync(null, "justifyLastLine", null, "0");
                await writer.WriteAttributeStringAsync(null, "shrinkToFit", null, "0");
                await writer.WriteAttributeStringAsync(null, "readingOrder", null, "0");
                await writer.WriteEndElementAsync(); // alignment
                await writer.WriteStartElementAsync(null, "protection", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "locked", null, "1");
                await writer.WriteAttributeStringAsync(null, "hidden", null, "0");
                await writer.WriteEndElementAsync(); // protection
                await writer.WriteEndElementAsync(); // xf
                // Header
                await writer.WriteStartElementAsync(null, "xf", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "numFmtId", null, "0");
                await writer.WriteAttributeStringAsync(null, "fontId", null, "1");
                await writer.WriteAttributeStringAsync(null, "fillId", null, "0");
                await writer.WriteAttributeStringAsync(null, "borderId", null, "1");
                await writer.WriteAttributeStringAsync(null, "xfId", null, "0");
                await writer.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyFill", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyBorder", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyProtection", null, "1");
                await writer.WriteStartElementAsync(null, "alignment", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "horizontal", null, "general");
                await writer.WriteAttributeStringAsync(null, "vertical", null, "bottom");
                await writer.WriteAttributeStringAsync(null, "textRotation", null, "0");
                await writer.WriteAttributeStringAsync(null, "wrapText", null, "0");
                await writer.WriteAttributeStringAsync(null, "indent", null, "0");
                await writer.WriteAttributeStringAsync(null, "relativeIndent", null, "0");
                await writer.WriteAttributeStringAsync(null, "justifyLastLine", null, "0");
                await writer.WriteAttributeStringAsync(null, "shrinkToFit", null, "0");
                await writer.WriteAttributeStringAsync(null, "readingOrder", null, "0");
                await writer.WriteEndElementAsync(); // alignment
                await writer.WriteStartElementAsync(null, "protection", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "locked", null, "1");
                await writer.WriteAttributeStringAsync(null, "hidden", null, "0");
                await writer.WriteEndElementAsync(); // protection
                await writer.WriteEndElementAsync(); // xf
                // Date
                await writer.WriteStartElementAsync(null, "xf", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "numFmtId", null, "14");
                await writer.WriteAttributeStringAsync(null, "fontId", null, "0");
                await writer.WriteAttributeStringAsync(null, "fillId", null, "0");
                await writer.WriteAttributeStringAsync(null, "borderId", null, "0");
                await writer.WriteAttributeStringAsync(null, "xfId", null, "0");
                await writer.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyFill", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyBorder", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyProtection", null, "1");
                await writer.WriteStartElementAsync(null, "alignment", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "horizontal", null, "general");
                await writer.WriteAttributeStringAsync(null, "vertical", null, "bottom");
                await writer.WriteAttributeStringAsync(null, "textRotation", null, "0");
                await writer.WriteAttributeStringAsync(null, "wrapText", null, "0");
                await writer.WriteAttributeStringAsync(null, "indent", null, "0");
                await writer.WriteAttributeStringAsync(null, "relativeIndent", null, "0");
                await writer.WriteAttributeStringAsync(null, "justifyLastLine", null, "0");
                await writer.WriteAttributeStringAsync(null, "shrinkToFit", null, "0");
                await writer.WriteAttributeStringAsync(null, "readingOrder", null, "0");
                await writer.WriteEndElementAsync(); // alignment
                await writer.WriteStartElementAsync(null, "protection", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "locked", null, "1");
                await writer.WriteAttributeStringAsync(null, "hidden", null, "0");
                await writer.WriteEndElementAsync(); // protection
                await writer.WriteEndElementAsync(); // xf
                // Date time
                await writer.WriteStartElementAsync(null, "xf", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "numFmtId", null, "22");
                await writer.WriteAttributeStringAsync(null, "fontId", null, "0");
                await writer.WriteAttributeStringAsync(null, "fillId", null, "0");
                await writer.WriteAttributeStringAsync(null, "borderId", null, "0");
                await writer.WriteAttributeStringAsync(null, "xfId", null, "0");
                await writer.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyFill", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyBorder", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyProtection", null, "1");
                await writer.WriteStartElementAsync(null, "alignment", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "horizontal", null, "general");
                await writer.WriteAttributeStringAsync(null, "vertical", null, "bottom");
                await writer.WriteAttributeStringAsync(null, "textRotation", null, "0");
                await writer.WriteAttributeStringAsync(null, "wrapText", null, "0");
                await writer.WriteAttributeStringAsync(null, "indent", null, "0");
                await writer.WriteAttributeStringAsync(null, "relativeIndent", null, "0");
                await writer.WriteAttributeStringAsync(null, "justifyLastLine", null, "0");
                await writer.WriteAttributeStringAsync(null, "shrinkToFit", null, "0");
                await writer.WriteAttributeStringAsync(null, "readingOrder", null, "0");
                await writer.WriteEndElementAsync(); // alignment
                await writer.WriteStartElementAsync(null, "protection", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "locked", null, "1");
                await writer.WriteAttributeStringAsync(null, "hidden", null, "0");
                await writer.WriteEndElementAsync(); // protection
                await writer.WriteEndElementAsync(); // xf
                // Decimal
                await writer.WriteStartElementAsync(null, "xf", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "numFmtId", null, "4");
                await writer.WriteAttributeStringAsync(null, "fontId", null, "0");
                await writer.WriteAttributeStringAsync(null, "fillId", null, "0");
                await writer.WriteAttributeStringAsync(null, "borderId", null, "0");
                await writer.WriteAttributeStringAsync(null, "xfId", null, "0");
                await writer.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyFill", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyBorder", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyProtection", null, "1");
                await writer.WriteStartElementAsync(null, "alignment", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "horizontal", null, "general");
                await writer.WriteAttributeStringAsync(null, "vertical", null, "bottom");
                await writer.WriteAttributeStringAsync(null, "textRotation", null, "0");
                await writer.WriteAttributeStringAsync(null, "wrapText", null, "0");
                await writer.WriteAttributeStringAsync(null, "indent", null, "0");
                await writer.WriteAttributeStringAsync(null, "relativeIndent", null, "0");
                await writer.WriteAttributeStringAsync(null, "justifyLastLine", null, "0");
                await writer.WriteAttributeStringAsync(null, "shrinkToFit", null, "0");
                await writer.WriteAttributeStringAsync(null, "readingOrder", null, "0");
                await writer.WriteEndElementAsync(); // alignment
                await writer.WriteStartElementAsync(null, "protection", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "locked", null, "1");
                await writer.WriteAttributeStringAsync(null, "hidden", null, "0");
                await writer.WriteEndElementAsync(); // protection
                await writer.WriteEndElementAsync(); // xf
                // Integer
                await writer.WriteStartElementAsync(null, "xf", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "numFmtId", null, "3");
                await writer.WriteAttributeStringAsync(null, "fontId", null, "0");
                await writer.WriteAttributeStringAsync(null, "fillId", null, "0");
                await writer.WriteAttributeStringAsync(null, "borderId", null, "0");
                await writer.WriteAttributeStringAsync(null, "xfId", null, "0");
                await writer.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyFill", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyBorder", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyProtection", null, "1");
                await writer.WriteStartElementAsync(null, "alignment", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "horizontal", null, "general");
                await writer.WriteAttributeStringAsync(null, "vertical", null, "bottom");
                await writer.WriteAttributeStringAsync(null, "textRotation", null, "0");
                await writer.WriteAttributeStringAsync(null, "wrapText", null, "0");
                await writer.WriteAttributeStringAsync(null, "indent", null, "0");
                await writer.WriteAttributeStringAsync(null, "relativeIndent", null, "0");
                await writer.WriteAttributeStringAsync(null, "justifyLastLine", null, "0");
                await writer.WriteAttributeStringAsync(null, "shrinkToFit", null, "0");
                await writer.WriteAttributeStringAsync(null, "readingOrder", null, "0");
                await writer.WriteEndElementAsync(); // alignment
                await writer.WriteStartElementAsync(null, "protection", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "locked", null, "1");
                await writer.WriteAttributeStringAsync(null, "hidden", null, "0");
                await writer.WriteEndElementAsync(); // protection
                await writer.WriteEndElementAsync(); // xf
                // Hyperlink
                await writer.WriteStartElementAsync(null, "xf", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "numFmtId", null, "0");
                await writer.WriteAttributeStringAsync(null, "fontId", null, "2");
                await writer.WriteAttributeStringAsync(null, "fillId", null, "0");
                await writer.WriteAttributeStringAsync(null, "borderId", null, "0");
                await writer.WriteAttributeStringAsync(null, "xfId", null, "0");
                await writer.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyFill", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyBorder", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
                await writer.WriteAttributeStringAsync(null, "applyProtection", null, "1");
                await writer.WriteStartElementAsync(null, "alignment", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "horizontal", null, "general");
                await writer.WriteAttributeStringAsync(null, "vertical", null, "bottom");
                await writer.WriteAttributeStringAsync(null, "textRotation", null, "0");
                await writer.WriteAttributeStringAsync(null, "wrapText", null, "0");
                await writer.WriteAttributeStringAsync(null, "indent", null, "0");
                await writer.WriteAttributeStringAsync(null, "relativeIndent", null, "0");
                await writer.WriteAttributeStringAsync(null, "justifyLastLine", null, "0");
                await writer.WriteAttributeStringAsync(null, "shrinkToFit", null, "0");
                await writer.WriteAttributeStringAsync(null, "readingOrder", null, "0");
                await writer.WriteEndElementAsync(); // alignment
                await writer.WriteStartElementAsync(null, "protection", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "locked", null, "1");
                await writer.WriteAttributeStringAsync(null, "hidden", null, "0");
                await writer.WriteEndElementAsync(); // protection
                await writer.WriteEndElementAsync(); // xf
                await writer.WriteEndElementAsync(); // cellXfs

                await writer.WriteStartElementAsync(null, "cellStyles", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "count", null, "1");
                await writer.WriteStartElementAsync(null, "cellStyle", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "name", null, "Normal");
                await writer.WriteAttributeStringAsync(null, "xfId", null, "0");
                await writer.WriteAttributeStringAsync(null, "builtinId", null, "0");
                await writer.WriteEndElementAsync(); // cellStyle
                await writer.WriteEndElementAsync(); // cellStyles

                await writer.WriteEndElementAsync(); // styleSheet
                await writer.WriteEndDocumentAsync();
            }
        
            _introWritten = true;
        }
    }

    private async Task WriteOutroAsync()
    {
        // Write the [Content-Types].xml file.
        using (var writer = XmlWriterFactory.Create(_zip.WriteToStream("[Content_Types].xml", new ZipWriterEntryOptions {})))
        {
            await writer.WriteStartDocumentAsync();
            await writer.WriteStartElementAsync(null, "Types", ContentTypesXmlNamespace);
            await writer.WriteStartElementAsync(null, "Default", ContentTypesXmlNamespace);
            await writer.WriteAttributeStringAsync(null, "Extension", null, "rels");
            await writer.WriteAttributeStringAsync(null, "ContentType", null, "application/vnd.openxmlformats-package.relationships+xml");
            await writer.WriteEndElementAsync(); // Default
            await writer.WriteStartElementAsync(null, "Default", ContentTypesXmlNamespace);
            await writer.WriteAttributeStringAsync(null, "Extension", null, "xml");
            await writer.WriteAttributeStringAsync(null, "ContentType", null, "application/xml");
            await writer.WriteEndElementAsync(); // Default
            await writer.WriteStartElementAsync(null, "Override", ContentTypesXmlNamespace);
            await writer.WriteAttributeStringAsync(null, "PartName", null, "/xl/workbook.xml");
            await writer.WriteAttributeStringAsync(null, "ContentType", null, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
            await writer.WriteEndElementAsync(); // Override
            await writer.WriteStartElementAsync(null, "Override", ContentTypesXmlNamespace);
            await writer.WriteAttributeStringAsync(null, "PartName", null, "/xl/styles.xml");
            await writer.WriteAttributeStringAsync(null, "ContentType", null, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
            await writer.WriteEndElementAsync(); // Override
            foreach (var sheet in _sheets)
            {
                await writer.WriteStartElementAsync(null, "Override", ContentTypesXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "PartName", null, $"/xl/worksheets/sheet{sheet.Key}.xml");
                await writer.WriteAttributeStringAsync(null, "ContentType", null, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                await writer.WriteEndElementAsync(); // Override
            }
            await writer.WriteEndElementAsync(); // Types
            await writer.WriteEndDocumentAsync();
        }
        // Write the workbook XML.
        using (var writer = XmlWriterFactory.Create(_zip.WriteToStream("/xl/workbook.xml", new ZipWriterEntryOptions {})))
        {
            await writer.WriteStartDocumentAsync();
            await writer.WriteStartElementAsync(null, "workbook", SpreadsheetMlXmlNamespace);
            await writer.WriteStartElementAsync(null, "sheets", SpreadsheetMlXmlNamespace);
            foreach (var sheet in _sheets)
            {
                await writer.WriteStartElementAsync(null, "sheet", SpreadsheetMlXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "name", null, sheet.Value);
                await writer.WriteAttributeStringAsync(null, "sheetId", null, sheet.Key.ToString(CultureInfo.InvariantCulture));
                await writer.WriteAttributeStringAsync(null, "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", $"rIdSheet{sheet.Key.ToString(CultureInfo.InvariantCulture)}");
                await writer.WriteEndElementAsync(); // sheet
            }
            await writer.WriteEndElementAsync(); // sheets
            await writer.WriteEndElementAsync(); // workbook
            await writer.WriteEndDocumentAsync();
        }
        // Write the workbook relationships XML.
        using (var writer = XmlWriterFactory.Create(_zip.WriteToStream("/xl/_rels/workbook.xml.rels", new ZipWriterEntryOptions {})))
        {
            await writer.WriteStartDocumentAsync();
            await writer.WriteStartElementAsync(null, "Relationships", RelationshipsXmlNamespace);
            await writer.WriteStartElementAsync(null, "Relationship", RelationshipsXmlNamespace);
            await writer.WriteAttributeStringAsync(null, "Type", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
            await writer.WriteAttributeStringAsync(null, "Target", null, "/xl/styles.xml");
            await writer.WriteAttributeStringAsync(null, "Id", null, "rIdStyle1");
            await writer.WriteEndElementAsync(); // Relationship
            foreach (var sheet in _sheets)
            {
                await writer.WriteStartElementAsync(null, "Relationship", RelationshipsXmlNamespace);
                await writer.WriteAttributeStringAsync(null, "Type", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                await writer.WriteAttributeStringAsync(null, "Target", null, $"/xl/worksheets/sheet{sheet.Key}.xml");
                await writer.WriteAttributeStringAsync(null, "Id", null, $"rIdSheet{sheet.Key}");
                await writer.WriteEndElementAsync(); // Relationship
            }
            await writer.WriteEndElementAsync(); // Relationships
            await writer.WriteEndDocumentAsync();
        }
    }
}
