using System.Text;
using System.Xml;

namespace DoIt.ExcelWriter;

static class XmlWriterFactory
{
    public static XmlWriter Create(Stream stream)
    {
        return XmlWriter.Create(stream, new XmlWriterSettings { Async = true, CloseOutput = true, Encoding = Encoding.UTF8, CheckCharacters = false });
    }
}
