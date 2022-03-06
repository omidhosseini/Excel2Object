using System.IO;

namespace Excel2Object.Extensions
{
    public class ObjectToExcelFileResult
    {
        public string ContentType { get; set; }
        public string FileName { get; set; }
        public byte[] File { get; internal set; }
    }
}
