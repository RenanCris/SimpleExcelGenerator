using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleExcelGenerator.Response
{
    public class ExcelResultResponse
    {
        public long Size { get;  set; }
        public String SizeToString { get;  set; }
        public TimeSpan TimeGenerate { get;  set; }
        public string ContentBase64 { get;  set; }
        public string TypeFile { get; set; }
    }
}
