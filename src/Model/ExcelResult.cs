using SimpleExcelGenerator.Helper;
using SimpleExcelGenerator.Response;
using System;

namespace SimpleExcelGenerator.Model
{
    public class ExcelResult
    {
        public long Size { get; private set; }
        public String SizeToString { get; private set; }
        public TimeSpan TimeGenerate { get; private set; }
        public byte[] Content { get; private set; }

        public ExcelResult(TimeSpan timeGenerate, byte[] content)
        {
            TimeGenerate = timeGenerate;
            Content = content;
            Size = content is null ?  0 : content.Length;
            SizeToString = Size.BytesToString();
        }


        public override string ToString()
        {
            return $"Size: {this.Size}, Size Description {this.SizeToString}, TotalMilliseconds: {this.TimeGenerate.TotalMilliseconds}";
        }

        public ExcelResultResponse ToResponse()
        {
            return new ExcelResultResponse { 
             Size = Size, 
             SizeToString = SizeToString,
             ContentBase64= Convert.ToBase64String(Content),
             TimeGenerate  = TimeGenerate
            };
        }
    }
}
