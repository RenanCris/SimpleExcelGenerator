using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleExcelGenerator.Atributtes
{
   
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttr : Attribute
    {
        public string DisplayName;
        public string Format;

        public ExcelColumnAttr(string displayName)
        {
            this.DisplayName = displayName;
        }

        public ExcelColumnAttr(string displayName, string format)
        {
            this.DisplayName = displayName;
            this.Format = format;
        }
    }
}
