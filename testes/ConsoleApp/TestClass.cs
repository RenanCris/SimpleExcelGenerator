using SimpleExcelGenerator.Atributtes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp
{
    public class TestClass
    {
        [ExcelColumnAttr("My Value")]
        public int Value { get; set; }

        [ExcelColumnAttr("My Custom Name")]
        public string Description { get; set; }

        public TestClass(int value, string description)
        {
            Value = value;
            Description = description;
        }
    }
}
