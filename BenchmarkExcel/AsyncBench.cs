using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using MiniExcelLibs;
using SimpleExcelGenerator;
using SimpleExcelGenerator.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace BenchmarkExcel
{
    public class AsyncBench
    {
        [Benchmark(Description = "Async Generate Static Object")]
        public async Task ExcelCreateAsyncTestStaticObject()
        {
            var controllTime = new ExecutionControlTime();
            var excelGenerator = new ExcelGenerator(controllTime);

            excelGenerator.AddSheet("Test One", new[] { new { Test = "Test" }, new { Test = "Test 2" }, new { Test = "Test 3" } });

            await excelGenerator.GenerateAsync();
        }

        [Benchmark(Description = "Async Generate Static Object MiniExcel")]
        public async Task ExcelCreateAsyncTestStaticObjectMiniExcel()
        {
            byte[] bytes = default;

            using var stream = new MemoryStream();
            await stream.SaveAsAsync(value: new[] { new { Test = "Test" }, new { Test = "Test 2" }, new { Test = "Test 3" } }, excelType: ExcelType.XLSX);

            bytes = stream?.ToArray();
        }

        [Benchmark(Description = "Async Generate Static Object 10000")]
        public async Task ExcelCreateAsyncTestStaticObject_10000()
        {
            var controllTime = new ExecutionControlTime();
            var excelGenerator = new ExcelGenerator(controllTime);

            excelGenerator.AddSheet("Test One", Data.GetValues(10000));

            await excelGenerator.GenerateAsync();
        }

        [Benchmark(Description = "Async Generate Static Object MiniExcel 10000")]
        public async Task ExcelCreateAsyncTestStaticObjectMiniExcel_10000()
        {
            byte[] bytes = default;

            using var stream = new MemoryStream();
            await stream.SaveAsAsync(value: Data.GetValues(10000), excelType: ExcelType.XLSX);

            bytes = stream?.ToArray();
        }
    }
}
