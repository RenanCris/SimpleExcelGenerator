
using ConsoleApp;
using Microsoft.Extensions.DependencyInjection;
using SimpleExcelGenerator.Extensions;
using SimpleExcelGenerator.Interfaces;

var serviceCollection = new ServiceCollection();

serviceCollection.AddSimpleExcelGenerator();

var provider = serviceCollection.BuildServiceProvider();

var excelGenerator = provider.GetService<IExcelGenerator>();

var dataExample = 
        new List<TestClass> { 
            new TestClass(1,"Test"),
            new TestClass(2,"Other Test") }
        .AsEnumerable();


var resultAsync = await excelGenerator
                    .AddSheet("SheetTestName", dataExample)
                    .AddSheet("Other SheetTestName", dataExample)
                    .AddSheet("Other Sheet Final Test", dataExample)
                    .GenerateAsync();

Console.WriteLine($"Result Async {resultAsync}");

var resultSync =  excelGenerator
                    .AddSheet("SheetTestName", dataExample)
                    .AddSheet("Other SheetTestName", dataExample)
                    .AddSheet("Other Sheet Final Test", dataExample)
                    .GenerateSync();

Console.WriteLine($"Result Sync {resultSync}");



var resultBytes = await excelGenerator
                    .AddSheet("SheetTestName", dataExample)
                    .AddSheet("Other SheetTestName", dataExample)
                    .AddSheet("Other Sheet Final Test", dataExample)
                    .GetBytesAsync();

Console.WriteLine($"Result Bytes {resultBytes.Length}");


await excelGenerator
                    .AddSheet("SheetTestName", dataExample)
                    .AddSheet("Other SheetTestName", dataExample)
                    .AddSheet("Other Sheet Final Test", dataExample)
                    .SaveAsync(@$"C:\Temp\{Guid.NewGuid()}.xlsx");

var path = Directory.GetFiles(@$"C:\Temp\");

Console.WriteLine($"Async Save {path.Length}");


 excelGenerator
        .AddSheet("SheetTestName", dataExample)
        .AddSheet("Other SheetTestName", dataExample)
        .AddSheet("Other Sheet Final Test", dataExample)
        .SaveSync(@$"C:\Temp\{Guid.NewGuid()}.xlsx");


Console.WriteLine($"Sync Save {path.Length}");

