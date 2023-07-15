#   SimpleExcelGenerator

This repository contains the library that provides functionality for wrapping [MiniExcel](https://github.com/mini-software/MiniExcel) code and adding it in your application via the dependency injection mechanism. That wrapper provides easy generator file .xlsx using objects in memory.

# Package

	NuGet\Install-Package SimpleExcelGenerator -Version 1.0.3

# Usage

In your class Program.cs add this code:

    builder.Services.AddSimpleExcelGenerator();
This interface for generate your file xlsx:

    IExcelGenerator _excelGenerator

## Functions

    public interface IExcelGenerator: IDisposable
    {
        Task<ExcelResult> GenerateAsync(CancellationToken cancellationToken = default);
        ExcelResult GenerateSync();
        Task<byte[]> GetBytesAsync(CancellationToken cancellationToken = default);
        Task SaveAsync(string path, CancellationToken cancellationToken = default);
        void SaveSync(string path);
        IExcelGenerator AddSheet(string name, IEnumerable<object> data);
    }

## Usage Example

        var dataExample = 
	        new List<TestClass> { 
	            new TestClass(1,"Test"),
	            new TestClass(2,"Other Test") }
	        .AsEnumerable();
        
        var resultAsync = await excelGenerator
                    .AddSheet("SheetTestName", dataExample)
                    .AddSheet("Other SheetTestName", dataExample)
                    .AddSheet("Other SheetFinalTest", dataExample)
                    .GenerateAsync();
                    
        Console.WriteLine($"Result Async {resultAsync}");
                    
**Result:** `$"Size: {this.Size}, Size Descption {this.SizeToString}, TotalMilliseconds: {this.TimeGenerate.TotalMilliseconds}";`

![console](https://user-images.githubusercontent.com/7238977/210172567-7f1364f4-e787-4e9e-b96e-5948264c8d25.PNG)


## To perform a test , locally

Use the docker image and run via api. 

    docker build -t excel:latest .
    docker run  --name mynewapi -p 5000:5000 -p 5001:5001 -e ASPNETCORE_ENVIRONMENT=Development -e ASPNETCORE_HTTP_PORT=http://:+5001 -e ASPNETCORE_URLS=http://+:5000 excel:latest

![Capturar](https://user-images.githubusercontent.com/7238977/210172547-84506608-918d-4c75-b203-ad5aa3916c89.PNG)

## Download Result
![excel](https://user-images.githubusercontent.com/7238977/210172653-410df30c-3d0a-4cd7-8e19-a8d2cd9183f8.PNG)

## Custom Header Column Attr

    [ExcelColumnAttr("My Value")]
    public int Value { get; set; }


![Attr](https://user-images.githubusercontent.com/7238977/210172710-a5f8911a-d7fd-49b7-8aa3-9b050ee9e495.PNG)




