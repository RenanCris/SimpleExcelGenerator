using AutoFixture;
using FluentAssertions;
using Moq;
using SimpleExcelGenerator;
using SimpleExcelGenerator.Exceptions;
using SimpleExcelGenerator.Interfaces;

namespace UnitTest
{
    public class ExcelGeneratorTest
    {

        private readonly Mock<IExecutionControlTime> executionControlTimeMock;
        private readonly Fixture fixture;

        public ExcelGeneratorTest()
        {
            executionControlTimeMock = new();
            fixture = new();
        }

        [Fact]
        public void AddSheet_AddNewTab()
        {
            var nameSheet = fixture.Create<string>();

            var excelGenerator = new ExcelGenerator(executionControlTimeMock.Object);
            excelGenerator.AddSheet(nameSheet, new[] { new { Test = "Test" } }.AsEnumerable());

            excelGenerator.ExcelGenerators.Should().HaveCount(1);
            excelGenerator.ExcelGenerators.First().Item1.Should().NotBeNull();
            excelGenerator.ExcelGenerators.First().Item1.Should().NotBeEmpty();
            excelGenerator.ExcelGenerators.First().Item1.Should().Be(nameSheet);

        }

        [Fact]
        public void AddSheet_AddNewTab_TrowsExecption_ByNameDuplicated()
        {
            var nameSheet = fixture.Create<string>();

            var excelGenerator = new ExcelGenerator(executionControlTimeMock.Object);
            excelGenerator.AddSheet(nameSheet, new[] { new { Test = "Test" } }.AsEnumerable());

            Assert.Throws<InvalidNameSheetCustomException>(() => excelGenerator.AddSheet(nameSheet, new[] { new { Test = "Test" } }.AsEnumerable()));

            excelGenerator.ExcelGenerators.Should().HaveCount(1);
        }

        [Fact]
        public async Task GenerateAsync_Generate_Success_ExcelResult()
        {
            var nameSheet = fixture.Create<string>();

            var excelGenerator = new ExcelGenerator(new ExecutionControlTime());
            excelGenerator.AddSheet(nameSheet, new[] { new { Test = "Test" } }.AsEnumerable());

            var result = await excelGenerator.GenerateAsync();

            excelGenerator.ExcelGenerators.Should().HaveCount(0);
            result.Should().NotBeNull();
            result.Size.Should().BeGreaterThan(0);
            result.Content.Should().HaveCountGreaterThan(0);
            result.TimeGenerate.Should().BeGreaterThan(TimeSpan.FromMilliseconds(1));
        }

        [Fact]
        public async Task GetBytesAsync_Generate_Success_ExcelResult()
        {
            var nameSheet = fixture.Create<string>();

            var excelGenerator = new ExcelGenerator(new ExecutionControlTime());
            excelGenerator.AddSheet(nameSheet, new[] { new { Test = "Test" } }.AsEnumerable());

            var result = await excelGenerator.GetBytesAsync();

            excelGenerator.ExcelGenerators.Should().HaveCount(0);
            result.Should().NotBeNull();
            result.Should().HaveCountGreaterThan(0);
        }

        [Fact]
        public void GenerateSync_Generate_Success_ExcelResult()
        {
            var nameSheet = fixture.Create<string>();

            var excelGenerator = new ExcelGenerator(new ExecutionControlTime());
            excelGenerator.AddSheet(nameSheet, new[] { new { Test = "Test" } }.AsEnumerable());

            var result = excelGenerator.GenerateSync();

            excelGenerator.ExcelGenerators.Should().HaveCount(0);
            result.Should().NotBeNull();
            result.Size.Should().BeGreaterThan(0);
            result.Content.Should().HaveCountGreaterThan(0);
            result.TimeGenerate.Should().BeGreaterThan(TimeSpan.FromMilliseconds(1));
        }

        [Fact]
        public async Task GenerateAsync_NotGenerate_Fail_ExcelResult()
        {
            var nameSheet = fixture.Create<string>();

            var excelGenerator = new ExcelGenerator(new ExecutionControlTime());

            await Assert.ThrowsAsync<ArgumentException>(async () => await excelGenerator.GenerateAsync()); 
        }

        [Fact]
        public async Task SaveAsync_Generate_Success_ExcelResult()
        {
            var nameSheet = fixture.Create<string>();
            var nameFile = fixture.Create<Guid>();

            var excelGenerator = new ExcelGenerator(new ExecutionControlTime());
            excelGenerator.AddSheet(nameSheet, new[] { new { Test = "Test" } }.AsEnumerable());

            var path = @$"C:\Temp\{nameFile}.xlsx";

            await excelGenerator.SaveAsync(path);

            File.ReadAllBytes(path).Should().HaveCountGreaterThan(0);
        }

        [Fact]
        public void SaveSync_Generate_Success_ExcelResult()
        {
            var nameSheet = fixture.Create<string>();
            var nameFile = fixture.Create<Guid>();

            var excelGenerator = new ExcelGenerator(new ExecutionControlTime());
            excelGenerator.AddSheet(nameSheet, new[] { new { Test = "Test" } }.AsEnumerable());

            var path = @$"C:\Temp\{nameFile}.xlsx";

             excelGenerator.SaveSync(path);

            File.ReadAllBytes(path).Should().HaveCountGreaterThan(0);
        }
    }
}