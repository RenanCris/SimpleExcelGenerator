using Bogus;
using Microsoft.AspNetCore.Mvc;
using SimpleExcelGenerator.Interfaces;
using System.Collections;
using System.IO;

namespace WebAPI.Controllers
{
    [ApiController]
    [Route("api/simple-excel-generator")]
    public class SimpleExcelGeneratorController : ControllerBase
    {

        private readonly ILogger<SimpleExcelGeneratorController> _logger;
        private readonly IExcelGenerator _excelGenerator;

        public SimpleExcelGeneratorController(ILogger<SimpleExcelGeneratorController> logger, IExcelGenerator excelGenerator)
        {
            _logger = logger;
            _excelGenerator = excelGenerator;
        }

        [HttpGet("download-async")]
        public async Task<IActionResult> DownloadAsync(CancellationToken cancellationToken)
        {
            var occupations = new string[] { "gardener", "teacher", "writer", "programmer" };
            var faker = new Faker<TestExcel>()
                .RuleFor(u => u.Name, f => f.Name.FullName())
                .RuleFor(u => u.Occupation, f => f.PickRandom(occupations));

            var data = faker.Generate(2000);

            var result = await _excelGenerator
                            .AddSheet("Test", data)
                            .GenerateAsync(cancellationToken);

            return new FileContentResult(result.Content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                FileDownloadName = "demo.xlsx"
            };
        }

        [HttpGet("download-sync")]
        public IActionResult DownloadSync(CancellationToken cancellationToken)
        {
            var occupations = new string[] { "gardener", "teacher", "writer", "programmer" };
            var faker = new Faker<TestExcel>()
                .RuleFor(u => u.Name, f => f.Name.FullName())
                .RuleFor(u => u.Occupation, f => f.PickRandom(occupations));

            var data = faker.Generate(2000);

            var result =  _excelGenerator
                            .AddSheet("Test", data)
                            .GenerateSync();

            return new FileContentResult(result.Content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                FileDownloadName = "demo.xlsx"
            };
        }

        [HttpGet("result-file")]
        public async Task<IActionResult> ResultFileAsync(CancellationToken cancellationToken)
        {
            var occupations = new string[] { "gardener", "teacher", "writer", "programmer" };
            var faker = new Faker<TestExcel>()
                .RuleFor(u => u.Name, f => f.Name.FullName())
                .RuleFor(u => u.Occupation, f => f.PickRandom(occupations));

            var data = faker.Generate(2000);

            var result = await _excelGenerator
                            .AddSheet("Test", data)
                            .GenerateAsync(cancellationToken);

            return Ok(result.ToResponse());
        }
    }
}