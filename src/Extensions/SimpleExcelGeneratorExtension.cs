using Microsoft.Extensions.DependencyInjection;
using SimpleExcelGenerator.Interfaces;

namespace SimpleExcelGenerator.Extensions
{
    public static class SimpleExcelGeneratorExtension
    {
        public static IServiceCollection AddSimpleExcelGenerator(this IServiceCollection service) 
        {
            service.AddScoped<IExcelGenerator, ExcelGenerator>();
            service.AddScoped<IExecutionControlTime, ExecutionControlTime>();

            return service;
        }
    }
}
