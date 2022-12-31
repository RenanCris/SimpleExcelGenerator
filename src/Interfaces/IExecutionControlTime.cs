using System;
using System.Threading.Tasks;

namespace SimpleExcelGenerator.Interfaces
{
    public interface IExecutionControlTime 
    {
        TimeSpan Execute(Action action);
        Task<TimeSpan> ExecuteAsync(Func<Task> action);
    }
}
