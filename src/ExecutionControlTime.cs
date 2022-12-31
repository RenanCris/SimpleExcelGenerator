using System;
using System.Diagnostics;
using System.Threading.Tasks;
using SimpleExcelGenerator.Interfaces;

namespace SimpleExcelGenerator
{
    public class ExecutionControlTime : IExecutionControlTime
    {
        public TimeSpan Execute(Action action)
        {
            Stopwatch stopWatch = new();
            stopWatch.Start();

            action();

            stopWatch.Stop();
            return stopWatch.Elapsed;
        }

        public async Task<TimeSpan> ExecuteAsync(Func<Task> action)
        {
            Stopwatch stopWatch = new();
            stopWatch.Start();

            await action();

            stopWatch.Stop();
            return stopWatch.Elapsed;
        }
    }
}
