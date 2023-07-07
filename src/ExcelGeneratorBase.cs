using SimpleExcelGenerator.Exceptions;
using SimpleExcelGenerator.Helper;
using SimpleExcelGenerator.Interfaces;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace SimpleExcelGenerator
{
    public abstract class ExcelGeneratorBase 
    {
        public IList<Tuple<string, IEnumerable<object>>> ExcelGenerators { get; private set; }
        protected DataSet _dataSetAllValues;

        protected readonly IExecutionControlTime _executionControlTime;
        public ExcelGeneratorBase(IExecutionControlTime executionControlTime)
        {
            _executionControlTime = executionControlTime;
            ExcelGenerators = new List<Tuple<string, IEnumerable<object>>>();
        }
        protected void SetupDataSet()
        {
            _dataSetAllValues = new DataSet();

            try
            {
                foreach (var configData in ExcelGenerators)
                {
                    var (nameTable, data) = configData;

                    if (data is null || (data != null && !data.Any())) continue;

                    var dataTableComplete = data.ToDataTable(nameTable);

                    _dataSetAllValues.Tables.Add(dataTableComplete);
                }
            }
            catch (Exception ex)
            {
                throw new SimpleExcelCustomException("[SetupDataSet] Error!", ex);
            }

            ValidExecution();
        }

        private void ValidExecution()
        {
            if (_dataSetAllValues.Tables.Count == 0)
                throw new ArgumentException("The content is invalid. At least one sheet needs to be inserted into the spreadsheet in order to generate the file.");
        }

        protected async Task<byte[]> ExecuteBaseAsync(Func<Task<byte[]>> action) 
        {
            SetupDataSet();
            var result = await action();
            Clear();

            return result;
        }

        protected async Task ExecuteBaseAsync(Func<Task> action)
        {
            SetupDataSet();
            await action();
            Clear();
        }

        protected void ExecuteBaseSync(Action action)
        {
            SetupDataSet();
            action();
            Clear();
        }

        protected byte[] ExecuteBaseSync(Func<byte[]> action)
        {
            SetupDataSet();
            byte[] result = action();
            Clear();

            return result;
        }

        protected virtual void Clear()
        {
            _dataSetAllValues?.Dispose();
            ExcelGenerators.Clear();
        }
    }
}
