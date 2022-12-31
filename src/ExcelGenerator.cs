using MiniExcelLibs;
using SimpleExcelGenerator.Exceptions;
using SimpleExcelGenerator.Helper;
using SimpleExcelGenerator.Interfaces;
using SimpleExcelGenerator.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace SimpleExcelGenerator
{
    public class ExcelGenerator : IExcelGenerator
    {
        public IList<Tuple<string, IEnumerable<object>>> ExcelGenerators { get; private set; }
        private DataSet _dataSetAllValues;
        private bool disposedValue;

        private readonly IExecutionControlTime _executionControlTime;

        public ExcelGenerator(IExecutionControlTime executionControlTime)
        {
            _executionControlTime = executionControlTime;
            ExcelGenerators = new List<Tuple<string, IEnumerable<object>>>();
        }

        public IExcelGenerator AddSheet(string name, IEnumerable<object> data)
        {
            if (ExcelGenerators.Any(x => x.Item1 == name)) throw new InvalidNameSheetCustomException("A sheet with that name has already been added.");

            ExcelGenerators.Add(new Tuple<string, IEnumerable<object>>(name, data));

            return this;
        }

        private void SetupDataSet()
        {
            _dataSetAllValues = new DataSet();
            
            foreach (var configData in ExcelGenerators)
            {
                var (nameTable, data) = configData;

                _dataSetAllValues.Tables.Add(data.ToDataTable(nameTable));
            }

            ValidExecution();
        }

        private void ValidExecution() 
        {
            if (_dataSetAllValues.Tables.Count == 0)
                throw new ArgumentException("ExcelGenerator is invalid, not set sheet in instance from this object.");
        }

        public async Task<ExcelResult> GenerateAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                byte[] bytes = default;

                var completTime = await _executionControlTime.ExecuteAsync(async () =>
                {
                    SetupDataSet();

                    using var stream = new MemoryStream();
                    await stream.SaveAsAsync(value: _dataSetAllValues, excelType: ExcelType.XLSX, cancellationToken: cancellationToken);

                    bytes = stream?.ToArray();
                });
                
                ExcelGenerators.Clear();

                return new ExcelResult(completTime, bytes);

            }
            catch
            {
                throw;
            }
        }

        public ExcelResult GenerateSync()
        {
            try
            {
                byte[] bytes = default;

                var completTime = _executionControlTime.Execute(() =>
                {
                    SetupDataSet();

                    using var stream = new MemoryStream();
                    stream.SaveAs(value: _dataSetAllValues, excelType: ExcelType.XLSX);

                    bytes = stream.ToArray();
                });

                ExcelGenerators.Clear();

                return new ExcelResult(completTime, bytes);

            }
            catch
            {
                throw;
            }
        }

        public async Task<byte[]> GetBytesAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                var result = await GenerateAsync(cancellationToken);
                return result?.Content;
            }
            catch
            {
                throw;
            }
        }

        public async Task SaveAsync(string path, CancellationToken cancellationToken = default)
        {
            try
            {
                SetupDataSet();
                await MiniExcel.SaveAsAsync(path, value: _dataSetAllValues, excelType: ExcelType.XLSX, cancellationToken: cancellationToken);

                ExcelGenerators.Clear();
            }
            catch
            {
                throw;
            }
        }

        public void SaveSync(string path)
        {
            try
            {
                SetupDataSet();
                MiniExcel.SaveAs(path, value: _dataSetAllValues, excelType: ExcelType.XLSX);

                ExcelGenerators.Clear();
            }
            catch
            {
                throw;
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    _dataSetAllValues.Dispose();
                    ExcelGenerators.Clear();
                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

    }
}
