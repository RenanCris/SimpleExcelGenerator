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
    public class ExcelGenerator : ExcelGeneratorBase , IExcelGenerator
    {
        private bool disposedValue;
        public ExcelGenerator(IExecutionControlTime executionControlTime) : base(executionControlTime)
        {
        }

        public IExcelGenerator AddSheet(string name, IEnumerable<object> data)
        {
            if (ExcelGenerators.Any(x => x.Item1 == name)) throw new InvalidNameSheetCustomException("A sheet with that name has already been added.");

            ExcelGenerators.Add(new Tuple<string, IEnumerable<object>>(name, data));

            return this;
        }

        public async Task<ExcelResult> GenerateAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                byte[] bytes = default;

                var completTime = await _executionControlTime.ExecuteAsync(async () =>
                {
                    bytes = await ExecuteBaseAsync(async () =>
                    {
                        using var stream = new MemoryStream();
                        await stream.SaveAsAsync(value: _dataSetAllValues, excelType: ExcelType.XLSX, cancellationToken: cancellationToken);

                        return stream?.ToArray();
                    });
                });

                return new ExcelResult(completTime, bytes);

            }
            catch(Exception ex)
            {
                throw new SimpleExcelCustomException("[GenerateAsync] Error!", ex);
            }
        }

        public ExcelResult GenerateSync()
        {
            try
            {
                byte[] bytes = default;

                var completTime = _executionControlTime.Execute(() =>
                {
                    bytes = ExecuteBaseSync(() =>
                    {
                        using var stream = new MemoryStream();
                        stream.SaveAs(value: _dataSetAllValues, excelType: ExcelType.XLSX);

                        return stream.ToArray();
                    });
                });

                return new ExcelResult(completTime, bytes);

            }
            catch (Exception ex)
            {
                throw new SimpleExcelCustomException("[GenerateSync] Error!", ex);
            }
        }

        public async Task<byte[]> GetBytesAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                var result = await GenerateAsync(cancellationToken);
                return result?.Content;
            }
            catch (Exception ex)
            {
                throw new SimpleExcelCustomException("[GetBytesAsync] Error!", ex);
            }
        }

        public async Task SaveAsync(string path, CancellationToken cancellationToken = default)
        {
            try
            {
                await ExecuteBaseAsync(async () =>
                {
                    await MiniExcel.SaveAsAsync(path, value: _dataSetAllValues, excelType: ExcelType.XLSX, cancellationToken: cancellationToken);
                });
            }
            catch (Exception ex)
            {
                throw new SimpleExcelCustomException("[SaveAsync] Error!", ex);
            }
        }

        public void SaveSync(string path)
        {
            try
            {
                ExecuteBaseSync(() => MiniExcel.SaveAs(path, value: _dataSetAllValues, excelType: ExcelType.XLSX));
            }
            catch (Exception ex)
            {
                throw new SimpleExcelCustomException("[SaveSync] Error!", ex);
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
