using SimpleExcelGenerator.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SimpleExcelGenerator.Interfaces
{
    public interface IExcelGenerator: IDisposable
    {
        Task<ExcelResult> GenerateAsync(CancellationToken cancellationToken = default);
        ExcelResult GenerateSync();
        Task<byte[]> GetBytesAsync(CancellationToken cancellationToken = default);
        Task SaveAsync(string path, CancellationToken cancellationToken = default);
        void SaveSync(string path);
        IExcelGenerator AddSheet(string name, IEnumerable<object> data);
    }
}
