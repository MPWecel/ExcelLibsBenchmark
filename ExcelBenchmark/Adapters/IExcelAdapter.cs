using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBenchmark.Adapters;
public interface IExcelAdapter
{
    string LibraryName { get; }

    //rowNo & columnNo are 0-based indices; formatting them to library specific 0-based index or 1-based index is handled by each interface implementation
    object ReadCell(string filePath, int rowNo, int columnNo);
    void WriteRows(string filePath, IEnumerable<(string Id, string Name)> rows);
}
