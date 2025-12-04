using MiniExcelLibs;

namespace ExcelBenchmark.Adapters;
public sealed class MiniExcelAdapter : IExcelAdapter
{
    public string LibraryName => "MiniExcel";

    public object ReadCell(string filePath, int rowNo, int columnNo)
    {
        List<string> sheetNames = MiniExcel.GetSheetNames(filePath);
        IDictionary<string, object> row = (IDictionary<string, object>)(MiniExcel.Query(filePath, sheetName: sheetNames.First()).First());
        object cellValue = row.Values.ElementAt(columnNo);
        return cellValue;
    }

    public void WriteRows(string filePath, IEnumerable<(string Id, string Name)> rows)
    { 
        var list = rows.Select(x => new {Id=x.Id, Name=x.Name});
        MiniExcel.SaveAs(filePath, list);
    }

    public override string ToString() => $"{LibraryName}Adapter";
}
