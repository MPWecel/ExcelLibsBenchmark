using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelBenchmark.Adapters;
public sealed class NpoiAdapter : IExcelAdapter
{
    public string LibraryName => "NPOI";

    public object ReadCell(string filePath, int rowNo, int columnNo)
    {
        using FileStream fs = new(filePath, FileMode.Open, FileAccess.Read);
        XSSFWorkbook workbook = new(fs);
        ISheet worksheet = workbook.GetSheetAt(0);
        IRow row = worksheet.GetRow(ConvertIndex(rowNo));
        ICell cell = row.GetCell(ConvertIndex(columnNo));
        object value = cell.StringCellValue;
        return value;
    }

    public void WriteRows(
                            string filePath, 
                            IEnumerable<(string Id, string Name)> rows
                         )
    { 
        XSSFWorkbook workbook = new();
        ISheet worksheet = workbook.CreateSheet("Data");

        int rowNo = 0;
        foreach (var row in rows)
        {
            IRow rowObject = worksheet.CreateRow(rowNo++);
            rowObject.CreateCell(0, CellType.String).SetCellValue(row.Id);
            rowObject.CreateCell(1, CellType.String).SetCellValue(row.Name);
        }

        using FileStream fileStream = new(filePath, FileMode.Create, FileAccess.Write);
        workbook.Write(fileStream);
    }

    public override string ToString() => $"{LibraryName}Adapter";

    //NPOI uses zero-based row/column indices
    private int ConvertIndex(int index) => index;
}
