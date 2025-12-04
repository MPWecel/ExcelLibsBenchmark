using OfficeOpenXml;

namespace ExcelBenchmark.Adapters;
public sealed class EpplusAdapter : IExcelAdapter
{
    public string LibraryName => "EPPlus";

    public EpplusAdapter()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }
    public object ReadCell(string filePath, int rowNo, int columnNo)
    {
        FileInfo info = new(filePath);
        using ExcelPackage package = new(info);
        ExcelWorkbook workbook = package.Workbook;
        ExcelWorksheet worksheet = workbook.Worksheets[0];
        ExcelRange cell = worksheet.Cells[
                                            ConvertIndex(rowNo),
                                            ConvertIndex(columnNo)
                                         ];
        object value = cell.Value;
        return value;
    }

    public void WriteRows(
                            string filePath, 
                            IEnumerable<(string Id, string Name)> rows
                         )
    {
        using ExcelPackage package = new();
        ExcelWorkbook workbook = package.Workbook;
        ExcelWorksheet worksheet = workbook.Worksheets.Add("Data");

        List<(string Id, string Name)> rowList = rows.ToList();
        for(int i=0; i<rowList.Count; i++)
        {
            worksheet.Cells[ConvertIndex(i), 1].Value = rowList[i].Id;
            worksheet.Cells[ConvertIndex(i), 2].Value = rowList[i].Name;
        }

        FileInfo info = new(filePath);
        package.SaveAs(info);
    }

    public override string ToString() => $"{LibraryName}Adapter";

    //EPPlus uses 1-based indices for rows and columns
    private int ConvertIndex(int index) => index + 1;
}
