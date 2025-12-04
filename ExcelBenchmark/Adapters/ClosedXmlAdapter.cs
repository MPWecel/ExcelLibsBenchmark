using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;

namespace ExcelBenchmark.Adapters;
public sealed class ClosedXmlAdapter : IExcelAdapter
{
    public string LibraryName => "ClosedXML";

    public object ReadCell(string filePath, int rowNo, int columnNo)
    {
        using XLWorkbook workbook = new(filePath);
        IXLWorksheet worksheet = workbook.Worksheet(1);
        IXLCell cell = worksheet.Cell(
                                        ConvertIndex(rowNo),
                                        ConvertIndex(columnNo)
                                     );
        object value = cell.Value;
        return value;
    }

    public void WriteRows(
                            string filePath, 
                            IEnumerable<(string Id, string Name)> rows
                         )
    { 
        using XLWorkbook workbook = new();
        IXLWorksheet worksheet = workbook.AddWorksheet("Data");

        int rowNo = 1;
        foreach (var row in rows)
        {
            worksheet.Cell(rowNo, 1).SetValue(row.Id);
            worksheet.Cell(rowNo++, 2).SetValue(row.Name);
        }

        workbook.SaveAs(filePath);
    }

    public override string ToString() => $"{LibraryName}Adapter";

    //ClosedXML uses 1-based indices for rows and columns
    private int ConvertIndex(int index) => index + 1;
}
