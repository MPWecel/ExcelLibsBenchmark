using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using BenchmarkDotNet.Attributes;
using NPOI;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using OfficeOpenXml;
using ClosedXML.Excel;
using MiniExcelLibs;
using Org.BouncyCastle.Asn1.Mozilla;
using ExcelBenchmark.Adapters;

namespace ExcelBenchmark.Benchmarks;

[MemoryDiagnoser]
public class ExcelReadBenchmark
{
    private const string _directory = "C:\\Users\\MWecel\\source\\repos\\ExcelBenchmark\\ExcelBenchmark\\TestFiles\\";  //To run enter valid input directory
    private const string _fileName = "testFile1.xlsx";

    public string FilePath => $"{_directory}{_fileName}";

    private readonly List<IExcelAdapter> _adapters = new()
    {
        new NpoiAdapter(),
        new EpplusAdapter(),
        new ClosedXmlAdapter(),
        new MiniExcelAdapter()
    };

    public IEnumerable<IExcelAdapter> Adapters => _adapters;

    [Benchmark]
    [ArgumentsSource(nameof(Adapters))]
    public object ReadFirstCell(IExcelAdapter adapter) => adapter.ReadCell(FilePath, 0, 0);

}
