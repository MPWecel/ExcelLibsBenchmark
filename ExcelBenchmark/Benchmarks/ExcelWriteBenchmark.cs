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
public class ExcelWriteBenchmark
{
    private const string _directory = "Outputs";    //To run enter valid output directory
    private string GenerateFilename(string adapterName) => $"{adapterName}_output.xlsx";
    private string OutputPath(string filename) => $"{_directory}\\{filename}";

    private readonly List<(string Id, string Name)> _data = Enumerable.Range(1, 1000000).Select(x=>($"{x}", $"Element\t{x}")).ToList();

    private readonly List<IExcelAdapter> _adapters = new()
    {
        new NpoiAdapter(),
        new EpplusAdapter(),
        new ClosedXmlAdapter(),
        new MiniExcelAdapter()
    };

    public IEnumerable<IExcelAdapter> Adapters => _adapters;

    [GlobalSetup]
    public void Setup() => Directory.CreateDirectory(_directory);

    [Benchmark]
    [ArgumentsSource(nameof(Adapters))]
    public void WriteSampleData(IExcelAdapter adapter) => adapter.WriteRows(OutputPath(GenerateFilename(adapter.LibraryName)), _data);

}
