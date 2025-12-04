using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BenchmarkDotNet.Running;

namespace ExcelBenchmark.Benchmarks;
public static class GlobalBenchmarkRunner
{
    public static void Run()
    {
        BenchmarkRunner.Run<ExcelReadBenchmark>();
        BenchmarkRunner.Run<ExcelWriteBenchmark>();
    }
}
