using System;
using BenchmarkDotNet.Running;
using ExcelBenchmark.Benchmarks;

namespace ExcelBenchmark;

public class Program
{
    public static void Main(string[] args) => GlobalBenchmarkRunner.Run();
}
