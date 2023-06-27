using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Running;
using BenchmarkExcel;


BenchmarkSwitcher.FromTypes(new[] { typeof(AsyncBench) }).Run(args, new Config());
