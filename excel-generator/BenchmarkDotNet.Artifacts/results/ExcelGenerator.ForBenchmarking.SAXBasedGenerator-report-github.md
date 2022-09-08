``` ini

BenchmarkDotNet=v0.13.2, OS=Windows 10 (10.0.19044.1889/21H2/November2021Update)
Intel Core i7-7700 CPU 3.60GHz (Kaby Lake), 1 CPU, 8 logical and 4 physical cores
.NET SDK=6.0.400
  [Host]     : .NET 6.0.8 (6.0.822.36306), X64 RyuJIT AVX2
  DefaultJob : .NET 6.0.8 (6.0.822.36306), X64 RyuJIT AVX2


```
| Method |    Mean |   Error |  StdDev |        Gen0 |      Gen1 |      Gen2 | Allocated |
|------- |--------:|--------:|--------:|------------:|----------:|----------:|----------:|
|    SAX | 10.13 s | 0.072 s | 0.068 s | 789000.0000 | 9000.0000 | 8000.0000 |   3.86 GB |
