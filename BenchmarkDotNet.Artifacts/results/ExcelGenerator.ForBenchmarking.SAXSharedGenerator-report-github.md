``` ini

BenchmarkDotNet=v0.13.2, OS=Windows 10 (10.0.19044.2006/21H2/November2021Update)
Intel Core i7-7700 CPU 3.60GHz (Kaby Lake), 1 CPU, 8 logical and 4 physical cores
.NET SDK=6.0.401
  [Host]     : .NET 6.0.9 (6.0.922.41905), X64 RyuJIT AVX2
  DefaultJob : .NET 6.0.9 (6.0.922.41905), X64 RyuJIT AVX2


```
|     Method |    Mean |   Error |  StdDev |        Gen0 |      Gen1 |      Gen2 | Allocated |
|----------- |--------:|--------:|--------:|------------:|----------:|----------:|----------:|
| SAX_SHARED | 12.84 s | 0.173 s | 0.170 s | 807000.0000 | 3000.0000 | 2000.0000 |   4.33 GB |
