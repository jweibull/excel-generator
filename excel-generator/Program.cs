using BenchmarkDotNet.Running;
using ExcelGenerator.ForBenchmarking;
using ExcelGenerator.Generators;

//var summary1 = BenchmarkRunner.Run<ClosedXMLBasedGenerator>();
//var summary2 = BenchmarkRunner.Run<DOMBasedGenerator>();
//var summary3 = BenchmarkRunner.Run<SAXBasedGenerator>();
//var summary4 = BenchmarkRunner.Run<EPPlusFreeXMLBasedGenerator>();
var summary5 = BenchmarkRunner.Run<SAXSharedGenerator>();

//var runner = new SaxLib();
//runner.Run();

