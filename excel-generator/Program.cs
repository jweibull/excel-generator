using BenchmarkDotNet.Running;
using ExcelGenerator.ForBenchmarking;

var summary1 = BenchmarkRunner.Run<ClosedXMLBasedGenerator>();
var summary2 = BenchmarkRunner.Run<DOMBasedGenerator>();
var summary3 = BenchmarkRunner.Run<SAXBasedGenerator>();
var summary4 = BenchmarkRunner.Run<EPPlusFreeXMLBasedGenerator>();



