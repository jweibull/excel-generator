using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using x14 = DocumentFormat.OpenXml.Office2010.Excel;
using x15 = DocumentFormat.OpenXml.Office2013.Excel;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Text.RegularExpressions;
using Dia2Lib;
using Microsoft.Diagnostics.Runtime.DacInterface;
using Newtonsoft.Json;

namespace rbkApiModules.Utilities.Excel;

public class Runner
{
    public void Run()
    {
        var saxLib = new SaxLib();

        var serializer = new JsonSerializer();

        ModelData? modelData;

        string inputPath = Directory.GetCurrentDirectory();
        inputPath = Path.Combine(inputPath, "..", "..", "..", "input", "excel.json");

        using (StreamReader sr = new StreamReader(inputPath))
        using (var jsonTextReader = new JsonTextReader(sr))
        {
            modelData = serializer.Deserialize<ModelData>(jsonTextReader);
        }

        if (modelData == null)
        {
            throw new Exception("Data input was null.");
        }


        string outputPath = Directory.GetCurrentDirectory();
        outputPath = Path.Combine(outputPath, "..", "..", "..", "output");
        
        for (int i = 0; i < 1; i++)
        {
            var nameCounter = 1;
            var baseFilename = "output";
            var filename = baseFilename;
            while (File.Exists(Path.Combine(outputPath, filename + ".xlsx")))
            {
                filename = baseFilename + nameCounter++;
            }
            filename = Path.Combine(outputPath, filename + ".xlsx");

            var stream = saxLib.CreatePackage(modelData.WorkbookModel);

            using (var fileStream = File.Create(filename))
            {
                stream.Seek(0, SeekOrigin.Begin);
                stream.CopyTo(fileStream);
           }
        }       
    }
}
