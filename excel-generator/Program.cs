using ExcelGenerator.Generators;
;

var excelLib = new MinimalBoilerPlate();

string path = Directory.GetCurrentDirectory();
path = Path.Combine(path, "..", "..", "..", "output");
for (int i = 0; i < 1; i++)
{
    var nameCounter = 1;
    var baseFilename = "output";
    var filename = baseFilename;
    while (File.Exists(Path.Combine(path, filename + ".xlsx")))
    {
        filename = baseFilename + nameCounter++;
    }
    filename = Path.Combine(path, filename + ".xlsx");

    excelLib.CreatePackage(filename);
}




