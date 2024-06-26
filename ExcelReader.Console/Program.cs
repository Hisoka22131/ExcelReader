using ExcelReader.Console.Models;
using ExcelReader.Core;

var filePath = Path.Combine(GetWwwRootPath(), "Example.xlsx");
var fileBytes = File.ReadAllBytes(filePath);
if (!File.Exists(filePath))
{
    Console.WriteLine("Файл не найден.");
    return;
}
IExcelReader<User> reader = new ExcelReader<User>();

var result = await reader.ParseAsync(fileBytes);

foreach (var item in result)
{
    Console.WriteLine(item);
}

Console.WriteLine();
Console.WriteLine("Parsing completed.");

return;

static string GetWwwRootPath()
{
    var basePath = AppContext.BaseDirectory;
    var projectPath = Directory.GetParent(basePath)!.Parent!.Parent!.Parent!.FullName;
    return Path.Combine(projectPath, "wwwroot");
}