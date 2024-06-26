using System.Reflection;
using OfficeOpenXml;

namespace ExcelReader.Core;

public sealed class ExcelReader<T> : IExcelReader<T> where T : class, new()
{
    public async Task<List<T>> ParseAsync(byte[] fileBytes, CancellationToken ct = default)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var result = new List<T>();

        try
        {
            using var package = new ExcelPackage(new MemoryStream(fileBytes));
            var worksheet = package.Workbook.Worksheets[0];

            var rowCount = worksheet.Dimension.Rows;
            var colCount = worksheet.Dimension.Columns;

            var headers = new List<string>();
            for (var col = 1; col <= colCount; col++)
            {
                headers.Add(worksheet.Cells[1, col].Text);
            }

            for (var row = 2; row <= rowCount; row++)
            {
                var obj = new T();
                var objType = typeof(T);

                for (var col = 1; col <= colCount; col++)
                {
                    var header = headers[col - 1];
                    var cellValue = worksheet.Cells[row, col].Text;

                    var property = objType.GetProperty(header, BindingFlags.Public | BindingFlags.Instance);
                    if (property != null && property.CanWrite)
                    {
                        object convertedValue = ConvertValue(cellValue, property.PropertyType);
                        property.SetValue(obj, convertedValue);
                    }
                }

                result.Add(obj);
            }
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw;
        }

        return await Task.FromResult(result);
    }
    
    private object ConvertValue(string value, Type targetType)
    {
        if (targetType == typeof(string))
        {
            return value;
        }
        if (targetType == typeof(int) || targetType == typeof(int?))
        {
            return int.TryParse(value, out var result) ? result : (int?)null;
        }
        if (targetType == typeof(double) || targetType == typeof(double?))
        {
            return double.TryParse(value, out var result) ? result : (double?)null;
        }
        if (targetType == typeof(DateTime) || targetType == typeof(DateTime?))
        {
            return DateTime.TryParse(value, out var result) ? result : (DateTime?)null;
        }
        if (targetType == typeof(DateTimeOffset) || targetType == typeof(DateTimeOffset?))
        {
            return DateTimeOffset.TryParse(value, out var result) ? result : (DateTimeOffset?)null;
        }
        
        throw new InvalidCastException($"Cannot convert '{value}' to {targetType.Name}");
    }
}