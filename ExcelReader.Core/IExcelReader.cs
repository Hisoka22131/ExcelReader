namespace ExcelReader.Core;

public interface IExcelReader<T> where T: class
{
    Task<List<T>> ParseAsync(byte[] fileBytes, CancellationToken ct = default);
}