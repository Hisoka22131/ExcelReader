namespace ExcelReader.Console.Models;

public class User
{
    public string Name { get; set; }
    public int Age { get; set; }
    public DateTimeOffset DateOfBirthday { get; set; }
    
    public override string ToString()
    {
        return $"Name: {Name}\n" +
               $"Age: {Age}\n" +
               $"DateOfBirthday: {DateOfBirthday}";
    }
}