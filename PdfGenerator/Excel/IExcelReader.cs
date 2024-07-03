public interface IExcelReader
{
    IEnumerable<Student> ReadStudents(string filePath);
}