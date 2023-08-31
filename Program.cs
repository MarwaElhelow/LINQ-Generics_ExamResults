using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

public class ExamResult<T>
{
    public string StudentName { get; set; }
    public T Result { get; set; }

    public ExamResult(string studentName, T result)
    {
        StudentName = studentName;
        Result = result;
    }
}

public class ExamResultManager<T>
    where T : IComparable<T>
{
    private List<ExamResult<T>> examResults;

    public ExamResultManager()
    {
        examResults = new List<ExamResult<T>>();
    }

    public void AddResult(string studentName, T result)
    {
        ExamResult<T> examResult = new ExamResult<T>(studentName, result);
        examResults.Add(examResult);
    }

    public void PrintResults()
    {
        Console.WriteLine("Exam Results:");
        Console.WriteLine("------------------");
        foreach (var result in examResults)
        {
            Console.WriteLine("Student: " + result.StudentName + ", Result: " + result.Result);
        }
        Console.WriteLine("------------------");
    }

    public List<ExamResult<T>> GetPassedResults()
    {
        return examResults.Where(r => r.Result.Equals("Pass")).ToList();
    }

    public List<ExamResult<T>> GetFailedResults()
    {
        return examResults.Where(r => r.Result.Equals("Failed")).ToList();
    }

    public List<ExamResult<T>> GetTopScorers(int count)
    {
        List<ExamResult<T>> topScorers = new List<ExamResult<T>>();

        foreach (var result in examResults)
        {
            if (result.Result is int || result.Result is double || result.Result is string)
            {
                topScorers.Add(result);
            }
            else
            {
                Console.WriteLine("Invalid score format for student: " + result.StudentName);
            }
        }

        return topScorers.OrderByDescending(r => r.Result).Take(count).ToList();
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        string filePath = "D:\\Automation Learning\\Self Assignments\\ExamResults_Generics\\Exams_Scores.xlsx";
        string sheetName = "Scores_Sheet";

        if (File.Exists(filePath))
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet != null)
                {
                    int rowCount = worksheet.Dimension.Rows;

                    ExamResultManager<string> examResultManagerString = new ExamResultManager<string>();

                    for (int row = 2; row <= rowCount; row++)
                    {
                        string studentName = worksheet.Cells[row, 1].Value?.ToString();
                        string scoreStr = worksheet.Cells[row, 2].Value?.ToString();
                        string result = worksheet.Cells[row, 3].Value?.ToString();

                        if (int.TryParse(scoreStr, out int score))
                        {
                            examResultManagerString.AddResult(studentName, result);
                        }
                        else
                        {
                            Console.WriteLine("Invalid score at row " + row);
                        }
                    }

                    examResultManagerString.PrintResults();

                    Console.WriteLine("Passed Results:");
                    List<ExamResult<string>> passedResults = examResultManagerString.GetPassedResults();
                    foreach (var result in passedResults)
                    {
                        Console.WriteLine("Student: " + result.StudentName + ", Result: " + result.Result);
                    }

                    Console.WriteLine("Failed Results:");
                    List<ExamResult<string>> failedResults = examResultManagerString.GetFailedResults();
                    foreach (var result in failedResults)
                    {
                        Console.WriteLine("Student: " + result.StudentName + ", Result: " + result.Result);
                    }

                    Console.WriteLine("Top Scorers:");
                    List<ExamResult<string>> topScorers = examResultManagerString.GetTopScorers(3);
                    foreach (var result in topScorers)
                    {
                        Console.WriteLine("Student: " + result.StudentName + ", Result: " + result.Result);
                    }
                }
                //else if(){

                //}
                else
                {
                    Console.WriteLine("Worksheet '" + sheetName + "' not found in the Excel file.");
                }
            }
        }
        else
        {
            Console.WriteLine("Excel file not found at the specified location: " + filePath);
        }

        Console.ReadLine();
    }
}