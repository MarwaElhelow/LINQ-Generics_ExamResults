using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

public class ExamResult<T>
{
    public string StudentName { get; set; }
    public int Score { get; set; }
    public T Result { get; set; }

    public ExamResult(string studentName, int score, T result)
    {
        StudentName = studentName;
        Score = score;
        Result = result;
    }
}

public class ExamResultManager<T> where T : IComparable<T>
{
    private List<ExamResult<T>> examResults;

    public ExamResultManager()
    {
        examResults = new List<ExamResult<T>>();
    }

    public void AddResult(string studentName, int score, T result)
    {
        ExamResult<T> examResult = new ExamResult<T>(studentName, score, result);
        examResults.Add(examResult);
    }

    public void PrintResults()
    {
        Console.WriteLine("Exam Results:");
        Console.WriteLine("------------------");
        foreach (var result in examResults)
        {
            Console.WriteLine("Student: " + result.StudentName + ", Score: " + result.Score + ", Result: " + result.Result);
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

    public List<ExamResult<T>> GetTopScorer()
    {
        int maxScore = examResults.Max(r => r.Score);
        return examResults.Where(r => r.Score == maxScore).ToList();
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

                    ExamResultManager<string> examResultManager = new ExamResultManager<string>();

                    for (int row = 2; row <= rowCount; row++)
                    {
                        string studentName = worksheet.Cells[row, 1].Value?.ToString();
                        int score = int.Parse(worksheet.Cells[row, 2].Value?.ToString());
                        string result = worksheet.Cells[row, 3].Value?.ToString();

                        examResultManager.AddResult(studentName, score, result);
                    }

                    examResultManager.PrintResults();

                    Console.WriteLine("Passed Results:");
                    List<ExamResult<string>> passedResults = examResultManager.GetPassedResults();
                    foreach (var result in passedResults)
                    {
                        Console.WriteLine("Student: " + result.StudentName + ", Result: " + result.Result);
                    }
                    Console.WriteLine("------------------");

                    Console.WriteLine("Failed Results:");
                    List<ExamResult<string>> failedResults = examResultManager.GetFailedResults();
                    foreach (var result in failedResults)
                    {
                        Console.WriteLine("Student: " + result.StudentName + ", Result: " + result.Result);
                    }
                    Console.WriteLine("------------------");

                    Console.WriteLine("Top Scorer(s):");
                    List<ExamResult<string>> topScorers = examResultManager.GetTopScorer();
                    foreach (var result in topScorers)
                    {
                        Console.WriteLine("Student: " + result.StudentName + ", Score: " + result.Score);
                    }
                }
                else
                {
                    Console.WriteLine("Worksheet '" + sheetName + "' not found in the Excel file.");
                }
            }
        }
        else
        {
            Console.WriteLine("Excel file not found at the specified path.");
        }

        Console.ReadLine();
    }

    //Add Comment for Test Fleet IDE
}