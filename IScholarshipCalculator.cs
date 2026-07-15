using System.Collections.Generic;

public interface IScholarshipCalculator
{
    double CalculateAverageGrade(List<int> grades);
    int DetermineScholarship(double avgGrade);
    string GetStudentStatus(double avgGrade);
    string Currency { get; set; }
}