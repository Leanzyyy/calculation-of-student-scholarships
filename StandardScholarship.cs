using System.Collections.Generic;

public class StandardScholarship : IScholarshipCalculator
{
    public string Currency { get; set; } = "руб.";

    public double CalculateAverageGrade(List<int> grades)
    {
        double sum = 0;
        foreach (int g in grades) sum += g;
        return sum / grades.Count;
    }

    public int DetermineScholarship(double avgGrade)
    {
        if (avgGrade >= 4.8) return 5000;
        if (avgGrade >= 4.0) return 3000;
        if (avgGrade >= 3.0) return 1500;
        return 0;
    }

    public string GetStudentStatus(double avgGrade)
    {
        if (avgGrade >= 4.8) return "Отличник";
        if (avgGrade >= 4.0) return "Хорошист";
        if (avgGrade >= 3.0) return "Троечник";
        return "Неуспевающий";
    }
}