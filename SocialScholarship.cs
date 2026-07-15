using System.Collections.Generic;

public class SocialScholarship : IScholarshipCalculator
{
    public string Currency { get; set; } = "руб.";
    private bool _isLargeFamily;

    public SocialScholarship(bool isLargeFamily)
    {
        _isLargeFamily = isLargeFamily;
    }

    public double CalculateAverageGrade(List<int> grades)
    {
        int count = 0;
        double sum = 0;
        foreach (int g in grades)
            if (g >= 3) { sum += g; count++; }
        return count == 0 ? 0 : sum / count;
    }

    public int DetermineScholarship(double avgGrade)
    {
        int baseScholarship = 0;
        if (avgGrade >= 4.8) baseScholarship = 5000;
        else if (avgGrade >= 4.0) baseScholarship = 3000;
        else if (avgGrade >= 3.0) baseScholarship = 1500;

        if (_isLargeFamily && avgGrade >= 3.0)
            return baseScholarship + 2000;
        return baseScholarship;
    }

    public string GetStudentStatus(double avgGrade)
    {
        if (_isLargeFamily && avgGrade >= 3.0)
            return "Социальная (многодетная семья)";
        if (avgGrade >= 4.8) return "Отличник";
        if (avgGrade >= 4.0) return "Хорошист";
        if (avgGrade >= 3.0) return "Троечник";
        return "Неуспевающий";
    }
}