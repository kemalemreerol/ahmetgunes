using System.Globalization;

namespace TranscriptApp.Models;

public sealed record TranscriptReport(
    string SourceFileName,
    string? StudentNumber,
    string? StudentFullName,
    string? StudentIdentityNumber,
    string? FacultyName,
    string? ProgramCode,
    TranscriptCalculationBasis CalculationBasis,
    TranscriptScope Scope,
    IReadOnlyList<TranscriptCourse> Courses,
    IReadOnlyList<TranscriptPeriodAverage> PeriodAverages,
    decimal? OverallAverage,
    TranscriptSummary Summary)
{
    public IReadOnlyList<IGrouping<TranscriptPeriod, TranscriptCourse>> GroupedCourses =>
        Courses
            .GroupBy(course => new TranscriptPeriod(course.Year, course.Semester))
            .OrderBy(group => group.Key)
            .ToList();

    public TranscriptPeriodAverage? GetAverage(TranscriptPeriod period) =>
        PeriodAverages.FirstOrDefault(average => average.Period == period);

    public decimal TotalCredit => Courses.Sum(course => ParseNumber(course.Credit));

    public decimal TotalEcts => Courses.Sum(course => ParseNumber(course.Ects));

    private static decimal ParseNumber(string value) =>
        decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out var number)
            ? number
            : 0;
}
