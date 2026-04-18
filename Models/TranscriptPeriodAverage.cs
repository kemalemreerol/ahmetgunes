namespace TranscriptApp.Models;

public sealed record TranscriptPeriodAverage(
    TranscriptPeriod Period,
    decimal? Average,
    decimal WeightTotal,
    int IncludedCourseCount,
    decimal TotalCredit,
    decimal TotalEcts);
