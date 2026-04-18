namespace TranscriptApp.Models;

public sealed record TranscriptPeriod(string Year, string Semester) : IComparable<TranscriptPeriod>
{
    public int CompareTo(TranscriptPeriod? other)
    {
        if (other is null)
        {
            return 1;
        }

        var yearCompare = ParseYear(Year).CompareTo(ParseYear(other.Year));
        if (yearCompare != 0)
        {
            return yearCompare;
        }

        return SemesterOrder(Semester).CompareTo(SemesterOrder(other.Semester));
    }

    private static int ParseYear(string value) =>
        int.TryParse(value, out var year) ? year : int.MaxValue;

    private static int SemesterOrder(string value) =>
        value.Trim().ToUpperInvariant() switch
        {
            "G" => 1,
            "B" => 2,
            "Y" => 3,
            _ => 99
        };
}
