using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using TranscriptApp.Models;

namespace TranscriptApp.Services;

public sealed partial class TranscriptParser
{
    private const int HeaderSearchLimit = 20;
    private readonly AcademicUnitResolver _academicUnitResolver;

    public TranscriptParser(AcademicUnitResolver academicUnitResolver)
    {
        _academicUnitResolver = academicUnitResolver;
    }

    public TranscriptReport Parse(
        Stream workbookStream,
        string sourceFileName,
        TranscriptCalculationBasis calculationBasis = TranscriptCalculationBasis.Ects,
        TranscriptScope scope = TranscriptScope.AllCourses)
    {
        try
        {
            using var workbook = new XLWorkbook(workbookStream);
            var worksheet = workbook.Worksheets.FirstOrDefault()
                ?? throw new TranscriptProcessingException("Excel dosyasinda okunabilir bir sayfa bulunamadi.");

            var header = FindHeader(worksheet);
            if (scope == TranscriptScope.FirstTwoYears && header.NyyColumn <= 0)
            {
                throw new TranscriptProcessingException("İlk iki yıl filtresi için NYY sütunu bulunamadı.");
            }

            var rows = worksheet.RangeUsed()?.RowCount() ?? 0;
            var courses = new List<TranscriptCourse>();
            var emptySkipped = 0;
            var starredSkipped = 0;
            var invalidSkipped = 0;

            var lastDataRow = worksheet.LastRowUsed()?.RowNumber() ?? header.RowNumber;
            for (var rowNumber = header.RowNumber + 1; rowNumber <= lastDataRow; rowNumber++)
            {
                var rawNotes = GetCellText(worksheet, rowNumber, header.NotesColumn);
                var trimmedNotes = rawNotes.Trim();

                if (string.IsNullOrWhiteSpace(trimmedNotes))
                {
                    emptySkipped++;
                    continue;
                }

                if (trimmedNotes.EndsWith('*'))
                {
                    starredSkipped++;
                    continue;
                }

                if (!TryParseNotes(trimmedNotes, out var year, out var semester, out var letterGrade, out var transferCredit))
                {
                    invalidSkipped++;
                    continue;
                }

                var nyy = GetCellText(worksheet, rowNumber, header.NyyColumn);
                if (scope == TranscriptScope.FirstTwoYears && !IsFirstTwoYearNyy(nyy))
                {
                    continue;
                }

                courses.Add(new TranscriptCourse(
                    rowNumber,
                    year,
                    semester,
                    nyy,
                    GetCellText(worksheet, rowNumber, header.CourseCodeColumn),
                    GetCellText(worksheet, rowNumber, header.CourseNameColumn),
                    transferCredit ?? FormatCredit(GetCellText(worksheet, rowNumber, header.CreditColumn)),
                    FormatNumber(GetCellText(worksheet, rowNumber, header.EctsColumn)),
                    letterGrade));
            }

            if (courses.Count == 0)
            {
                throw new TranscriptProcessingException("Raporlanacak gecerli ders satiri bulunamadi.");
            }

            var orderedCourses = courses
                .OrderBy(course => new TranscriptPeriod(course.Year, course.Semester))
                .ThenBy(course => course.SourceRow)
                .ToList();
            var periodAverages = CalculatePeriodAverages(orderedCourses, calculationBasis);
            var overallAverage = CalculateAverage(orderedCourses, calculationBasis).Average;

            var studentInfo = ParseStudentInfo(worksheet.Cell(1, 1).GetString());
            var academicUnit = _academicUnitResolver.Resolve(studentInfo.StudentNumber);
            var summary = new TranscriptSummary(
                rows,
                emptySkipped,
                starredSkipped,
                invalidSkipped,
                orderedCourses.Count);

            return new TranscriptReport(
                sourceFileName,
                studentInfo.StudentNumber,
                null,
                null,
                academicUnit?.FacultyName,
                academicUnit?.DepartmentName ?? studentInfo.ProgramCode,
                calculationBasis,
                scope,
                orderedCourses,
                periodAverages,
                overallAverage,
                summary);
        }
        catch (TranscriptProcessingException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new TranscriptProcessingException("Excel dosyasi okunamadi. Lutfen dosya bicimini kontrol edin.", ex);
        }
    }

    private static HeaderMap FindHeader(IXLWorksheet worksheet)
    {
        var lastRow = Math.Min(HeaderSearchLimit, worksheet.LastRowUsed()?.RowNumber() ?? HeaderSearchLimit);
        var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;

        for (var rowNumber = 1; rowNumber <= lastRow; rowNumber++)
        {
            var columns = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (var columnNumber = 1; columnNumber <= lastColumn; columnNumber++)
            {
                var value = worksheet.Cell(rowNumber, columnNumber).GetString().Trim();
                if (!string.IsNullOrEmpty(value) && !columns.ContainsKey(value))
                {
                    columns[value] = columnNumber;
                }
            }

            if (columns.TryGetValue("notlar", out var notesColumn))
            {
                return new HeaderMap(
                    rowNumber,
                    notesColumn,
                    GetColumn(columns, "Ders Kodu"),
                    GetColumn(columns, "Ders Adı", "Ders Adi"),
                    GetColumn(columns, "Krd"),
                    GetColumn(columns, "AKTS"),
                    GetColumn(columns, "NYY"));
            }
        }

        throw new TranscriptProcessingException("notlar sutunu bulunamadi.");
    }

    private static int GetColumn(IReadOnlyDictionary<string, int> columns, params string[] names)
    {
        foreach (var name in names)
        {
            if (columns.TryGetValue(name, out var column))
            {
                return column;
            }
        }

        return 0;
    }

    private static string GetCellText(IXLWorksheet worksheet, int rowNumber, int columnNumber)
    {
        if (columnNumber <= 0)
        {
            return string.Empty;
        }

        return worksheet.Cell(rowNumber, columnNumber).GetFormattedString().Trim();
    }

    private static string FormatCredit(string value)
    {
        var withoutDetails = ParenthesizedDetailsRegex().Replace(value, string.Empty).Trim();
        return FormatNumber(withoutDetails);
    }

    private static string FormatNumber(string value)
    {
        var normalized = value.Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return string.Empty;
        }

        if (!decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.InvariantCulture, out var number))
        {
            return normalized;
        }

        return number.ToString("0.############################", CultureInfo.InvariantCulture);
    }

    private static bool IsFirstTwoYearNyy(string value) =>
        int.TryParse(value.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var nyy)
        && nyy is >= 1 and <= 4;

    private static IReadOnlyList<TranscriptPeriodAverage> CalculatePeriodAverages(
        IReadOnlyList<TranscriptCourse> courses,
        TranscriptCalculationBasis calculationBasis)
    {
        return courses
            .GroupBy(course => new TranscriptPeriod(course.Year, course.Semester))
            .OrderBy(group => group.Key)
            .Select(group =>
            {
                var result = CalculateAverage(group, calculationBasis);
                return new TranscriptPeriodAverage(
                    group.Key,
                    result.Average,
                    result.WeightTotal,
                    result.IncludedCourseCount,
                    SumCourseNumber(group, course => course.Credit),
                    SumCourseNumber(group, course => course.Ects));
            })
            .ToList();
    }

    private static decimal SumCourseNumber(IEnumerable<TranscriptCourse> courses, Func<TranscriptCourse, string> selector) =>
        courses.Sum(course =>
            decimal.TryParse(selector(course), NumberStyles.Number, CultureInfo.InvariantCulture, out var number)
                ? number
                : 0);

    private static AverageResult CalculateAverage(
        IEnumerable<TranscriptCourse> courses,
        TranscriptCalculationBasis calculationBasis)
    {
        decimal weightedScore = 0;
        decimal weightTotal = 0;
        var includedCourseCount = 0;

        foreach (var course in courses)
        {
            if (!TryGetGradeCoefficient(course.LetterGrade, out var coefficient)
                || !TryGetWeight(course, calculationBasis, out var weight)
                || weight <= 0)
            {
                continue;
            }

            weightedScore += coefficient * weight;
            weightTotal += weight;
            includedCourseCount++;
        }

        decimal? average = weightTotal == 0
            ? null
            : decimal.Round(weightedScore / weightTotal, 2, MidpointRounding.AwayFromZero);
        return new AverageResult(average, weightTotal, includedCourseCount);
    }

    private static bool TryGetWeight(
        TranscriptCourse course,
        TranscriptCalculationBasis calculationBasis,
        out decimal weight)
    {
        var value = calculationBasis == TranscriptCalculationBasis.Credit ? course.Credit : course.Ects;
        return decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out weight);
    }

    private static bool TryGetGradeCoefficient(string letterGrade, out decimal coefficient)
    {
        switch (letterGrade.Trim().ToUpperInvariant())
        {
            case "AA":
                coefficient = 4.00m;
                return true;
            case "BA":
                coefficient = 3.50m;
                return true;
            case "BB":
                coefficient = 3.00m;
                return true;
            case "CB":
                coefficient = 2.50m;
                return true;
            case "CC":
                coefficient = 2.00m;
                return true;
            case "DC":
                coefficient = 1.50m;
                return true;
            case "DD":
                coefficient = 1.00m;
                return true;
            case "FF":
            case "DZ":
                coefficient = 0.00m;
                return true;
            default:
                coefficient = 0;
                return false;
        }
    }

    public static bool TryParseNotes(
        string notes,
        out string year,
        out string semester,
        out string letterGrade,
        out string? transferCredit)
    {
        var value = notes.Replace("\r\n", "\n", StringComparison.Ordinal).Trim();
        year = string.Empty;
        semester = string.Empty;
        letterGrade = string.Empty;
        transferCredit = null;

        var transferMatch = TransferNotesRegex().Match(value);
        if (transferMatch.Success)
        {
            year = transferMatch.Groups["year"].Value;
            semester = transferMatch.Groups["semester"].Value.ToUpperInvariant();
            letterGrade = transferMatch.Groups["grade"].Value.ToUpperInvariant();
            transferCredit = FormatNumber(transferMatch.Groups["credit"].Value.Replace(',', '.'));

            return YearRegex().IsMatch(year)
                && !string.IsNullOrWhiteSpace(semester)
                && GradeRegex().IsMatch(letterGrade);
        }

        if (value.Length == 8)
        {
            year = value[..4];
            semester = value[4].ToString(CultureInfo.InvariantCulture);
            letterGrade = value[^2..];
        }
        else if (value.Length > 8 && !value.EndsWith('*'))
        {
            year = value.Substring(value.Length - 8, 4);
            semester = value.Substring(value.Length - 4, 1);
            letterGrade = value[^2..];
        }
        else
        {
            return false;
        }

        return YearRegex().IsMatch(year)
            && !string.IsNullOrWhiteSpace(semester)
            && GradeRegex().IsMatch(letterGrade);
    }

    private static StudentInfo ParseStudentInfo(string value)
    {
        var match = StudentInfoRegex().Match(value);
        if (!match.Success)
        {
            return new StudentInfo(null, null);
        }

        return new StudentInfo(
            match.Groups["number"].Value.Trim(),
            match.Groups["program"].Value.Trim());
    }

    [GeneratedRegex(@"^\d{4}$")]
    private static partial Regex YearRegex();

    [GeneratedRegex(@"\s*\([^)]*\)")]
    private static partial Regex ParenthesizedDetailsRegex();

    [GeneratedRegex(@"^(?<year>\d{4})(?<semester>[GBY])\s+.+?\s+(?<grade>AA|BA|BB|CB|CC|DC|DD|FF|DZ|YT|YZ|MU)\s+(?<credit>\d+(?:[.,]\d+)?)\s*krd\s*$", RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant)]
    private static partial Regex TransferNotesRegex();

    [GeneratedRegex(@"^[A-Za-zÇĞİÖŞÜçğıöşü]{2}$")]
    private static partial Regex GradeRegex();

    [GeneratedRegex(@"Öğrenci\s*No\s*:\s*(?<number>[^,|]+).*?\bDal\s*:\s*(?<program>[^,|]+)", RegexOptions.IgnoreCase)]
    private static partial Regex StudentInfoRegex();

    private sealed record HeaderMap(
        int RowNumber,
        int NotesColumn,
        int CourseCodeColumn,
        int CourseNameColumn,
        int CreditColumn,
        int EctsColumn,
        int NyyColumn);

    private sealed record StudentInfo(string? StudentNumber, string? ProgramCode);

    private sealed record AverageResult(decimal? Average, decimal WeightTotal, int IncludedCourseCount);
}
