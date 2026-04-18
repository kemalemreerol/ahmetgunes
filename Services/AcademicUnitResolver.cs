using System.Globalization;
using System.Text;
using ClosedXML.Excel;

namespace TranscriptApp.Services;

public sealed class AcademicUnitResolver
{
    private readonly IWebHostEnvironment _environment;
    private readonly Lazy<IReadOnlyDictionary<string, AcademicUnit>> _units;

    public AcademicUnitResolver(IWebHostEnvironment environment)
    {
        _environment = environment;
        _units = new Lazy<IReadOnlyDictionary<string, AcademicUnit>>(LoadUnits);
    }

    public AcademicUnitResolution? Resolve(string? studentNumber)
    {
        var digits = new string((studentNumber ?? string.Empty).Where(char.IsDigit).ToArray());
        if (digits.Length < 9)
        {
            return null;
        }

        var units = _units.Value;
        if (units.Count == 0)
        {
            return null;
        }

        var facultyCode = digits[..2] + "000";
        var departmentPrefix = digits[..4];
        var educationMarker = digits[8];
        var department = FindDepartment(units, departmentPrefix, educationMarker);

        AcademicUnit? faculty = null;
        if (department is not null && !string.IsNullOrWhiteSpace(department.ParentId))
        {
            units.TryGetValue(department.ParentId, out faculty);
        }

        if (faculty is null)
        {
            units.TryGetValue(facultyCode, out faculty);
        }

        if (faculty is null && department is null)
        {
            return null;
        }

        return new AcademicUnitResolution(faculty?.Name, department?.Name);
    }

    private AcademicUnit? FindDepartment(
        IReadOnlyDictionary<string, AcademicUnit> units,
        string departmentPrefix,
        char educationMarker)
    {
        foreach (var code in GetDepartmentCandidates(departmentPrefix, educationMarker))
        {
            if (units.TryGetValue(code, out var unit) && unit.ParentId != "0")
            {
                return unit;
            }
        }

        var matchingUnits = units.Values
            .Where(unit => unit.Id.StartsWith(departmentPrefix, StringComparison.Ordinal) && unit.ParentId != "0")
            .ToList();

        if (matchingUnits.Count == 0)
        {
            return null;
        }

        var wantsSecondTeaching = educationMarker is '3' or '4';
        return matchingUnits.FirstOrDefault(unit => IsSecondTeaching(unit.Name) == wantsSecondTeaching)
            ?? matchingUnits.FirstOrDefault();
    }

    private static IEnumerable<string> GetDepartmentCandidates(string departmentPrefix, char educationMarker)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);

        foreach (var suffix in GetSuffixCandidates(educationMarker))
        {
            var code = departmentPrefix + suffix;
            if (seen.Add(code))
            {
                yield return code;
            }
        }
    }

    private static IEnumerable<char> GetSuffixCandidates(char educationMarker)
    {
        yield return educationMarker;

        if (educationMarker is '1' or '2')
        {
            yield return '1';
            yield return '2';
        }
        else if (educationMarker is '3' or '4')
        {
            yield return '3';
            yield return '4';
        }
    }

    private IReadOnlyDictionary<string, AcademicUnit> LoadUnits()
    {
        var path = Path.Combine(_environment.ContentRootPath, "ReferenceData", "AcademicUnits.xlsx");
        if (!File.Exists(path))
        {
            return new Dictionary<string, AcademicUnit>(StringComparer.Ordinal);
        }

        using var workbook = new XLWorkbook(path);
        var worksheet = workbook.Worksheets.FirstOrDefault();
        if (worksheet is null)
        {
            return new Dictionary<string, AcademicUnit>(StringComparer.Ordinal);
        }

        var header = FindHeader(worksheet);
        if (header is null)
        {
            return new Dictionary<string, AcademicUnit>(StringComparer.Ordinal);
        }

        var units = new Dictionary<string, AcademicUnit>(StringComparer.Ordinal);
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? header.Value.RowNumber;
        for (var rowNumber = header.Value.RowNumber + 1; rowNumber <= lastRow; rowNumber++)
        {
            var id = worksheet.Cell(rowNumber, header.Value.IdColumn).GetFormattedString().Trim();
            var name = NormalizeWhitespace(worksheet.Cell(rowNumber, header.Value.NameColumn).GetFormattedString());
            var parentId = worksheet.Cell(rowNumber, header.Value.ParentIdColumn).GetFormattedString().Trim();

            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            units[id] = new AcademicUnit(id, name, parentId);
        }

        return units;
    }

    private static AcademicUnitHeader? FindHeader(IXLWorksheet worksheet)
    {
        var lastRow = Math.Min(20, worksheet.LastRowUsed()?.RowNumber() ?? 20);
        var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;

        for (var rowNumber = 1; rowNumber <= lastRow; rowNumber++)
        {
            var columns = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (var columnNumber = 1; columnNumber <= lastColumn; columnNumber++)
            {
                var value = NormalizeHeader(worksheet.Cell(rowNumber, columnNumber).GetFormattedString());
                if (!string.IsNullOrWhiteSpace(value) && !columns.ContainsKey(value))
                {
                    columns[value] = columnNumber;
                }
            }

            if (columns.TryGetValue("birimid", out var idColumn))
            {
                return new AcademicUnitHeader(
                    rowNumber,
                    idColumn,
                    columns.TryGetValue("birimadi", out var nameColumn) ? nameColumn : 2,
                    columns.TryGetValue("parentid", out var parentIdColumn) ? parentIdColumn : 7);
            }
        }

        return null;
    }

    private static string NormalizeHeader(string value)
    {
        var mapped = value
            .Replace('İ', 'I')
            .Replace('ı', 'i')
            .Replace('Ğ', 'G')
            .Replace('ğ', 'g')
            .Replace('Ü', 'U')
            .Replace('ü', 'u')
            .Replace('Ş', 'S')
            .Replace('ş', 's')
            .Replace('Ö', 'O')
            .Replace('ö', 'o')
            .Replace('Ç', 'C')
            .Replace('ç', 'c');

        var normalized = mapped.Normalize(NormalizationForm.FormD);
        var builder = new StringBuilder(normalized.Length);
        foreach (var character in normalized)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(character) == UnicodeCategory.NonSpacingMark)
            {
                continue;
            }

            if (char.IsLetterOrDigit(character))
            {
                builder.Append(char.ToLowerInvariant(character));
            }
        }

        return builder.ToString();
    }

    private static bool IsSecondTeaching(string value)
    {
        var normalized = NormalizeHeader(value);
        return normalized.Contains("ikinci", StringComparison.Ordinal)
            || normalized.Contains("iiogretim", StringComparison.Ordinal);
    }

    private static string NormalizeWhitespace(string value) =>
        string.Join(' ', value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    private sealed record AcademicUnit(string Id, string Name, string ParentId);

    private readonly record struct AcademicUnitHeader(
        int RowNumber,
        int IdColumn,
        int NameColumn,
        int ParentIdColumn);
}

public sealed record AcademicUnitResolution(string? FacultyName, string? DepartmentName);
