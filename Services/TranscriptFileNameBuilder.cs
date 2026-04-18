using System.Text;
using TranscriptApp.Models;

namespace TranscriptApp.Services;

public static class TranscriptFileNameBuilder
{
    public static string Build(TranscriptReport report, string extension)
    {
        var suffix = string.IsNullOrWhiteSpace(report.StudentNumber)
            ? DateTime.Now.ToString("yyyyMMdd_HHmmss")
            : report.StudentNumber;

        return $"transkript_{Sanitize(suffix)}.{extension.TrimStart('.')}";
    }

    private static string Sanitize(string value)
    {
        var builder = new StringBuilder(value.Length);
        foreach (var ch in value)
        {
            if (char.IsLetterOrDigit(ch) || ch is '-' or '_')
            {
                builder.Append(ch);
            }
        }

        return builder.Length == 0 ? "ogrenci" : builder.ToString();
    }
}
