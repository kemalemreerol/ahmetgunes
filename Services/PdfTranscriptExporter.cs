using Microsoft.AspNetCore.Hosting;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using TranscriptApp.Models;

namespace TranscriptApp.Services;

public sealed class PdfTranscriptExporter
{
    private readonly string _logoPath;

    public PdfTranscriptExporter(IWebHostEnvironment environment)
    {
        _logoPath = Path.Combine(environment.WebRootPath, "images", "site-logo-outlined.png");
    }

    public byte[] Export(TranscriptReport report)
    {
        QuestPDF.Settings.License = LicenseType.Community;

        return Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.A4);
                page.Margin(28);
                page.DefaultTextStyle(style => style.FontFamily("Arial").FontSize(9));

                page.Header().Element(header => ComposeHeader(header, report));
                page.Content().PaddingTop(12).Element(content => ComposeContent(content, report));
                page.Footer().AlignCenter().Text(text =>
                {
                    text.Span("Sayfa ");
                    text.CurrentPageNumber();
                    text.Span(" / ");
                    text.TotalPages();
                });
            });
        }).GeneratePdf();
    }

    private void ComposeHeader(IContainer container, TranscriptReport report)
    {
        container.Column(column =>
        {
            column.Item().Row(row =>
            {
                row.ConstantItem(72).Height(68).AlignMiddle().Element(logo =>
                {
                    if (File.Exists(_logoPath))
                    {
                        logo.Width(58).Image(_logoPath).FitWidth();
                    }
                });
                row.RelativeItem().AlignMiddle().Text("ESKİŞEHİR OSMANGAZİ ÜNİVERSİTESİ").FontSize(18).Bold();
            });
            column.Item().PaddingTop(4).AlignRight().Text(DateTime.Now.ToString("dd.MM.yyyy HH:mm"));

            column.Item().PaddingTop(4).Element(info => ComposeStudentInfo(info, report));
            column.Item().PaddingTop(4).Text(text =>
            {
                text.Span("Hesaplama Türü: ").SemiBold();
                text.Span(CalculationBasisLabel(report.CalculationBasis));
                text.Span("   Kapsam: ").SemiBold();
                text.Span(ScopeLabel(report.Scope));
                text.Span("   Genel Ortalama: ").SemiBold();
                text.Span(AverageText(report.OverallAverage));
                text.Span("   Toplam Krd: ").SemiBold();
                text.Span(NumberText(report.TotalCredit));
                text.Span("   Toplam AKTS: ").SemiBold();
                text.Span(NumberText(report.TotalEcts));
            });

            column.Item().PaddingTop(8).LineHorizontal(1).LineColor(Colors.Grey.Lighten1);
        });
    }

    private static void ComposeStudentInfo(IContainer container, TranscriptReport report)
    {
        container.Width(360).Table(table =>
        {
            table.ColumnsDefinition(columns =>
            {
                columns.ConstantColumn(130);
                columns.ConstantColumn(10);
                columns.RelativeColumn();
            });

            InfoRow(table, "Öğrenci No", DisplayValue(report.StudentNumber));
            InfoRow(table, "Öğrenci Adı Soyadı", DisplayValue(report.StudentFullName));
            InfoRow(table, "Öğrenci TC/YU Numarası", DisplayValue(report.StudentIdentityNumber));
            InfoRow(table, "Fakülte", DisplayValue(report.FacultyName));
            InfoRow(table, "Bölüm", DisplayValue(report.ProgramCode));
        });
    }

    private static void InfoRow(TableDescriptor table, string label, string value)
    {
        table.Cell().Text(label).SemiBold();
        table.Cell().Text(":").SemiBold();
        table.Cell().Text(value);
    }

    private static void ComposeContent(IContainer container, TranscriptReport report)
    {
        container.Column(column =>
        {
            column.Spacing(10);

            foreach (var group in report.GroupedCourses)
            {
                column.Item().PaddingTop(5).Text($"{group.Key.Year} / {group.Key.Semester}").FontSize(12).Bold();
                column.Item().Text($"Dönem Ortalaması: {AverageText(report.GetAverage(group.Key)?.Average)}");
                column.Item().Element(table => ComposeCourseTable(table, group));
                if (report.GetAverage(group.Key) is { } periodAverage)
                {
                    column.Item().AlignRight().Text($"Toplam Krd: {NumberText(periodAverage.TotalCredit)}   Toplam AKTS: {NumberText(periodAverage.TotalEcts)}").SemiBold();
                }
            }

            column.Item().PaddingTop(8).AlignRight().Text(
                $"Genel Ortalama: {AverageText(report.OverallAverage)}   Toplam Krd: {NumberText(report.TotalCredit)}   Toplam AKTS: {NumberText(report.TotalEcts)}").FontSize(11).Bold();
        });
    }

    private static void ComposeCourseTable(IContainer container, IEnumerable<TranscriptCourse> courses)
    {
        container.Table(table =>
        {
            table.ColumnsDefinition(columns =>
            {
                columns.ConstantColumn(36);
                columns.ConstantColumn(38);
                columns.ConstantColumn(62);
                columns.RelativeColumn(3);
                columns.ConstantColumn(34);
                columns.ConstantColumn(38);
                columns.ConstantColumn(48);
            });

            table.Header(header =>
            {
                HeaderCell(header, "Yıl");
                HeaderCell(header, "Yarıyıl");
                HeaderCell(header, "Ders Kodu");
                HeaderCell(header, "Ders Adı");
                HeaderCell(header, "Krd");
                HeaderCell(header, "AKTS");
                HeaderCell(header, "Harf Notu");
            });

            foreach (var course in courses)
            {
                BodyCell(table, course.Year);
                BodyCell(table, course.Semester);
                BodyCell(table, course.CourseCode);
                BodyCell(table, course.CourseName);
                BodyCell(table, course.Credit);
                BodyCell(table, course.Ects);
                BodyCell(table, course.LetterGrade);
            }
        });
    }

    private static void HeaderCell(TableCellDescriptor header, string text) =>
        header.Cell().Element(HeaderStyle).Text(text).SemiBold();

    private static void BodyCell(TableDescriptor table, string text) =>
        table.Cell().Element(BodyStyle).Text(text);

    private static IContainer HeaderStyle(IContainer container) =>
        container
            .Border(0.5f)
            .BorderColor(Colors.Grey.Lighten1)
            .Background(Colors.Grey.Lighten3)
            .PaddingVertical(4)
            .PaddingHorizontal(5);

    private static IContainer BodyStyle(IContainer container) =>
        container
            .Border(0.5f)
            .BorderColor(Colors.Grey.Lighten2)
            .PaddingVertical(3)
            .PaddingHorizontal(5);

    private static string DisplayValue(string? value) =>
        string.IsNullOrWhiteSpace(value) ? "-" : value;

    private static string CalculationBasisLabel(TranscriptCalculationBasis calculationBasis) =>
        calculationBasis == TranscriptCalculationBasis.Credit ? "Krd" : "AKTS";

    private static string ScopeLabel(TranscriptScope scope) =>
        scope == TranscriptScope.FirstTwoYears ? "İlk iki yıl" : "Tüm dersler";

    private static string AverageText(decimal? average) =>
        average.HasValue ? average.Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture) : "-";

    private static string NumberText(decimal number) =>
        number.ToString("0.############################", System.Globalization.CultureInfo.InvariantCulture);
}
