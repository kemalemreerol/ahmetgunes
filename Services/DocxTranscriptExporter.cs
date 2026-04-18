using Microsoft.AspNetCore.Hosting;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TranscriptApp.Models;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace TranscriptApp.Services;

public sealed class DocxTranscriptExporter
{
    private readonly string _logoPath;

    public DocxTranscriptExporter(IWebHostEnvironment environment)
    {
        _logoPath = Path.Combine(environment.WebRootPath, "images", "site-logo-outlined.png");
    }

    public byte[] Export(TranscriptReport report)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var body = mainPart.Document.Body!;
            body.Append(CreateHeaderTable(mainPart, _logoPath));
            body.Append(CreateParagraph($"Oluşturma Tarihi: {DateTime.Now:dd.MM.yyyy HH:mm}"));
            body.Append(CreateStudentInfoTable(report));
            body.Append(CreateParagraph($"Hesaplama Türü: {CalculationBasisLabel(report.CalculationBasis)}   Kapsam: {ScopeLabel(report.Scope)}   Genel Ortalama: {AverageText(report.OverallAverage)}   Toplam Krd: {NumberText(report.TotalCredit)}   Toplam AKTS: {NumberText(report.TotalEcts)}"));

            foreach (var group in report.GroupedCourses)
            {
                body.Append(CreateParagraph($"{group.Key.Year} / {group.Key.Semester}", bold: true, fontSize: "24"));
                body.Append(CreateParagraph($"Dönem Ortalaması: {AverageText(report.GetAverage(group.Key)?.Average)}"));
                body.Append(CreateCourseTable(group));
                if (report.GetAverage(group.Key) is { } periodAverage)
                {
                    body.Append(CreateParagraph($"Toplam Krd: {NumberText(periodAverage.TotalCredit)}   Toplam AKTS: {NumberText(periodAverage.TotalEcts)}", bold: true, justification: JustificationValues.Right));
                }
            }

            body.Append(CreateParagraph($"Genel Ortalama: {AverageText(report.OverallAverage)}   Toplam Krd: {NumberText(report.TotalCredit)}   Toplam AKTS: {NumberText(report.TotalEcts)}", bold: true, fontSize: "24", justification: JustificationValues.Right));

            body.Append(new SectionProperties(
                new PageSize { Width = 11906, Height = 16838 },
                new PageMargin { Top = 720, Right = 540, Bottom = 720, Left = 540, Header = 360, Footer = 360, Gutter = 0 }));

            mainPart.Document.Save();
        }

        return stream.ToArray();
    }

    private static Paragraph CreateParagraph(
        string text,
        bool bold = false,
        string fontSize = "20",
        JustificationValues? justification = null)
    {
        var runProperties = new RunProperties(new FontSize { Val = fontSize });
        if (bold)
        {
            runProperties.Append(new Bold());
        }

        var paragraphProperties = new ParagraphProperties(new SpacingBetweenLines { After = "120" });
        if (justification.HasValue)
        {
            paragraphProperties.Append(new Justification { Val = justification.Value });
        }

        return new Paragraph(
            paragraphProperties,
            new Run(runProperties, new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static Table CreateHeaderTable(MainDocumentPart mainPart, string imagePath)
    {
        var table = new Table(
            new TableProperties(
                new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                CreateNoBorders()));

        var row = new TableRow();
        var logoCell = CreateCell(string.Empty, false, "950");
        if (File.Exists(imagePath))
        {
            logoCell.RemoveAllChildren<Paragraph>();
            logoCell.Append(CreateLogoParagraph(mainPart, imagePath));
        }

        var titleCell = CreateCell(string.Empty, false, "9250");
        titleCell.RemoveAllChildren<Paragraph>();
        titleCell.Append(CreateParagraph("ESKİŞEHİR OSMANGAZİ ÜNİVERSİTESİ", bold: true, fontSize: "32"));

        row.Append(logoCell, titleCell);
        table.Append(row);
        return table;
    }

    private static Paragraph CreateLogoParagraph(MainDocumentPart mainPart, string imagePath)
    {
        var imagePart = mainPart.AddImagePart(ImagePartType.Png);
        using (var stream = File.OpenRead(imagePath))
        {
            imagePart.FeedData(stream);
        }

        var relationshipId = mainPart.GetIdOfPart(imagePart);
        const long imageSizeEmu = 731520;

        var drawing = new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = imageSizeEmu, Cy = imageSizeEmu },
                new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new DW.DocProperties { Id = 1U, Name = "Üniversite Logosu" },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = "site-logo-outlined.png" },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0, Y = 0 },
                                    new A.Extents { Cx = imageSizeEmu, Cy = imageSizeEmu }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            });

        return new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { After = "120" }),
            new Run(drawing));
    }

    private static Table CreateStudentInfoTable(TranscriptReport report)
    {
        var table = new Table(
            new TableProperties(
                new TableWidth { Width = "5200", Type = TableWidthUnitValues.Dxa },
                CreateNoBorders()));

        table.Append(CreateInfoRow("Öğrenci No", DisplayValue(report.StudentNumber)));
        table.Append(CreateInfoRow("Öğrenci Adı Soyadı", DisplayValue(report.StudentFullName)));
        table.Append(CreateInfoRow("Öğrenci TC/YU Numarası", DisplayValue(report.StudentIdentityNumber)));
        table.Append(CreateInfoRow("Fakülte", DisplayValue(report.FacultyName)));
        table.Append(CreateInfoRow("Bölüm", DisplayValue(report.ProgramCode)));

        return table;
    }

    private static TableRow CreateInfoRow(string label, string value)
    {
        var row = new TableRow();
        row.Append(
            CreateInfoCell(label, "2400", bold: true),
            CreateInfoCell(":", "200", bold: true),
            CreateInfoCell(value, "2600"));
        return row;
    }

    private static Table CreateCourseTable(IEnumerable<TranscriptCourse> courses)
    {
        var table = new Table(
            new TableProperties(
                new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4 },
                    new BottomBorder { Val = BorderValues.Single, Size = 4 },
                    new LeftBorder { Val = BorderValues.Single, Size = 4 },
                    new RightBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 })));

        var widths = new[] { "650", "650", "1100", "5600", "500", "650", "850" };
        table.Append(CreateRow(new[] { "Yıl", "Yarıyıl", "Ders Kodu", "Ders Adı", "Krd", "AKTS", "Harf Notu" }, widths, isHeader: true));

        foreach (var course in courses)
        {
            table.Append(CreateRow(new[]
            {
                course.Year,
                course.Semester,
                course.CourseCode,
                course.CourseName,
                course.Credit,
                course.Ects,
                course.LetterGrade
            }, widths));
        }

        return table;
    }

    private static TableRow CreateRow(IEnumerable<string> values, IReadOnlyList<string>? widths = null, bool isHeader = false)
    {
        var row = new TableRow();
        var index = 0;
        foreach (var value in values)
        {
            row.Append(CreateCell(value, isHeader, widths is not null && index < widths.Count ? widths[index] : null));
            index++;
        }

        return row;
    }

    private static TableCell CreateCell(string value, bool isHeader, string? width = null)
    {
        var properties = new TableCellProperties(
            new TableCellMargin(
                new TopMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                new BottomMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                new LeftMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                new RightMargin { Width = "80", Type = TableWidthUnitValues.Dxa }));

        if (!string.IsNullOrWhiteSpace(width))
        {
            properties.Append(new TableCellWidth { Width = width, Type = TableWidthUnitValues.Dxa });
        }

        if (isHeader)
        {
            properties.Append(new Shading { Fill = "E9ECEF" });
        }

        return new TableCell(properties, CreateParagraph(value, isHeader));
    }

    private static TableCell CreateInfoCell(string value, string width, bool bold = false)
    {
        var properties = new TableCellProperties(
            new TableCellWidth { Width = width, Type = TableWidthUnitValues.Dxa },
            new TableCellMargin(
                new TopMargin { Width = "20", Type = TableWidthUnitValues.Dxa },
                new BottomMargin { Width = "20", Type = TableWidthUnitValues.Dxa },
                new LeftMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                new RightMargin { Width = "40", Type = TableWidthUnitValues.Dxa }));

        return new TableCell(properties, CreateParagraph(value, bold));
    }

    private static TableBorders CreateNoBorders() =>
        new(
            new TopBorder { Val = BorderValues.None },
            new BottomBorder { Val = BorderValues.None },
            new LeftBorder { Val = BorderValues.None },
            new RightBorder { Val = BorderValues.None },
            new InsideHorizontalBorder { Val = BorderValues.None },
            new InsideVerticalBorder { Val = BorderValues.None });

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
