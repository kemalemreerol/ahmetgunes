using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using TranscriptApp.Models;
using TranscriptApp.Services;

namespace TranscriptApp.Pages;

public class IndexModel : PageModel
{
    private readonly TranscriptParser _parser;
    private readonly UploadedWorkbookStore _store;
    private readonly PdfTranscriptExporter _pdfExporter;
    private readonly DocxTranscriptExporter _docxExporter;

    public IndexModel(
        TranscriptParser parser,
        UploadedWorkbookStore store,
        PdfTranscriptExporter pdfExporter,
        DocxTranscriptExporter docxExporter)
    {
        _parser = parser;
        _store = store;
        _pdfExporter = pdfExporter;
        _docxExporter = docxExporter;
    }

    [BindProperty]
    public IFormFile? Workbook { get; set; }

    [BindProperty]
    public string? UploadId { get; set; }

    [BindProperty]
    public TranscriptCalculationBasis CalculationBasis { get; set; } = TranscriptCalculationBasis.Ects;

    [BindProperty]
    public TranscriptScope Scope { get; set; } = TranscriptScope.AllCourses;

    public TranscriptReport? Report { get; private set; }

    public string? ErrorMessage { get; private set; }

    public async Task<IActionResult> OnPostUploadAsync(CancellationToken cancellationToken)
    {
        if (Workbook is null)
        {
            ErrorMessage = "Lütfen .xlsx uzantılı bir Excel dosyası seçin.";
            return Page();
        }

        try
        {
            var stored = await _store.SaveAsync(Workbook, cancellationToken);
            await using var stream = System.IO.File.OpenRead(stored.FilePath);

            Report = _parser.Parse(stream, stored.OriginalFileName, CalculationBasis, Scope);
            UploadId = stored.UploadId;

            return Page();
        }
        catch (TranscriptProcessingException ex)
        {
            ErrorMessage = ex.Message;
            return Page();
        }
    }

    public IActionResult OnPostDownloadPdf()
    {
        return BuildDownload(
            report => _pdfExporter.Export(report),
            "pdf",
            "application/pdf");
    }

    public IActionResult OnPostDownloadDocx()
    {
        return BuildDownload(
            report => _docxExporter.Export(report),
            "docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    }

    private IActionResult BuildDownload(Func<TranscriptReport, byte[]> exporter, string extension, string contentType)
    {
        try
        {
            var stored = _store.Get(UploadId ?? string.Empty);

            using var stream = System.IO.File.OpenRead(stored.FilePath);
            var report = _parser.Parse(stream, stored.OriginalFileName, CalculationBasis, Scope);
            var content = exporter(report);
            var downloadName = TranscriptFileNameBuilder.Build(report, extension);

            return File(content, contentType, downloadName);
        }
        catch (TranscriptProcessingException ex)
        {
            ErrorMessage = ex.Message;
            return Page();
        }
    }

    public static string CalculationBasisLabel(TranscriptCalculationBasis calculationBasis) =>
        calculationBasis == TranscriptCalculationBasis.Credit ? "Krd" : "AKTS";

    public static string ScopeLabel(TranscriptScope scope) =>
        scope == TranscriptScope.FirstTwoYears ? "İlk iki yıl" : "Tüm dersler";

    public static string AverageText(decimal? average) =>
        average.HasValue ? average.Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture) : "-";

    public static string NumberText(decimal number) =>
        number.ToString("0.############################", System.Globalization.CultureInfo.InvariantCulture);
}
