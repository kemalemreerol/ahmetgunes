using Microsoft.AspNetCore.Http;

namespace TranscriptApp.Services;

public sealed class UploadedWorkbookStore
{
    private readonly string _uploadRoot;

    public UploadedWorkbookStore(IWebHostEnvironment environment)
    {
        _uploadRoot = Path.Combine(environment.ContentRootPath, "App_Data", "uploads");
    }

    public async Task<StoredWorkbook> SaveAsync(IFormFile workbook, CancellationToken cancellationToken)
    {
        if (workbook.Length == 0)
        {
            throw new TranscriptProcessingException("Yuklenen Excel dosyasi bos gorunuyor.");
        }

        if (!string.Equals(Path.GetExtension(workbook.FileName), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            throw new TranscriptProcessingException("Lutfen .xlsx uzantili bir Excel dosyasi yukleyin.");
        }

        CleanupOldUploads();

        Directory.CreateDirectory(_uploadRoot);

        var uploadId = Guid.NewGuid().ToString("N");
        var filePath = Path.Combine(_uploadRoot, $"{uploadId}.xlsx");

        await using var output = File.Create(filePath);
        await workbook.CopyToAsync(output, cancellationToken);

        return new StoredWorkbook(uploadId, Path.GetFileName(workbook.FileName), filePath);
    }

    public StoredWorkbook Get(string uploadId)
    {
        if (string.IsNullOrWhiteSpace(uploadId) || uploadId.Any(ch => !Uri.IsHexDigit(ch)))
        {
            throw new TranscriptProcessingException("Yuklenen dosya oturumu bulunamadi. Lutfen dosyayi yeniden yukleyin.");
        }

        var filePath = Path.Combine(_uploadRoot, $"{uploadId}.xlsx");
        if (!File.Exists(filePath))
        {
            throw new TranscriptProcessingException("Yuklenen dosya bulunamadi. Lutfen dosyayi yeniden yukleyin.");
        }

        return new StoredWorkbook(uploadId, "transkript.xlsx", filePath);
    }

    private void CleanupOldUploads()
    {
        if (!Directory.Exists(_uploadRoot))
        {
            return;
        }

        var cutoff = DateTimeOffset.UtcNow.AddHours(-6);
        foreach (var file in Directory.EnumerateFiles(_uploadRoot, "*.xlsx"))
        {
            var info = new FileInfo(file);
            if (info.LastWriteTimeUtc < cutoff)
            {
                info.Delete();
            }
        }
    }
}
