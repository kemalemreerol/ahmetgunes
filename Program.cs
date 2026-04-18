using TranscriptApp.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorPages();
builder.Services.AddSingleton<AcademicUnitResolver>();
builder.Services.AddSingleton<TranscriptParser>();
builder.Services.AddSingleton<UploadedWorkbookStore>();
builder.Services.AddSingleton<PdfTranscriptExporter>();
builder.Services.AddSingleton<DocxTranscriptExporter>();

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseStaticFiles();

app.UseRouting();

app.UseAuthorization();

app.MapRazorPages();

app.Run();
