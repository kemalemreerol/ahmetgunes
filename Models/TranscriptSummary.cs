namespace TranscriptApp.Models;

public sealed record TranscriptSummary(
    int TotalRows,
    int EmptyNotesSkipped,
    int StarredNotesSkipped,
    int InvalidNotesSkipped,
    int IncludedRows);
