namespace TranscriptApp.Models;

public sealed record TranscriptCourse(
    int SourceRow,
    string Year,
    string Semester,
    string Nyy,
    string CourseCode,
    string CourseName,
    string Credit,
    string Ects,
    string LetterGrade);
