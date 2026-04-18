namespace TranscriptApp.Services;

public sealed class TranscriptProcessingException : Exception
{
    public TranscriptProcessingException(string message)
        : base(message)
    {
    }

    public TranscriptProcessingException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
