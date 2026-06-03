namespace DocConverterService.Services;

public sealed class WordConversionException : Exception
{
    public WordConversionException(string message) : base(message)
    {
    }

    public WordConversionException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
