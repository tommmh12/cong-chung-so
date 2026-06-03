namespace DocConverterService.Services;

public sealed class WordConversionTimeoutException : Exception
{
    public WordConversionTimeoutException(string message) : base(message)
    {
    }
}
