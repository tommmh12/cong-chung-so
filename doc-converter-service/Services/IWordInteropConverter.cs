namespace DocConverterService.Services;

public interface IWordInteropConverter
{
    System.Threading.Tasks.Task<WordConversionResult> ConvertDocToDocxAsync(
        string inputPath,
        string outputDirectory,
        string originalFileName,
        CancellationToken cancellationToken = default);
}
