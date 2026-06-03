namespace DocConverterService.Models;

public sealed class ConverterOptions
{
    public string TempRoot { get; set; } = "temp";

    public int TimeoutSeconds { get; set; } = 15;
}
