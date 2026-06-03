using DocConverterService.Models;
using DocConverterService.Services;
using Microsoft.Extensions.Options;

if (!OperatingSystem.IsWindows())
{
    throw new PlatformNotSupportedException("DocConverterService chỉ hỗ trợ chạy trên Windows có cài Microsoft Word.");
}

var builder = WebApplication.CreateBuilder(args);

builder.Services.Configure<ConverterOptions>(builder.Configuration.GetSection("Converter"));
builder.Services.AddSingleton<IWordInteropConverter, WordInteropConverter>();
builder.Services.AddHealthChecks();

var app = builder.Build();

app.MapGet("/health", () => Results.Ok(new
{
    status = "ok",
    service = "doc-converter-service",
    platform = Environment.OSVersion.VersionString
}));

app.MapPost("/api/convert-doc", async (
    HttpRequest request,
    IWordInteropConverter converter,
    IOptions<ConverterOptions> options,
    ILogger<Program> logger,
    CancellationToken cancellationToken) =>
{
    if (!request.HasFormContentType)
    {
        return Results.BadRequest(new { error = "Request phải là multipart/form-data." });
    }

    var form = await request.ReadFormAsync(cancellationToken);
    var file = form.Files["file"];
    if (file is null)
    {
        return Results.BadRequest(new { error = "Vui lòng gửi file .doc trong field 'file'." });
    }

    if (file.Length <= 0)
    {
        return Results.BadRequest(new { error = "File upload rỗng." });
    }

    var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
    if (!string.Equals(extension, ".doc", StringComparison.OrdinalIgnoreCase))
    {
        return Results.BadRequest(new { error = "Service chỉ chấp nhận file .doc." });
    }

    var requestId = Guid.NewGuid().ToString("N");
    var tempRoot = Path.GetFullPath(options.Value.TempRoot ?? "temp");
    var requestDirectory = Path.Combine(tempRoot, requestId);
    Directory.CreateDirectory(requestDirectory);

    var inputFileName = $"{Path.GetFileNameWithoutExtension(file.FileName)}{extension}";
    var inputPath = Path.Combine(requestDirectory, inputFileName);

    try
    {
        await using (var stream = File.Create(inputPath))
        {
            await file.CopyToAsync(stream, cancellationToken);
        }

        var conversion = await converter.ConvertDocToDocxAsync(
            inputPath,
            requestDirectory,
            file.FileName,
            cancellationToken);

        if (!File.Exists(conversion.OutputPath))
        {
            throw new WordConversionException("Không tìm thấy file .docx đầu ra sau khi chuyển đổi.");
        }

        var bytes = await File.ReadAllBytesAsync(conversion.OutputPath, cancellationToken);
        return Results.File(
            bytes,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            conversion.OutputFileName);
    }
    catch (WordConversionTimeoutException ex)
    {
        logger.LogError(ex, "Timeout while converting file {FileName}.", file.FileName);
        return Results.Json(new { error = "Xử lý quá thời gian cho phép" }, statusCode: StatusCodes.Status500InternalServerError);
    }
    catch (WordConversionException ex)
    {
        logger.LogError(ex, "Word conversion failed for {FileName}.", file.FileName);
        return Results.Json(new { error = ex.Message }, statusCode: StatusCodes.Status500InternalServerError);
    }
    catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
    {
        logger.LogWarning("Request cancelled while converting file {FileName}.", file.FileName);
        return Results.Json(new { error = "Yêu cầu đã bị hủy." }, statusCode: 499);
    }
    catch (Exception ex)
    {
        logger.LogError(ex, "Unexpected error while converting file {FileName}.", file.FileName);
        return Results.Json(new { error = "Đã xảy ra lỗi trong quá trình chuyển đổi file." }, statusCode: StatusCodes.Status500InternalServerError);
    }
    finally
    {
        try
        {
            if (Directory.Exists(requestDirectory))
            {
                Directory.Delete(requestDirectory, recursive: true);
            }
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "Unable to delete temp directory {RequestDirectory}.", requestDirectory);
        }
    }
});

app.Run();
