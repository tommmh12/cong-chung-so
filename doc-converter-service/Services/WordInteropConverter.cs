using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using DocConverterService.Models;
using Microsoft.Extensions.Options;
using Word = Microsoft.Office.Interop.Word;

namespace DocConverterService.Services;

public sealed class WordInteropConverter : IWordInteropConverter
{
    private static readonly SemaphoreSlim ConversionLock = new(1, 1);
    private readonly ILogger<WordInteropConverter> _logger;
    private readonly ConverterOptions _options;

    public WordInteropConverter(
        ILogger<WordInteropConverter> logger,
        IOptions<ConverterOptions> options)
    {
        _logger = logger;
        _options = options.Value;
    }

    public async Task<WordConversionResult> ConvertDocToDocxAsync(
        string inputPath,
        string outputDirectory,
        string originalFileName,
        CancellationToken cancellationToken = default)
    {
        await ConversionLock.WaitAsync(cancellationToken);
        try
        {
            Directory.CreateDirectory(outputDirectory);

            var outputFileName = $"{Path.GetFileNameWithoutExtension(originalFileName)}.docx";
            var outputPath = Path.Combine(outputDirectory, outputFileName);
            var timeout = TimeSpan.FromSeconds(_options.TimeoutSeconds <= 0 ? 15 : _options.TimeoutSeconds);

            var execution = new WordConversionExecution(inputPath, outputPath, _logger);
            var conversionTask = execution.RunAsync();
            var completedTask = await Task.WhenAny(
                conversionTask,
                Task.Delay(timeout, CancellationToken.None));

            if (completedTask != conversionTask)
            {
                execution.TryKillWordProcess();
                await Task.WhenAny(
                    conversionTask,
                    Task.Delay(TimeSpan.FromSeconds(5), CancellationToken.None));
                throw new WordConversionTimeoutException("Xử lý quá thời gian cho phép");
            }

            await conversionTask;
            return new WordConversionResult(outputPath, outputFileName);
        }
        finally
        {
            ConversionLock.Release();
        }
    }

    private sealed class WordConversionExecution
    {
        private readonly string _inputPath;
        private readonly string _outputPath;
        private readonly ILogger _logger;
        private readonly TaskCompletionSource _completion = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _processId;

        public WordConversionExecution(string inputPath, string outputPath, ILogger logger)
        {
            _inputPath = inputPath;
            _outputPath = outputPath;
            _logger = logger;
        }

        public Task RunAsync()
        {
            var thread = new Thread(RunCore)
            {
                IsBackground = true,
                Name = $"word-converter-{Guid.NewGuid():N}"
            };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            return _completion.Task;
        }

        public void TryKillWordProcess()
        {
            var processId = Volatile.Read(ref _processId);
            if (processId <= 0)
            {
                return;
            }

            try
            {
                using var process = Process.GetProcessById(processId);
                if (!process.HasExited)
                {
                    _logger.LogWarning("Force killing WINWORD.EXE process {ProcessId} after timeout.", processId);
                    process.Kill(true);
                    process.WaitForExit(5000);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Unable to kill WINWORD.EXE process {ProcessId}.", processId);
            }
        }

        private void RunCore()
        {
            Word.Application? wordApp = null;
            Word.Document? document = null;
            Exception? conversionException = null;

            try
            {
                var existingProcessIds = SnapshotWordProcessIds();

                wordApp = new Word.Application
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                };

                Volatile.Write(ref _processId, TryResolveProcessId(existingProcessIds));

                object inputPath = _inputPath;
                object readOnly = true;
                object isVisible = false;
                object missing = Type.Missing;

                document = wordApp.Documents.Open(
                    ref inputPath,
                    ref missing,
                    ref readOnly,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref isVisible,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing);

                object outputPath = _outputPath;
                object fileFormat = Word.WdSaveFormat.wdFormatXMLDocument;

                document.SaveAs2(ref outputPath, ref fileFormat);
            }
            catch (Exception ex)
            {
                conversionException = new WordConversionException(
                    "Không thể chuyển file .doc sang .docx bằng Microsoft Word.",
                    ex);
            }
            finally
            {
                CloseDocument(document);
                QuitWord(wordApp);
                ReleaseComObject(document);
                ReleaseComObject(wordApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (conversionException is not null)
                {
                    _completion.TrySetException(conversionException);
                }
                else
                {
                    _completion.TrySetResult();
                }
            }
        }

        private static HashSet<int> SnapshotWordProcessIds()
        {
            return Process
                .GetProcessesByName("WINWORD")
                .Select(process => process.Id)
                .ToHashSet();
        }

        private static int TryResolveProcessId(HashSet<int> existingProcessIds)
        {
            var deadline = DateTime.UtcNow.AddSeconds(2);

            while (DateTime.UtcNow < deadline)
            {
                var newProcess = Process
                    .GetProcessesByName("WINWORD")
                    .Where(process => !existingProcessIds.Contains(process.Id))
                    .OrderByDescending(process =>
                    {
                        try
                        {
                            return process.StartTime;
                        }
                        catch
                        {
                            return DateTime.MinValue;
                        }
                    })
                    .FirstOrDefault();

                if (newProcess is not null)
                {
                    return newProcess.Id;
                }

                Thread.Sleep(100);
            }

            return 0;
        }

        private static void CloseDocument(Word.Document? document)
        {
            if (document is null)
            {
                return;
            }

            try
            {
                object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                object originalFormat = Type.Missing;
                object routeDocument = Type.Missing;
                document.Close(ref saveChanges, ref originalFormat, ref routeDocument);
            }
            catch
            {
            }
        }

        private static void QuitWord(Word.Application? wordApp)
        {
            if (wordApp is null)
            {
                return;
            }

            try
            {
                object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                object originalFormat = Type.Missing;
                object routeDocument = Type.Missing;
                wordApp.Quit(ref saveChanges, ref originalFormat, ref routeDocument);
            }
            catch
            {
            }
        }

        private static void ReleaseComObject(object? comObject)
        {
            if (comObject is null)
            {
                return;
            }

            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    while (Marshal.ReleaseComObject(comObject) > 0)
                    {
                    }
                }
            }
            catch
            {
            }
        }
    }
}
