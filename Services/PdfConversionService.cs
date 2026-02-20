using System.Collections.Concurrent;
using System.Diagnostics;

namespace DecoSOP.Services;

/// <summary>
/// Converts Office documents (DOC, DOCX, XLSX, XLS) to PDF using LibreOffice.
/// Caches converted PDFs in a previews directory so conversion only happens once.
/// </summary>
public static class PdfConversionService
{
    private static readonly HashSet<string> ConvertibleExtensions =
        [".doc", ".docx", ".xlsx", ".xls", ".ppt", ".pptx", ".odt", ".ods", ".odp", ".rtf"];

    // Track in-progress conversions to avoid duplicate work
    private static readonly ConcurrentDictionary<string, Task<string?>> ActiveConversions = new();

    private static string? _sofficePath;

    public static bool CanConvert(string fileName)
    {
        var ext = Path.GetExtension(fileName).ToLowerInvariant();
        return ConvertibleExtensions.Contains(ext);
    }

    /// <summary>
    /// Get or create a PDF preview for the given source file.
    /// Returns the path to the cached PDF, or null if conversion fails.
    /// </summary>
    public static async Task<string?> GetOrCreatePdfAsync(string sourceFilePath, string storedFileName)
    {
        var previewDir = GetPreviewDirectory();
        var pdfFileName = Path.ChangeExtension(storedFileName, ".pdf");
        var pdfPath = Path.Combine(previewDir, pdfFileName);

        // Return cached PDF if it exists and is newer than the source
        if (File.Exists(pdfPath))
        {
            var sourceTime = File.GetLastWriteTimeUtc(sourceFilePath);
            var pdfTime = File.GetLastWriteTimeUtc(pdfPath);
            if (pdfTime >= sourceTime)
                return pdfPath;
        }

        // Use ConcurrentDictionary to ensure only one conversion per file at a time
        var task = ActiveConversions.GetOrAdd(pdfPath, _ => ConvertToPdfAsync(sourceFilePath, previewDir, pdfPath));

        try
        {
            return await task;
        }
        finally
        {
            ActiveConversions.TryRemove(pdfPath, out _);
        }
    }

    private static async Task<string?> ConvertToPdfAsync(string sourceFilePath, string previewDir, string expectedPdfPath)
    {
        var soffice = FindSoffice();
        if (soffice is null)
            return null;

        try
        {
            // LibreOffice needs a unique user profile when running multiple instances
            var profileDir = Path.Combine(Path.GetTempPath(), $"libreoffice_convert_{Guid.NewGuid():N}");

            var psi = new ProcessStartInfo
            {
                FileName = soffice,
                Arguments = $"--headless --norestore --convert-to pdf --outdir \"{previewDir}\" \"{sourceFilePath}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                Environment =
                {
                    ["UserInstallation"] = $"file:///{profileDir.Replace('\\', '/')}"
                }
            };

            using var process = Process.Start(psi);
            if (process is null)
                return null;

            // Wait up to 60 seconds for conversion
            var completed = await WaitForExitAsync(process, TimeSpan.FromSeconds(60));

            // Clean up temp profile
            try { if (Directory.Exists(profileDir)) Directory.Delete(profileDir, true); } catch { }

            if (!completed)
            {
                try { process.Kill(); } catch { }
                return null;
            }

            // LibreOffice names the output based on the source filename
            // The source file is stored as "123_filename.doc" and output becomes "123_filename.pdf"
            var sourceNameWithoutExt = Path.GetFileNameWithoutExtension(sourceFilePath);
            var generatedPdfPath = Path.Combine(previewDir, sourceNameWithoutExt + ".pdf");

            if (File.Exists(generatedPdfPath))
            {
                // If the generated name differs from expected, rename it
                if (generatedPdfPath != expectedPdfPath && !File.Exists(expectedPdfPath))
                    File.Move(generatedPdfPath, expectedPdfPath);

                return File.Exists(expectedPdfPath) ? expectedPdfPath : generatedPdfPath;
            }

            // Check if the expected path exists (in case names matched)
            return File.Exists(expectedPdfPath) ? expectedPdfPath : null;
        }
        catch
        {
            return null;
        }
    }

    private static async Task<bool> WaitForExitAsync(Process process, TimeSpan timeout)
    {
        using var cts = new CancellationTokenSource(timeout);
        try
        {
            await process.WaitForExitAsync(cts.Token);
            return true;
        }
        catch (OperationCanceledException)
        {
            return false;
        }
    }

    private static string? FindSoffice()
    {
        if (_sofficePath is not null)
            return File.Exists(_sofficePath) ? _sofficePath : null;

        // Check common Windows install locations
        var candidates = new[]
        {
            @"C:\Program Files\LibreOffice\program\soffice.exe",
            @"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        };

        foreach (var path in candidates)
        {
            if (File.Exists(path))
            {
                _sofficePath = path;
                return path;
            }
        }

        // Try PATH
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "where",
                Arguments = "soffice",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };
            using var process = Process.Start(psi);
            if (process is not null)
            {
                var output = process.StandardOutput.ReadToEnd().Trim();
                process.WaitForExit(5000);
                if (!string.IsNullOrEmpty(output))
                {
                    var firstLine = output.Split('\n')[0].Trim();
                    if (File.Exists(firstLine))
                    {
                        _sofficePath = firstLine;
                        return firstLine;
                    }
                }
            }
        }
        catch { }

        return null;
    }

    public static bool IsLibreOfficeAvailable() => FindSoffice() is not null;

    private static string GetPreviewDirectory()
    {
        var dir = Path.Combine(DocumentService.DataDirectory, "uploads", "previews");
        Directory.CreateDirectory(dir);
        return dir;
    }
}
