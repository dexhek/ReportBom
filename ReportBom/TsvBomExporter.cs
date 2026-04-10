using System.Diagnostics;
using System.Text;

namespace ReportBom;

public static class TsvBomExporter
{
    private const string ExportFilePattern = "BOM_*.tsv";
    private static readonly TimeSpan StaleExportAge = TimeSpan.FromMinutes(5);

    public static void CleanupStaleExports(ProcessingStatus status)
    {
        string tempPath = Path.GetTempPath();

        try
        {
            foreach (string filePath in Directory.EnumerateFiles(tempPath, ExportFilePattern, SearchOption.TopDirectoryOnly))
            {
                try
                {
                    var fileInfo = new FileInfo(filePath);
                    if (DateTime.UtcNow - fileInfo.LastWriteTimeUtc < StaleExportAge)
                    {
                        continue;
                    }

                    fileInfo.Delete();
                }
                catch
                {
                    // Best effort cleanup: skip locked or inaccessible files.
                }
            }
        }
        catch (Exception ex)
        {
            status.ReportStep($"Pulizia export temporanei non riuscita: {ex.Message}");
        }
    }

    public static void CreateAndOpen(IReadOnlyList<string> tsvRows, ProcessingStatus status)
    {
        if (tsvRows.Count <= 1)
        {
            return;
        }

        status.ReportStep("Creazione del file BOM in corso...");

        string tsvPath = Path.Combine(Path.GetTempPath(), $"BOM_{Guid.NewGuid():N}.tsv");
        File.WriteAllLines(tsvPath, tsvRows, Encoding.UTF8);

        var viewerProcess = Process.Start(new ProcessStartInfo(tsvPath) { UseShellExecute = true });
        ScheduleCleanup(tsvPath, viewerProcess?.Id);
        status.ReportOutputReady(tsvPath);
    }

    private static void ScheduleCleanup(string tsvPath, int? viewerProcessId)
    {
        string escapedPath = EscapeForPowerShellSingleQuotedString(tsvPath);
        string processCondition = viewerProcessId is int pid && pid > 0
            ? $"$p = Get-Process -Id {pid} -ErrorAction SilentlyContinue; if ($p) {{ Wait-Process -Id {pid} -Timeout 14400 -ErrorAction SilentlyContinue | Out-Null }}"
            : string.Empty;

        string cleanupScript =
            "$path = '" + escapedPath + "'; " +
            processCondition +
            " $deadline = (Get-Date).AddHours(12); " +
            "while ((Get-Date) -lt $deadline) { " +
            "if (-not (Test-Path -LiteralPath $path)) { exit 0 } " +
            "try { Remove-Item -LiteralPath $path -Force -ErrorAction Stop; exit 0 } " +
            "catch { Start-Sleep -Seconds 10 } " +
            "}";

        var cleanupProcess = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = $"-NoProfile -WindowStyle Hidden -Command \"{cleanupScript}\"",
            CreateNoWindow = true,
            UseShellExecute = false
        };

        Process.Start(cleanupProcess);
    }

    private static string EscapeForPowerShellSingleQuotedString(string value)
    {
        return value.Replace("'", "''", StringComparison.Ordinal);
    }
}
