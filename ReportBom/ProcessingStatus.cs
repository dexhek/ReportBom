using System.Windows.Forms;

namespace ReportBom;

public sealed class ProcessingStatus
{
    private readonly object _sync = new();

    public void ReportStep(string message)
    {
        lock (_sync)
        {
            Console.WriteLine(message);
        }
    }

    public void ReportAssemblyProgress(int current, int total, string asmFile)
    {
        string fileName = Path.GetFileName(asmFile);
        ReportStep($"[{current}/{total}] Elaborazione assieme: {fileName}");
    }

    public void ReportOutputReady(string outputPath)
    {
        ReportStep($"File BOM creato: {outputPath}");
    }

    public void ShowWarning(string message, string title)
    {
        lock (_sync)
        {
            Console.WriteLine(message);
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    public void ShowError(string message, string title)
    {
        lock (_sync)
        {
            Console.Error.WriteLine(message);
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
