namespace ReportBom;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        var status = new ProcessingStatus();
        TsvBomExporter.CleanupStaleExports(status);

        using var session = SolidEdgeSession.Connect(status);
        if (session == null)
        {
            return;
        }

        var asmFiles = InputFileResolver.GetAssemblyFiles(args, status);
        if (asmFiles == null)
        {
            return;
        }

        status.ReportStep("Elaborazione BOM...");

        try
        {
            var processor = new BomProcessor(status);
            var tsvRows = processor.ProcessFiles(session.Application, asmFiles);
            TsvBomExporter.CreateAndOpen(tsvRows, status);
        }
        catch (Exception ex)
        {
            status.ShowError($"Si è verificato un errore imprevisto:\n{ex.Message}", "Errore");
        }
    }
}
