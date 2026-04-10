namespace ReportBom;

public static class InputFileResolver
{
    public static List<string> GetAssemblyFiles(string[] args, ProcessingStatus status)
    {
        var asmFiles = new List<string>();

        if (args.Length == 1 && args[0].EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                asmFiles.AddRange(File.ReadAllLines(args[0]).Select(line => line.Trim('"')));
            }
            catch (Exception ex)
            {
                status.ShowError($"Errore durante la lettura del file {args[0]}:\n{ex.Message}", "Errore File");
                return null;
            }
        }
        else
        {
            asmFiles.AddRange(args.Where(f => f.EndsWith(".asm", StringComparison.OrdinalIgnoreCase)));
        }

        return asmFiles;
    }
}
