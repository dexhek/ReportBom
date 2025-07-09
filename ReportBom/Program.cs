using SolidEdgeCommunity;
using SolidEdgeCommunity.Extensions;
using System.Diagnostics;

namespace ReportBom;

// Classe per gestire l'animazione dello spinner in console.
public class Spinner : IDisposable
{
    private const string Sequence = @"/-\|";
    private int _counter = 0;
    private readonly Thread _thread;
    private bool _active;

    public Spinner()
    {
        _thread = new Thread(Spin);
    }

    // Avvia l'animazione dello spinner.
    public void Start()
    {
        _active = true;
        if (!_thread.IsAlive)
        {
            _thread.Start();
        }
    }

    // Ferma l'animazione dello spinner.
    public void Stop()
    {
        _active = false;
        Console.Write("\b \b"); // Pulisce il carattere dello spinner.
    }

    private void Spin()
    {
        while (_active)
        {
            Thread.Sleep(100);
            Console.Write(Sequence[_counter++ % Sequence.Length]);
            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
        }
    }

    public void Dispose()
    {
        Stop();
        GC.SuppressFinalize(this);
    }
}

public class BomProcessor(SolidEdgeFramework.Application application, Spinner spinner)
{
    private readonly SolidEdgeFramework.Application _application = application;
    private readonly Spinner _spinner = spinner;

    public List<string> ProcessFiles(List<string> asmFiles)
    {
        var tsvRows = new List<string> { "Progressivo\tLivello\t\tNome file\t\tQuantità" };
        int progressivo = 1;

        if (asmFiles.Count > 0)
        {
            foreach (var asmFile in asmFiles)
            {
                ProcessAssemblyFile(asmFile, tsvRows, progressivo++);
            }
        }
        else
        {
            ProcessActiveDocument(tsvRows, progressivo);
        }

        return tsvRows;
    }

    private void ProcessAssemblyFile(string asmFile, List<string> tsvRows, int progressivo)
    {
        SolidEdgeAssembly.AssemblyDocument assemblyDocument = null;
        try
        {
            assemblyDocument = (SolidEdgeAssembly.AssemblyDocument)_application.Documents.Open(asmFile);
            var rootBomItem = CreateRootBomItem(assemblyDocument);
            PopulateBom("", assemblyDocument, rootBomItem);
            AddBomToTsv(tsvRows, rootBomItem, progressivo);
        }
        finally
        {
            assemblyDocument?.Close(false);
        }
    }

    private void ProcessActiveDocument(List<string> tsvRows, int progressivo)
    {
        var assemblyDocument = _application.GetActiveDocument<SolidEdgeAssembly.AssemblyDocument>(false);
        if (assemblyDocument != null)
        {
            var rootBomItem = CreateRootBomItem(assemblyDocument);
            PopulateBom("", assemblyDocument, rootBomItem);
            AddBomToTsv(tsvRows, rootBomItem, progressivo);
        }
        else
        {
            _spinner.Stop();
            MessageBox.Show("Nessun assieme di Solid Edge aperto o file di input fornito.", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private static BomItem CreateRootBomItem(SolidEdgeAssembly.AssemblyDocument assemblyDocument)
    {
        return new BomItem
        {
            FileName = Path.GetFileNameWithoutExtension(assemblyDocument.FullName)
        };
    }

    private static void AddBomToTsv(List<string> tsvRows, BomItem rootBomItem, int progressivo)
    {
        tsvRows.Add($"{progressivo}\t\t\t{rootBomItem.FileName}\t\t");
        foreach (var bomItem in rootBomItem.AllChildren)
        {
            tsvRows.Add($"\t{bomItem.LevelString}\t\t{bomItem.FileName}\t\t{bomItem.Quantity}");
        }
    }

    private static void PopulateBom(string levelString, SolidEdgeAssembly.AssemblyDocument assemblyDocument, BomItem parentBomItem)
    {
        int childIndex = 1;
        var (OccurrenceCount, UniqueOccurrences) = CollectOccurrenceData(assemblyDocument);

        foreach (var kvp in UniqueOccurrences)
        {
            var occurrence = kvp.Value;
            var lowerFileName = kvp.Key;
            string currentLevelString = string.IsNullOrEmpty(levelString) ? childIndex.ToString() : $"{levelString}.{childIndex}";

            var bomItem = new BomItem(occurrence, 0)
            {
                FileName = Path.GetFileNameWithoutExtension(occurrence.OccurrenceFileName),
                LevelString = currentLevelString,
                Quantity = OccurrenceCount[lowerFileName]
            };
            parentBomItem.Children.Add(bomItem);

            if (bomItem.IsSubassembly == true)
            {
                PopulateBom(currentLevelString, (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument, bomItem);
            }
            childIndex++;
        }
    }

    private static (Dictionary<string, int> OccurrenceCount, Dictionary<string, SolidEdgeAssembly.Occurrence> UniqueOccurrences) CollectOccurrenceData(SolidEdgeAssembly.AssemblyDocument assemblyDocument)
    {
        var occurrenceCount = new Dictionary<string, int>();
        var uniqueOccurrences = new Dictionary<string, SolidEdgeAssembly.Occurrence>();

        foreach (SolidEdgeAssembly.Occurrence occurrence in assemblyDocument.Occurrences)
        {
            if (!occurrence.IncludeInBom || occurrence.IsPatternItem || occurrence.OccurrenceDocument == null)
                continue;

            var lowerFileName = occurrence.OccurrenceFileName.ToLower();
            if (uniqueOccurrences.TryAdd(lowerFileName, occurrence))
            {
                occurrenceCount[lowerFileName] = 1;
            }
            else
            {
                occurrenceCount[lowerFileName]++;
            }
        }

        return (occurrenceCount, uniqueOccurrences);
    }
}

class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        SolidEdgeFramework.Application application = null;
        using var spinner = new Spinner();

        try
        {
            OleMessageFilter.Register();
            application = ConnectToSolidEdge(spinner);
            if (application == null) return;

            ConfigureSolidEdge(application);
            var asmFiles = GetInputFiles(args);
            if (asmFiles == null) return;

            Console.Write("Elaborazione in corso... ");
            spinner.Start();

            var processor = new BomProcessor(application, spinner);
            var tsvRows = processor.ProcessFiles(asmFiles);

            spinner.Stop();
            CreateAndOpenTsvFile(tsvRows);
        }
        catch (Exception ex)
        {
            spinner.Stop();
            MessageBox.Show($"Si è verificato un errore imprevisto:\n{ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            RestoreSolidEdgeSettings(application);
            OleMessageFilter.Unregister();
        }
    }

    private static SolidEdgeFramework.Application ConnectToSolidEdge(Spinner spinner)
    {
        try
        {
            Console.Write("Connessione a Solid Edge... ");
            spinner.Start();
            var application = SolidEdgeUtils.Connect(true);
            spinner.Stop();
            return application;
        }
        catch (Exception ex)
        {
            spinner.Stop();
            MessageBox.Show($"Impossibile avviare o connettersi a Solid Edge.\n{ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return null;
        }
    }

    private static void ConfigureSolidEdge(SolidEdgeFramework.Application application)
    {
        application.DelayCompute = true;
        application.DisplayAlerts = false;
        application.Interactive = false;
        application.ScreenUpdating = false;
    }

    private static List<string> GetInputFiles(string[] args)
    {
        var asmFiles = new List<string>();

        if (args.Length == 1 && args[0].ToLower().EndsWith(".txt"))
        {
            try
            {
                asmFiles.AddRange(File.ReadAllLines(args[0]));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Errore durante la lettura del file {args[0]}:\n{ex.Message}", "Errore File", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        else
        {
            asmFiles.AddRange(args.Where(f => f.ToLower().EndsWith(".asm")));
        }

        return asmFiles;
    }

    private static void CreateAndOpenTsvFile(List<string> tsvRows)
    {
        if (tsvRows.Count <= 1) return;

        Console.Write("\nCreazione del file BOM in corso... ");
        string tsvPath = Path.Combine(Path.GetTempPath(), $"BOM_{Guid.NewGuid()}.tsv");
        File.WriteAllLines(tsvPath, tsvRows, System.Text.Encoding.UTF8);
        Process.Start(new ProcessStartInfo(tsvPath) { UseShellExecute = true });
        Console.WriteLine("Fatto!");
    }

    private static void RestoreSolidEdgeSettings(SolidEdgeFramework.Application application)
    {
        if (application == null) return;

        application.DelayCompute = false;
        application.DisplayAlerts = true;
        application.Interactive = true;
        application.ScreenUpdating = true;
    }
}