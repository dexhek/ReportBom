using SolidEdgeCommunity.Extensions;

namespace ReportBom;

public sealed class BomProcessor(ProcessingStatus status)
{
    private readonly ProcessingStatus _status = status;

    public List<string> ProcessFiles(SolidEdgeFramework.Application application, IReadOnlyList<string> asmFiles)
    {
        var tsvRows = new List<string> { "Progressivo\tLivello\t\tNome file\t\tQuantità" };
        int progressivo = 1;

        if (asmFiles.Count > 0)
        {
            for (int i = 0; i < asmFiles.Count; i++)
            {
                string asmFile = asmFiles[i];
                _status.ReportAssemblyProgress(i + 1, asmFiles.Count, asmFile);
                ProcessAssemblyFile(application, asmFile, tsvRows, progressivo++);
            }
        }
        else
        {
            _status.ReportStep("Elaborazione assieme attivo...");
            ProcessActiveDocument(application, tsvRows, progressivo);
        }

        return tsvRows;
    }

    private void ProcessAssemblyFile(SolidEdgeFramework.Application application, string asmFile, List<string> tsvRows, int progressivo)
    {
        SolidEdgeAssembly.AssemblyDocument assemblyDocument = null;

        try
        {
            assemblyDocument = (SolidEdgeAssembly.AssemblyDocument)application.Documents.Open(asmFile);
            var rootBomItem = BomItem.CreateRoot(assemblyDocument.FullName);
            PopulateBom(string.Empty, assemblyDocument, rootBomItem);
            AddBomToTsv(tsvRows, rootBomItem, progressivo);
        }
        finally
        {
            assemblyDocument?.Close(false);
        }
    }

    private void ProcessActiveDocument(SolidEdgeFramework.Application application, List<string> tsvRows, int progressivo)
    {
        var assemblyDocument = application.GetActiveDocument<SolidEdgeAssembly.AssemblyDocument>(false);
        if (assemblyDocument != null)
        {
            var rootBomItem = BomItem.CreateRoot(assemblyDocument.FullName);
            PopulateBom(string.Empty, assemblyDocument, rootBomItem);
            AddBomToTsv(tsvRows, rootBomItem, progressivo);
            return;
        }

        _status.ShowWarning("Nessun assieme di Solid Edge aperto o file di input fornito.", "Avviso");
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
        var (occurrenceCount, uniqueOccurrences) = CollectOccurrenceData(assemblyDocument);

        if (uniqueOccurrences.Count == 1)
        {
            var kvp = uniqueOccurrences.First();
            var singleOccurrence = kvp.Value;
            int quantity = occurrenceCount[kvp.Key];

            if (singleOccurrence.Subassembly && quantity == 1)
            {
                PopulateBom(levelString, (SolidEdgeAssembly.AssemblyDocument)singleOccurrence.OccurrenceDocument, parentBomItem);
                return;
            }

            if (!singleOccurrence.Subassembly && quantity == 1)
            {
                // Assemblies with a single part remain collapsed on one line.
                return;
            }
        }

        int childIndex = 1;
        foreach (var kvp in uniqueOccurrences)
        {
            var occurrence = kvp.Value;
            int quantity = occurrenceCount[kvp.Key];
            string currentLevelString = string.IsNullOrEmpty(levelString) ? childIndex.ToString() : $"{levelString}.{childIndex}";

            var bomItem = BomItem.FromOccurrence(occurrence, currentLevelString, quantity);
            parentBomItem.Children.Add(bomItem);

            if (bomItem.IsSubassembly)
            {
                PopulateBom(currentLevelString, (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument, bomItem);
            }

            childIndex++;
        }
    }

    private static (Dictionary<string, int> OccurrenceCount, Dictionary<string, SolidEdgeAssembly.Occurrence> UniqueOccurrences) CollectOccurrenceData(SolidEdgeAssembly.AssemblyDocument assemblyDocument)
    {
        var occurrenceCount = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var uniqueOccurrences = new Dictionary<string, SolidEdgeAssembly.Occurrence>(StringComparer.OrdinalIgnoreCase);

        foreach (SolidEdgeAssembly.Occurrence occurrence in assemblyDocument.Occurrences)
        {
            if (!occurrence.IncludeInBom || occurrence.IsPatternItem || occurrence.OccurrenceDocument == null)
            {
                continue;
            }

            string occurrenceFileName = occurrence.OccurrenceFileName;
            if (uniqueOccurrences.TryAdd(occurrenceFileName, occurrence))
            {
                occurrenceCount[occurrenceFileName] = 1;
            }
            else
            {
                occurrenceCount[occurrenceFileName]++;
            }
        }

        return (occurrenceCount, uniqueOccurrences);
    }
}
