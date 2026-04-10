namespace ReportBom;

public class BomItem
{
    private List<BomItem> _children = [];

    public string FileName { get; set; }
    public string LevelString { get; set; } // String representation of BOM level (e.g., "1", "1.1")
    public int Quantity { get; set; }
    public bool IsSubassembly { get; set; }

    /// Returns all direct children.
    public List<BomItem> Children
    {
        get => _children;
        set => _children = value ?? [];
    }

    /// Returns all direct and descendant children.
    public IEnumerable<BomItem> AllChildren
    {
        get
        {
            foreach (var bomItem in Children)
            {
                yield return bomItem;

                if (bomItem.IsSubassembly)
                {
                    foreach (var childBomItem in bomItem.AllChildren)
                    {
                        yield return childBomItem;
                    }
                }
            }
        }
    }

    public static BomItem CreateRoot(string assemblyFullName)
    {
        return new BomItem
        {
            FileName = Path.GetFileNameWithoutExtension(assemblyFullName)
        };
    }

    public static BomItem FromOccurrence(SolidEdgeAssembly.Occurrence occurrence, string levelString, int quantity)
    {
        return new BomItem
        {
            FileName = Path.GetFileNameWithoutExtension(occurrence.OccurrenceFileName),
            LevelString = levelString,
            Quantity = quantity,
            IsSubassembly = occurrence.Subassembly
        };
    }

    public bool ShouldSerializeChildren()
    {
        return Children.Count > 0;
    }

    public override string ToString() => FileName ?? string.Empty;
}
