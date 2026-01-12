namespace ReportBom;

public class BomItem
{
    private List<BomItem> _children = [];

        public BomItem()
        {
            // Parameterless constructor for serialization.
        }

    public BomItem(SolidEdgeAssembly.Occurrence occurrence, int level)
        {
            Level = level;
            FileName = Path.GetFullPath(occurrence.OccurrenceFileName);
            IsMissing = occurrence.FileMissing();
            Quantity = 1;
            IsSubassembly = occurrence.Subassembly;

            // If the file exists, extract file properties.
            if (IsMissing == false)
            {
                var document = (SolidEdgeFramework.SolidEdgeDocument)occurrence.OccurrenceDocument;
                var summaryInfo = (SolidEdgeFramework.SummaryInfo)document.SummaryInfo;

                DocumentNumber = summaryInfo.DocumentNumber;
                Title = summaryInfo.Title;
                Revision = summaryInfo.RevisionNumber;
            }
        }

        public int? Level { get; set; }
        public string DocumentNumber { get; set; }
        public string Revision { get; set; }
        public string Title { get; set; }
        public int? Quantity { get; set; }
        public string FileName { get; set; }
        public string LevelString { get; set; } // String representation of BOM level (e.g., "1", "1.1")

        public bool? IsSubassembly { get; set; }

        public bool? IsMissing { get; set; }

        /// Returns all direct children.
        public List<BomItem> Children { get => _children; set => _children = value; }

        /// Returns all direct and descendant children.
        public IEnumerable<BomItem> AllChildren
        {
            get
            {
                foreach (var bomItem in Children)
                {
                    yield return bomItem;

                    if (bomItem.IsSubassembly == true)
                    {
                        foreach (var childBomItem in bomItem.AllChildren)
                        {
                            yield return childBomItem;
                        }
                    }
                }
            }
        }

        public bool ShouldSerializeChildren()
        {
            return Children.Count > 0;
        }

        public override string ToString() => FileName ?? string.Empty;
    }
