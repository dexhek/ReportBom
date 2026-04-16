using System.Reflection;
using SolidEdgeCommunity.AddIn;

namespace ReportBom.AddIn;

internal sealed class ReportBomAssemblyRibbon : Ribbon
{
    public ReportBomAssemblyRibbon()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resourceName = $"{GetType().FullName}.xml";
        LoadXml(assembly, resourceName);

        var button = GetButton(0);
        button.Label = "ReportBom";
        button.ScreenTip = "Esporta la BOM dell'assieme attivo.";
        button.SuperTip = button.ScreenTip;
        button.ShowImage = true;
        button.ShowLabel = true;
        button.Click += OnReportBomClick;
    }

    private static void OnReportBomClick(RibbonControl control)
    {
        SolidEdgeAddIn.LaunchReportBomExecutable();
    }
}
