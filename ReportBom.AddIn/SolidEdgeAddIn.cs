using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using SolidEdgeCommunity.AddIn;
using SolidEdgeFramework;
using System.Windows.Forms;

namespace ReportBom.AddIn;

[ComVisible(true)]
[Guid("004583E6-C0B8-4DD3-A8F7-E0DD7AC363E4")]
[ProgId(ProgIdValue)]
[ClassInterface(ClassInterfaceType.None)]
public sealed class SolidEdgeAddIn : SolidEdgeCommunity.AddIn.SolidEdgeAddIn
{
    private const string ProgIdValue = "ReportBom.AddIn";
    private const string AddInDisplayName = "ReportBom";
    private const string AddInDescription = "Avvia l'export BOM dal progetto ReportBom.";
    private const string SolidEdgeRootRegistryKey = @"Software\Siemens\Solid Edge";
    private const string SolidEdgeAddInCategoryId = "{26B1D2D1-2B03-11D2-B589-080036E8B802}";
    internal const string AssemblyEnvironmentCategoryId = "{26618395-09D6-11D1-BA07-080036230602}";
    internal const string ReportBomExecutableName = "ReportBom.exe";
    private const string NativeResourcesDllName = "ReportBomNativeResources.dll";

    public override string NativeResourcesDllPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, NativeResourcesDllName);

    public override void OnConnection(SolidEdgeFramework.Application application, SeConnectMode connectMode, SolidEdgeFramework.AddIn addInInstance)
    {
        base.OnConnection(application, connectMode, addInInstance);
        AddInEx.GuiVersion = 3;
    }

    public override void OnCreateRibbon(RibbonController controller, Guid environmentCategory, bool firstTime)
    {
        if (environmentCategory.Equals(new Guid(AssemblyEnvironmentCategoryId)))
        {
            controller.Add<ReportBomAssemblyRibbon>(environmentCategory, firstTime);
        }
    }

    [ComRegisterFunction]
    public static void OnRegister(Type type)
    {
        RegisterSolidEdgeComCategories(type);
        RegisterSolidEdgeAddIn(type);
    }

    [ComUnregisterFunction]
    public static void OnUnregister(Type type)
    {
        UnregisterSolidEdgeComCategories(type);
        UnregisterSolidEdgeAddIn(type);
    }

    private static void RegisterSolidEdgeComCategories(Type type)
    {
        var clsidPath = $@"CLSID\{type.GUID:B}";
        using var clsidKey = Registry.ClassesRoot.CreateSubKey(clsidPath);
        if (clsidKey == null)
        {
            return;
        }

        clsidKey.SetValue("409", AddInDisplayName);
        clsidKey.SetValue("AutoConnect", 1, RegistryValueKind.DWord);

        using var implementedCategoryKey = clsidKey.CreateSubKey($@"Implemented Categories\{SolidEdgeAddInCategoryId}");
        using var environmentCategoryKey = clsidKey.CreateSubKey($@"Environment Categories\{AssemblyEnvironmentCategoryId}");
        using var summaryKey = clsidKey.CreateSubKey("Summary");
        summaryKey?.SetValue("409", AddInDescription);
    }

    private static void UnregisterSolidEdgeComCategories(Type type)
    {
        var clsidPath = $@"CLSID\{type.GUID:B}";
        using var clsidKey = Registry.ClassesRoot.OpenSubKey(clsidPath, writable: true);
        clsidKey?.DeleteSubKeyTree("Implemented Categories", throwOnMissingSubKey: false);
        clsidKey?.DeleteSubKeyTree("Environment Categories", throwOnMissingSubKey: false);
        clsidKey?.DeleteSubKeyTree("Summary", throwOnMissingSubKey: false);
    }

    private static void RegisterSolidEdgeAddIn(Type type)
    {
        using var solidEdgeRoot = Registry.CurrentUser.OpenSubKey(SolidEdgeRootRegistryKey, writable: true);
        if (solidEdgeRoot == null)
        {
            return;
        }

        foreach (var versionKeyName in solidEdgeRoot.GetSubKeyNames())
        {
            if (!versionKeyName.StartsWith("Version ", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            using var addInsKey = solidEdgeRoot.CreateSubKey($@"{versionKeyName}\AddIns");
            using var addInKey = addInsKey?.CreateSubKey(type.GUID.ToString("B").ToUpperInvariant());
            addInKey?.SetValue(null, AddInDisplayName);
            addInKey?.SetValue("AutoConnect", 1, RegistryValueKind.DWord);
        }
    }

    private static void UnregisterSolidEdgeAddIn(Type type)
    {
        using var solidEdgeRoot = Registry.CurrentUser.OpenSubKey(SolidEdgeRootRegistryKey, writable: true);
        if (solidEdgeRoot == null)
        {
            return;
        }

        foreach (var versionKeyName in solidEdgeRoot.GetSubKeyNames())
        {
            if (!versionKeyName.StartsWith("Version ", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            using var addInsKey = solidEdgeRoot.OpenSubKey($@"{versionKeyName}\AddIns", writable: true);
            addInsKey?.DeleteSubKeyTree(type.GUID.ToString("B").ToUpperInvariant(), throwOnMissingSubKey: false);
        }
    }

    internal static void LaunchReportBomExecutable()
    {
        var executablePath = ResolveExecutablePath();
        if (string.IsNullOrWhiteSpace(executablePath))
        {
            MessageBox.Show(
                "Non trovo ReportBom.exe. Compila prima il progetto ReportBom oppure copia i file di output accanto all'add-in.",
                "ReportBom",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = executablePath,
                WorkingDirectory = Path.GetDirectoryName(executablePath),
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Impossibile avviare {ReportBomExecutableName}:{System.Environment.NewLine}{ex.Message}",
                "ReportBom",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private static string ResolveExecutablePath()
    {
        var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var candidatePaths = new[]
        {
            Path.Combine(baseDirectory, "ReportBom", ReportBomExecutableName),
            Path.Combine(baseDirectory, ReportBomExecutableName),
            Path.GetFullPath(Path.Combine(baseDirectory, @"..\..\..\ReportBom\bin\Debug\net8.0-windows\ReportBom.exe")),
            Path.GetFullPath(Path.Combine(baseDirectory, @"..\..\..\ReportBom\bin\Release\net8.0-windows\ReportBom.exe"))
        };

        foreach (var candidatePath in candidatePaths)
        {
            if (File.Exists(candidatePath))
            {
                return candidatePath;
            }
        }

        return null;
    }
}
