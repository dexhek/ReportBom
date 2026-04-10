using SolidEdgeCommunity;

namespace ReportBom;

public sealed class SolidEdgeSession : IDisposable
{
    private bool _disposed;

    private SolidEdgeSession(SolidEdgeFramework.Application application)
    {
        Application = application;
        Configure(Application);
    }

    public SolidEdgeFramework.Application Application { get; }

    public static SolidEdgeSession Connect(ProcessingStatus status)
    {
        try
        {
            status.ReportStep("Connessione a Solid Edge...");
            OleMessageFilter.Register();
            var application = SolidEdgeUtils.Connect(true);
            return new SolidEdgeSession(application);
        }
        catch (Exception ex)
        {
            OleMessageFilter.Unregister();
            status.ShowError($"Impossibile avviare o connettersi a Solid Edge.\n{ex.Message}", "Errore");
            return null;
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        Restore(Application);
        OleMessageFilter.Unregister();
        _disposed = true;
        GC.SuppressFinalize(this);
    }

    private static void Configure(SolidEdgeFramework.Application application)
    {
        application.DelayCompute = true;
        application.DisplayAlerts = false;
        application.Interactive = false;
        application.ScreenUpdating = false;
    }

    private static void Restore(SolidEdgeFramework.Application application)
    {
        if (application == null)
        {
            return;
        }

        application.DelayCompute = false;
        application.DisplayAlerts = true;
        application.Interactive = true;
        application.ScreenUpdating = true;
    }
}
