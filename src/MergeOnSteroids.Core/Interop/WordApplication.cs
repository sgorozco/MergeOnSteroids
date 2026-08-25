using System.Runtime.InteropServices;

namespace MergeOnSteroids.Core.Interop;

/// <summary>
/// The one Word process the editor talks to, kept alive between operations: starting
/// Word costs seconds, so opening a fragment should not pay for it every time.
///
/// Not thread-safe — Word is an STA COM server, so create and use this from a single
/// STA thread.
/// </summary>
public sealed class WordApplication : IDisposable
{
    private const int WdDoNotSaveChanges = 0;

    private dynamic? _app;

    /// <summary>The running Word, started on first use.</summary>
    public dynamic Instance
    {
        get
        {
            var running = _app;
            // dynamic defeats the compiler's null analysis; the check above is the guarantee
            if (running is not null && IsAlive(running)) return running!;

            var app = Start();
            _app = app;
            return app;
        }
    }

    /// <summary>Brings Word to the front for the user to edit in.</summary>
    public void Show()
    {
        var app = Instance;
        app.Visible = true;
        try { app.Activate(); } catch (COMException) { /* focus is best effort */ }
    }

    /// <summary>Puts Word away again without closing it.</summary>
    public void Hide()
    {
        if (_app is null) return;
        try { _app.Visible = false; } catch (COMException) { /* already gone */ }
    }

    private static dynamic Start()
    {
        var wordType = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException(
                "Microsoft Word is not installed (Word.Application COM class not found).");

        dynamic app = Activator.CreateInstance(wordType)
            ?? throw new InvalidOperationException("Could not start Microsoft Word.");
        app.Visible = false;
        app.DisplayAlerts = 0;
        return app;
    }

    /// <summary>True while the reference still answers — the user may have closed Word themselves.</summary>
    private static bool IsAlive(dynamic app)
    {
        try
        {
            _ = app.Version;
            return true;
        }
        catch (COMException)
        {
            return false;
        }
    }

    public void Dispose()
    {
        if (_app is null) return;
        var app = _app;
        _app = null;

        try { app.Quit(WdDoNotSaveChanges); } catch (COMException) { /* already gone */ }

        // Deliberately NOT Marshal.ReleaseComObject: once Word has quit, releasing the
        // reference to the dead server blocks on an RPC timeout for ~27 seconds. Dropping
        // it costs nothing, and the process we were talking to is already gone.
    }
}
