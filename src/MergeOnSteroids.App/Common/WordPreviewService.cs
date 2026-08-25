using System.Collections.Concurrent;
using System.Windows.Threading;
using MergeOnSteroids.Core.Blocks;
using MergeOnSteroids.Core.Interop;

namespace MergeOnSteroids.App.Common;

/// <summary>
/// Keeps one hidden Word instance laying out paragraph previews on its own STA
/// thread, so the editor never blocks on COM while you type. Requests are keyed by
/// everything that affects the rendering, so repeating a paragraph — or coming back
/// to one — costs nothing.
/// </summary>
public sealed class WordPreviewService : IDisposable
{
    private const int MaxCachedRenders = 400;

    private readonly Dispatcher _dispatcher;
    private readonly Action<string> _log;
    private readonly BlockingCollection<Job> _queue = new();
    private readonly Dictionary<string, byte[]> _cache = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<ParagraphBlock, string> _latest = new();

    private Thread? _worker;
    private bool _wordUnavailable;

    private sealed record Job(ParagraphBlock Block, ParagraphPreviewRequest Request, Action<ParagraphBlock, byte[]> Done);

    public WordPreviewService(Dispatcher dispatcher, Action<string> log)
    {
        _dispatcher = dispatcher;
        _log = log;
    }

    /// <summary>Queues a paragraph for rendering; the callback runs on the UI thread.</summary>
    public void Request(ParagraphBlock block, ParagraphPreviewRequest request, Action<ParagraphBlock, byte[]> done)
    {
        if (_wordUnavailable || _queue.IsAddingCompleted) return;

        // Remember what this block is waiting for, so anything queued before the
        // last keystroke can be dropped rather than rendered and thrown away.
        _latest[block] = request.Key;

        EnsureWorker();
        try
        {
            _queue.Add(new Job(block, request, done));
        }
        catch (InvalidOperationException)
        {
            // queue closed while we were starting up
        }
    }

    private void EnsureWorker()
    {
        if (_worker is not null) return;

        _worker = new Thread(Work)
        {
            IsBackground = true,
            Name = "Word paragraph previews",
        };
        _worker.SetApartmentState(ApartmentState.STA);   // Word is an STA COM server
        _worker.Start();
    }

    private void Work()
    {
        WordParagraphPreview? renderer = null;
        try
        {
            foreach (var job in _queue.GetConsumingEnumerable())
            {
                if (_latest.TryGetValue(job.Block, out var wanted) && wanted != job.Request.Key)
                    continue;   // superseded while it waited

                if (_cache.TryGetValue(job.Request.Key, out var cached))
                {
                    Post(job, cached);
                    continue;
                }

                byte[] png;
                try
                {
                    renderer ??= new WordParagraphPreview();
                    png = renderer.Render(job.Request);
                }
                catch (Exception ex)
                {
                    _wordUnavailable = true;
                    _dispatcher.BeginInvoke(() => _log(
                        $"warning: Word previews turned off — {ex.Message}"));
                    return;
                }

                if (_cache.Count >= MaxCachedRenders) _cache.Clear();
                _cache[job.Request.Key] = png;
                Post(job, png);
            }
        }
        finally
        {
            renderer?.Dispose();
        }
    }

    private void Post(Job job, byte[] png) =>
        _dispatcher.BeginInvoke(() =>
        {
            // only paint it if the block still wants this exact rendering
            if (!_latest.TryGetValue(job.Block, out var wanted) || wanted == job.Request.Key)
                job.Done(job.Block, png);
        });

    /// <summary>Forgets which blocks were waiting — the old program's are gone.</summary>
    public void Reset() => _latest.Clear();

    public void Dispose()
    {
        _queue.CompleteAdding();
        _worker?.Join(TimeSpan.FromSeconds(10));   // let Word close its documents
        _worker = null;
        _queue.Dispose();
    }
}
