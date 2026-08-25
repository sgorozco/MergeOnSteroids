using System.Collections.Concurrent;
using System.Windows.Threading;
using MergeOnSteroids.Core.Blocks;
using MergeOnSteroids.Core.Interop;

namespace MergeOnSteroids.App.Common;

/// <summary>
/// Owns the editor's single Word process and the STA thread that talks to it, so no
/// COM call ever happens on the UI thread. Two kinds of work go through it: paragraph
/// previews, keyed and cached so repeating a paragraph costs nothing, and whatever a
/// fragment edit needs, via <see cref="InvokeAsync"/>.
///
/// Word stays running between operations. Starting it costs seconds and stopping it
/// can cost far more, so the editor pays that once rather than per edit.
/// </summary>
public sealed class WordService : IDisposable
{
    private const int MaxCachedRenders = 400;

    private readonly Dispatcher _dispatcher;
    private readonly Action<string> _log;
    private readonly BlockingCollection<Job> _queue = new();
    private readonly Dictionary<string, byte[]> _cache = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<ParagraphBlock, string> _latest = new();

    private readonly WordApplication _word = new();
    private Thread? _worker;
    private bool _wordUnavailable;

    private abstract record Job;

    private sealed record PreviewJob(
        ParagraphBlock Block, ParagraphPreviewRequest Request, Action<ParagraphBlock, byte[]> Done) : Job;

    /// <summary>Arbitrary Word work, run on the Word thread and reported back through the task.</summary>
    private sealed record InvokeJob(Func<WordApplication, object?> Work, TaskCompletionSource<object?> Completion) : Job;

    public WordService(Dispatcher dispatcher, Action<string> log)
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
            _queue.Add(new PreviewJob(block, request, done));
        }
        catch (InvalidOperationException)
        {
            // queue closed while we were starting up
        }
    }

    /// <summary>
    /// Runs <paramref name="work"/> against Word on the Word thread. Everything the
    /// fragment editing flow does goes through here, so the UI thread stays free even
    /// while Word is starting.
    /// </summary>
    public async Task<T> InvokeAsync<T>(Func<WordApplication, T> work)
    {
        var completion = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        EnsureWorker();
        _queue.Add(new InvokeJob(word => work(word), completion));
        return (T)(await completion.Task)!;
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
                switch (job)
                {
                    case InvokeJob invoke:
                        try
                        {
                            invoke.Completion.SetResult(invoke.Work(_word));
                        }
                        catch (Exception ex)
                        {
                            invoke.Completion.SetException(ex);
                        }
                        break;

                    case PreviewJob preview:
                        RenderPreview(preview, ref renderer);
                        break;
                }
            }
        }
        finally
        {
            renderer?.Dispose();
            _word.Dispose();      // quits Word; the reference is dropped, never released
        }
    }

    private void RenderPreview(PreviewJob job, ref WordParagraphPreview? renderer)
    {
        if (_latest.TryGetValue(job.Block, out var wanted) && wanted != job.Request.Key)
            return;   // superseded while it waited

        if (_cache.TryGetValue(job.Request.Key, out var cached))
        {
            Post(job, cached);
            return;
        }

        byte[] png;
        try
        {
            renderer ??= new WordParagraphPreview(_word);
            png = renderer.Render(job.Request);
        }
        catch (Exception ex)
        {
            _wordUnavailable = true;
            _dispatcher.BeginInvoke(() => _log($"warning: Word previews turned off — {ex.Message}"));
            return;
        }

        if (_cache.Count >= MaxCachedRenders) _cache.Clear();
        _cache[job.Request.Key] = png;
        Post(job, png);
    }

    private void Post(PreviewJob job, byte[] png) =>
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
