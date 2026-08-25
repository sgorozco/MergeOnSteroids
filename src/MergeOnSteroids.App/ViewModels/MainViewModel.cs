using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Threading;
using MergeOnSteroids.App.Common;
using MergeOnSteroids.Core;
using MergeOnSteroids.Core.Blocks;
using MergeOnSteroids.Core.Data;
using MergeOnSteroids.Core.Fragments;
using MergeOnSteroids.Core.Interop;
using MergeOnSteroids.Core.Runtime;
using MergeOnSteroids.Core.Samples;
using Microsoft.Win32;

namespace MergeOnSteroids.App.ViewModels;

public sealed class MainViewModel : INotifyPropertyChanged
{
    /// <summary>Product name, used as the caption of every dialog.</summary>
    public const string AppName = "Merge on Steroids";

    public static MainViewModel? Current { get; private set; }

    private readonly Dispatcher _dispatcher = Application.Current.Dispatcher;
    private readonly StringBuilder _log = new();
    private CancellationTokenSource? _runCts;

    // File sources re-read their columns shortly after the path/sheet/header row
    // changes; database sources only when asked, since that means talking to a server.
    private readonly DispatcherTimer _schemaTimer;
    private readonly List<FileSourceBlockBase> _schemaQueue = [];
    private readonly HashSet<SourceBlockBase> _schemaReading = [];

    // Paragraph previews drawn by Word, when switched on.
    private readonly DispatcherTimer _previewTimer;
    private readonly List<ParagraphBlock> _previewQueue = [];
    private WordPreviewService? _wordPreviews;
    private bool _wordPreviewsEnabled;

    private ProgramModel _program = new();
    private string? _currentFilePath;
    private bool _isDirty;
    private bool _isRunning;
    private bool _showWordWindow;
    private Block? _selectedBlock;

    public MainViewModel()
    {
        Current = this;
        Palette = BuildPalette();
        PaletteView = System.Windows.Data.CollectionViewSource.GetDefaultView(Palette);
        PaletteView.GroupDescriptions.Add(
            new System.Windows.Data.PropertyGroupDescription(nameof(PaletteItem.Category)));

        NewCommand = new RelayCommand(NewProgram);
        OpenCommand = new RelayCommand(OpenProgram);
        SaveCommand = new RelayCommand(() => SaveProgram(saveAs: false));
        SaveAsCommand = new RelayCommand(() => SaveProgram(saveAs: true));
        LoadSampleCommand = new RelayCommand(LoadSample);
        RunPreviewCommand = new RelayCommand(() => _ = RunAsync(useWord: false), () => !IsRunning);
        RunWordCommand = new RelayCommand(() => _ = RunAsync(useWord: true), () => !IsRunning);
        StopCommand = new RelayCommand(() => _runCts?.Cancel(), () => IsRunning);
        OpenOutputFolderCommand = new RelayCommand(OpenOutputFolder);
        DeleteSelectedCommand = new RelayCommand(DeleteSelected, () => SelectedBlock is not null);
        ClearLogCommand = new RelayCommand(() => { _log.Clear(); OnPropertyChanged(nameof(LogText)); });
        EditFragmentCommand = new ParamRelayCommand(
            p => EditFragment(p as WordFragmentBlock),
            p => !IsRunning && p is WordFragmentBlock);
        RefreshFragmentCommand = new ParamRelayCommand(
            p => RefreshFragment(p as WordFragmentBlock),
            p => !IsRunning && p is WordFragmentBlock { FragmentFile.Length: > 0 });
        ToggleThemeCommand = new RelayCommand(ThemeManager.Toggle);
        ReadSchemaCommand = new ParamRelayCommand(
            p => _ = ReadSchemaAsync(p as SourceBlockBase),
            p => p is FileSourceBlockBase { FilePath.Length: > 0 }
                or DatabaseSourceBlock { ConnectionString.Length: > 0, Query.Length: > 0 });
        InsertColumnCommand = new ParamRelayCommand(p => InsertColumnReference(p as SourceColumnInfo));

        ThemeManager.ThemeChanged += OnThemeChanged;
        _schemaTimer = new DispatcherTimer(DispatcherPriority.Background, _dispatcher)
        {
            Interval = TimeSpan.FromMilliseconds(700)
        };
        _schemaTimer.Tick += (_, _) => { _schemaTimer.Stop(); ReadQueuedSchemas(); };

        _previewTimer = new DispatcherTimer(DispatcherPriority.Background, _dispatcher)
        {
            Interval = TimeSpan.FromMilliseconds(500)
        };
        _previewTimer.Tick += (_, _) => { _previewTimer.Stop(); RenderQueuedPreviews(); };
        _wordPreviewsEnabled = UserSettings.Load().WordPreviews ?? true;

        LoadProgram(new ProgramModel(), null);
    }

    // ------------------------------------------------------------ properties

    public ProgramModel Program
    {
        get => _program;
        private set { _program = value; OnPropertyChanged(); }
    }

    public string? CurrentFilePath
    {
        get => _currentFilePath;
        private set { _currentFilePath = value; OnPropertyChanged(); OnPropertyChanged(nameof(WindowTitle)); }
    }

    public bool IsDirty
    {
        get => _isDirty;
        private set { _isDirty = value; OnPropertyChanged(); OnPropertyChanged(nameof(WindowTitle)); }
    }

    public bool IsRunning
    {
        get => _isRunning;
        private set { _isRunning = value; OnPropertyChanged(); }
    }

    public bool ShowWordWindow
    {
        get => _showWordWindow;
        set { _showWordWindow = value; OnPropertyChanged(); }
    }

    public Block? SelectedBlock
    {
        get => _selectedBlock;
        set { _selectedBlock = value; OnPropertyChanged(); SelectionChanged?.Invoke(); }
    }

    public string WindowTitle =>
        $"{AppName} — {(CurrentFilePath is null ? "unsaved program" : Path.GetFileName(CurrentFilePath))}{(IsDirty ? " *" : "")}";

    public string LogText => _log.ToString();

    public ObservableCollection<string> AvailableSources { get; } = [];
    public ObservableCollection<string> OutputFiles { get; } = [];
    public List<PaletteItem> Palette { get; }
    public ICollectionView PaletteView { get; }
    public string[] ParagraphStyleNames => ParagraphStyles.All;

    /// <summary>Raised when SelectedBlock changes (BlockChrome uses it to show selection).</summary>
    public event Action? SelectionChanged;

    // -------------------------------------------------------------- commands

    public RelayCommand NewCommand { get; }
    public RelayCommand OpenCommand { get; }
    public RelayCommand SaveCommand { get; }
    public RelayCommand SaveAsCommand { get; }
    public RelayCommand LoadSampleCommand { get; }
    public RelayCommand RunPreviewCommand { get; }
    public RelayCommand RunWordCommand { get; }
    public RelayCommand StopCommand { get; }
    public RelayCommand OpenOutputFolderCommand { get; }
    public RelayCommand DeleteSelectedCommand { get; }
    public RelayCommand ClearLogCommand { get; }
    public ParamRelayCommand EditFragmentCommand { get; }
    public ParamRelayCommand RefreshFragmentCommand { get; }
    public ParamRelayCommand ReadSchemaCommand { get; }
    public ParamRelayCommand InsertColumnCommand { get; }

    // -------------------------------------------------------- file handling

    private void NewProgram()
    {
        if (!ConfirmDiscard()) return;
        LoadProgram(new ProgramModel(), null);
    }

    private void OpenProgram()
    {
        if (!ConfirmDiscard()) return;
        var dialog = new OpenFileDialog
        {
            Filter = $"{AppName} program (*.mos.json)|*.mos.json|JSON files (*.json)|*.json|All files (*.*)|*.*"
        };
        if (dialog.ShowDialog() != true) return;
        try
        {
            LoadProgram(ProgramSerializer.Load(dialog.FileName), dialog.FileName);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Could not open program:\n{ex.Message}", AppName,
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SaveProgram(bool saveAs)
    {
        var path = CurrentFilePath;
        if (saveAs || path is null)
        {
            var dialog = new SaveFileDialog
            {
                Filter = $"{AppName} program (*.mos.json)|*.mos.json",
                FileName = SanitizeName(Program.Name) + ProgramSerializer.FileExtension
            };
            if (dialog.ShowDialog() != true) return;
            path = dialog.FileName;
        }
        try
        {
            ProgramSerializer.Save(Program, path);
            CurrentFilePath = path;
            IsDirty = false;
            Log($"Saved: {path}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Could not save program:\n{ex.Message}", AppName,
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void LoadSample()
    {
        if (!ConfirmDiscard()) return;
        try
        {
            var folder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "MergeOnSteroids", "Sample");
            var path = SampleFactory.CreateSample(folder);
            LoadProgram(ProgramSerializer.Load(path), path);
            Log($"Sample created in {folder}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Could not create sample:\n{ex.Message}", AppName,
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>Loads a program file (used for command-line "open with").</summary>
    public void TryLoadFile(string path)
    {
        try
        {
            LoadProgram(ProgramSerializer.Load(path), Path.GetFullPath(path));
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Could not open program:\n{ex.Message}", AppName,
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public bool ConfirmDiscard()
    {
        if (!IsDirty) return true;
        var answer = MessageBox.Show(
            "The current program has unsaved changes. Discard them?", AppName,
            MessageBoxButton.YesNo, MessageBoxImage.Warning);
        return answer == MessageBoxResult.Yes;
    }

    private void LoadProgram(ProgramModel program, string? path)
    {
        UnhookTree(Program);
        Program = program;
        Program.PropertyChanged += OnProgramPropertyChanged;
        HookTree(program);
        CurrentFilePath = path;
        SelectedBlock = null;
        RefreshSources();
        HydrateFragments();
        foreach (var source in Program.AllBlocks().OfType<FileSourceBlockBase>()) _ = ReadSchemaAsync(source);
        _wordPreviews?.Reset();
        QueueAllPreviews();
        IsDirty = false;   // hydration touches blocks; that is not a user edit
    }

    // --------------------------------------------------------- tree tracking

    private void OnProgramPropertyChanged(object? sender, PropertyChangedEventArgs e) => IsDirty = true;

    private void HookTree(ProgramModel program) => HookList(program.Blocks);

    private void UnhookTree(ProgramModel program)
    {
        program.PropertyChanged -= OnProgramPropertyChanged;
        UnhookList(program.Blocks);
    }

    private void HookList(BlockCollection list)
    {
        list.CollectionChanged += OnTreeCollectionChanged;
        foreach (var b in list) HookBlock(b);
    }

    private void UnhookList(BlockCollection list)
    {
        list.CollectionChanged -= OnTreeCollectionChanged;
        foreach (var b in list) UnhookBlock(b);
    }

    private void HookBlock(Block b)
    {
        b.PropertyChanged += OnBlockPropertyChanged;
        foreach (var l in b.ChildLists()) HookList(l);
    }

    private void UnhookBlock(Block b)
    {
        b.PropertyChanged -= OnBlockPropertyChanged;
        foreach (var l in b.ChildLists()) UnhookList(l);
    }

    private void OnTreeCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.OldItems is not null)
            foreach (Block b in e.OldItems) UnhookBlock(b);
        if (e.NewItems is not null)
            foreach (Block b in e.NewItems) HookBlock(b);
        IsDirty = true;
        RefreshSources();
    }

    private void OnBlockPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (IsDesignTimeProperty(e.PropertyName)) return;   // read back from a file, not a user edit

        IsDirty = true;
        if (e.PropertyName == nameof(SourceBlockBase.Name))
            RefreshSources();

        if (sender is FileSourceBlockBase source && e.PropertyName is
            nameof(FileSourceBlockBase.FilePath) or nameof(FileSourceBlockBase.HeaderRow) or
            nameof(ExcelSourceBlock.SheetName))
            QueueSchemaRead(source);

        if (sender is ParagraphBlock paragraph && e.PropertyName is
            nameof(ParagraphBlock.TextTemplate) or nameof(ParagraphBlock.Style) or
            nameof(ParagraphBlock.Bold) or nameof(ParagraphBlock.Italic))
            QueuePreview(paragraph);

        // a document's template decides how every paragraph inside it looks
        if (sender is NewDocumentBlock && e.PropertyName == nameof(NewDocumentBlock.TemplatePath))
            QueueAllPreviews();
    }

    /// <summary>
    /// Block properties the editor fills in from the files a block points at
    /// (fragment previews, Excel columns). Changing them is not a program edit.
    /// </summary>
    private static bool IsDesignTimeProperty(string? name) => name is
        nameof(WordFragmentBlock.PlainText) or
        nameof(WordFragmentBlock.PreviewImage) or
        nameof(SourceBlockBase.DetectedColumns) or
        nameof(SourceBlockBase.HasDetectedColumns) or
        nameof(SourceBlockBase.SchemaStatus) or
        nameof(ExcelSourceBlock.DetectedSheets);

    private void RefreshSources()
    {
        var names = Program.AllBlocks()
            .OfType<SourceBlockBase>()
            .Select(s => s.Name)
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        AvailableSources.Clear();
        foreach (var n in names) AvailableSources.Add(n);
    }

    // ---------------------------------------------------------------- editing

    /// <summary>
    /// The heart of the fragment experiment: open the fragment's own .docx in a
    /// real Word window, let the user edit it there, then save it back and refresh
    /// the block's Word-rendered preview.
    /// </summary>
    private void EditFragment(WordFragmentBlock? block)
    {
        if (block is null) return;
        if (CurrentFilePath is null)
        {
            MessageBox.Show(
                "Save the program first.\n\nFragments are stored as small .docx files next to the " +
                "program file, so the program needs a location before one can be created.",
                AppName, MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var baseFolder = Path.GetDirectoryName(CurrentFilePath)!;
        var relative = string.IsNullOrWhiteSpace(block.FragmentFile)
            ? FragmentFiles.DefaultRelativePath(block.Id)
            : block.FragmentFile;
        var fragmentPath = Path.GetFullPath(Path.Combine(baseFolder, relative));

        Core.Interop.WordFragmentEditSession session;
        try
        {
            System.Windows.Input.Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            session = File.Exists(fragmentPath)
                ? Core.Interop.WordFragmentEditSession.OpenFile(fragmentPath)
                : Core.Interop.WordFragmentEditSession.Start(
                    string.IsNullOrWhiteSpace(block.LegacyFragmentXml) ? null : block.LegacyFragmentXml);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Could not start Microsoft Word:\n{ex.Message}", AppName,
                MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }
        finally
        {
            System.Windows.Input.Mouse.OverrideCursor = null;
        }

        try
        {
            var dialog = new Views.WordEditDialog { Owner = Application.Current.MainWindow };
            if (dialog.ShowDialog() == true)
                ApplyCapture(block, fragmentPath, session.SaveAs(fragmentPath), baseFolder);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Could not save the fragment back from Word (was the document closed?):\n{ex.Message}",
                AppName, MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        finally
        {
            try { session.Dispose(); } catch { /* Word may already be gone */ }
        }
    }

    /// <summary>
    /// Re-reads a fragment file — use it after editing the .docx directly in Word
    /// (from Explorer, say), or after pointing the block at a different file.
    /// </summary>
    private void RefreshFragment(WordFragmentBlock? block)
    {
        if (block is null || string.IsNullOrWhiteSpace(block.FragmentFile)) return;

        var baseFolder = CurrentFilePath is not null ? Path.GetDirectoryName(CurrentFilePath)! : RunBaseFolder;
        var fragmentPath = Path.GetFullPath(Path.Combine(baseFolder, block.FragmentFile));
        if (!File.Exists(fragmentPath))
        {
            MessageBox.Show($"Fragment file not found:\n{fragmentPath}", AppName,
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            System.Windows.Input.Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            using var session = Core.Interop.WordFragmentEditSession.OpenFile(fragmentPath, visible: false);
            ApplyCapture(block, fragmentPath, session.Capture(), baseFolder);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Could not read the fragment:\n{ex.Message}", AppName,
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        finally
        {
            System.Windows.Input.Mouse.OverrideCursor = null;
        }
    }

    private void ApplyCapture(
        WordFragmentBlock block, string fragmentPath, Core.Interop.FragmentCapture capture, string baseFolder)
    {
        FragmentFiles.WritePreview(fragmentPath, capture.PreviewPng);
        block.FragmentFile = FragmentFiles.MakeRelative(baseFolder, fragmentPath);
        block.LegacyFragmentXml = null;   // migrated to a file
        block.PlainText = capture.PlainText;
        block.PreviewImage = capture.PreviewPng.Length > 0 ? capture.PreviewPng : null;
        Log($"Fragment '{block.FragmentFile}' updated — " +
            $"{capture.PlainText.Split('\n').Length} paragraph(s).");
    }

    /// <summary>Loads each fragment's text and preview image from its sidecar files.</summary>
    private void HydrateFragments()
    {
        if (CurrentFilePath is null) return;
        var baseFolder = Path.GetDirectoryName(CurrentFilePath)!;

        foreach (var block in Program.AllBlocks().OfType<WordFragmentBlock>())
        {
            if (string.IsNullOrWhiteSpace(block.FragmentFile)) continue;
            var path = Path.GetFullPath(Path.Combine(baseFolder, block.FragmentFile));
            if (!File.Exists(path))
            {
                Log($"warning: fragment file missing — {block.FragmentFile}");
                continue;
            }
            block.PreviewImage = FragmentFiles.TryReadPreview(path);
            try
            {
                block.PlainText = FragmentFiles.ReadPlainText(path);
            }
            catch (Exception ex)
            {
                Log($"warning: could not read fragment '{block.FragmentFile}': {ex.Message}");
            }
        }
    }

    // ------------------------------------------------- Word paragraph previews

    /// <summary>
    /// Draw paragraph previews by asking Word to lay them out, instead of
    /// approximating Word's styles in WPF. Costs a hidden Word instance.
    /// </summary>
    public bool WordPreviewsEnabled
    {
        get => _wordPreviewsEnabled;
        set
        {
            if (_wordPreviewsEnabled == value) return;
            _wordPreviewsEnabled = value;
            OnPropertyChanged();
            UserSettings.Update(s => s with { WordPreviews = value });

            if (value)
            {
                Log("Word previews on — paragraphs are drawn by Word, using the template's own styles.");
                QueueAllPreviews();
            }
            else
            {
                StopWordPreviews();
                Log("Word previews off.");
            }
        }
    }

    private void QueuePreview(ParagraphBlock block)
    {
        if (!WordPreviewsEnabled) return;
        if (!_previewQueue.Contains(block)) _previewQueue.Add(block);
        _previewTimer.Stop();
        _previewTimer.Start();
    }

    private void QueueAllPreviews()
    {
        foreach (var paragraph in Program.AllBlocks().OfType<ParagraphBlock>())
            QueuePreview(paragraph);
    }

    private void RenderQueuedPreviews()
    {
        var pending = _previewQueue.ToList();
        _previewQueue.Clear();
        if (!WordPreviewsEnabled || pending.Count == 0) return;

        _wordPreviews ??= new WordPreviewService(_dispatcher, Log);
        foreach (var block in pending)
            _wordPreviews.Request(block, PreviewRequest(block),
                (b, png) => b.PreviewImage = png.Length > 0 ? png : null);
    }

    /// <summary>
    /// What this paragraph should look like — including the template of the
    /// document it sits in, so the preview shows that template's real styles.
    /// </summary>
    private ParagraphPreviewRequest PreviewRequest(ParagraphBlock block) =>
        new(EnclosingTemplatePath(block), block.TextTemplate, block.Style, block.Bold, block.Italic);

    /// <summary>
    /// The .dotx/.docx of the nearest enclosing 'new document' block, if it names one
    /// that exists. A template path built from {expressions} cannot be resolved with
    /// no record in scope, so those fall back to Word's blank document.
    /// </summary>
    private string? EnclosingTemplatePath(Block block)
    {
        for (var owner = block.ParentCollection?.Owner; owner is not null; owner = owner.ParentCollection?.Owner)
        {
            if (owner is not NewDocumentBlock document) continue;
            var template = document.TemplatePath;
            if (string.IsNullOrWhiteSpace(template) || template.Contains('{')) return null;

            var path = Path.GetFullPath(Path.Combine(RunBaseFolder, template.Trim()));
            return File.Exists(path) ? path : null;
        }
        return null;
    }

    private void StopWordPreviews()
    {
        _previewQueue.Clear();
        _previewTimer.Stop();
        _wordPreviews?.Dispose();
        _wordPreviews = null;
        foreach (var paragraph in Program.AllBlocks().OfType<ParagraphBlock>())
            paragraph.PreviewImage = null;
    }

    /// <summary>Closes the hidden Word instance when the editor shuts down.</summary>
    public void Shutdown()
    {
        _wordPreviews?.Dispose();
        _wordPreviews = null;
    }

    // --------------------------------------------------- data source columns

    /// <summary>Schedules a (debounced) re-read of the file a block points at.</summary>
    private void QueueSchemaRead(FileSourceBlockBase block)
    {
        if (!_schemaQueue.Contains(block)) _schemaQueue.Add(block);
        _schemaTimer.Stop();
        _schemaTimer.Start();
    }

    private void ReadQueuedSchemas()
    {
        var pending = _schemaQueue.ToList();
        _schemaQueue.Clear();
        foreach (var block in pending) _ = ReadSchemaAsync(block);
    }

    /// <summary>
    /// Reads the columns a source produces — out of the file for CSV and Excel, out
    /// of the query's result schema for a database — so the block can show them.
    /// The read itself runs off the UI thread: files can be big and servers slow.
    /// </summary>
    private async Task ReadSchemaAsync(SourceBlockBase? block)
    {
        if (block is null || !_schemaReading.Add(block)) return;
        try
        {
            var read = PrepareRead(block);
            if (read is null) return;

            block.SchemaStatus = "reading…";
            try
            {
                var schema = await Task.Run(read);
                SetSchema(block, schema, schema.Summary);
            }
            catch (Exception ex)
            {
                SetSchema(block, null, FirstLine(ex.Message));
            }
        }
        finally
        {
            _schemaReading.Remove(block);
        }
    }

    /// <summary>
    /// Everything the read needs, captured on the UI thread. Null when there is
    /// nothing to read yet — the block says why.
    /// </summary>
    private Func<SourceSchema>? PrepareRead(SourceBlockBase block)
    {
        switch (block)
        {
            case FileSourceBlockBase file:
                {
                    if (string.IsNullOrWhiteSpace(file.FilePath))
                    {
                        SetSchema(file, null, "");
                        return null;
                    }

                    var path = Path.GetFullPath(Path.Combine(RunBaseFolder, file.FilePath.Trim()));
                    if (!File.Exists(path))
                    {
                        SetSchema(file, null, CurrentFilePath is null && !Path.IsPathRooted(file.FilePath)
                            ? "save the program next to the data file, or use a full path"
                            : $"file not found: {path}");
                        return null;
                    }

                    var headerRow = file.HeaderRow;
                    if (file is ExcelSourceBlock excel)
                    {
                        var sheetName = excel.SheetName;
                        return () => DataSourceLoaders.PeekExcel(path, sheetName, headerRow);
                    }
                    return () => DataSourceLoaders.PeekCsv(path, headerRow);
                }

            case DatabaseSourceBlock db:
                {
                    var provider = db.Provider;
                    var connectionString = db.ConnectionString;
                    var query = db.Query;
                    return () => DataSourceLoaders.PeekDatabase(provider, connectionString, query);
                }

            default:
                return null;
        }
    }

    private static void SetSchema(SourceBlockBase block, SourceSchema? schema, string status)
    {
        block.DetectedColumns = schema?.Columns ?? [];
        if (schema is not null && block is ExcelSourceBlock excel) excel.DetectedSheets = schema.SheetNames;
        block.SchemaStatus = status;
    }

    /// <summary>Database errors are paragraphs; the block only has room for the gist.</summary>
    private static string FirstLine(string message)
    {
        var line = message.Split('\n')[0].Trim();
        return line.Length > 200 ? line[..200] + "…" : line;
    }

    /// <summary>Drops a column reference into whichever block input was last used.</summary>
    private void InsertColumnReference(SourceColumnInfo? column)
    {
        if (column is null) return;
        if (InputFocus.TryInsert(column.Reference)) return;

        try
        {
            Clipboard.SetText(column.Reference);
            Log($"Copied '{column.Reference}' to the clipboard — click inside a block input first to insert it there.");
        }
        catch (Exception ex)
        {
            Log($"warning: could not copy '{column.Reference}': {ex.Message}");
        }
    }

    private void DeleteSelected()
    {
        var block = SelectedBlock;
        if (block?.ParentCollection is { } parent)
        {
            parent.Remove(block);
            SelectedBlock = null;
        }
    }

    // ------------------------------------------------------------------- run

    private string RunBaseFolder =>
        CurrentFilePath is not null
            ? Path.GetDirectoryName(CurrentFilePath)!
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "MergeOnSteroids");

    private string ResolvedOutputFolder
    {
        get
        {
            var output = string.IsNullOrWhiteSpace(Program.OutputFolder) ? "output" : Program.OutputFolder;
            return Path.GetFullPath(Path.IsPathRooted(output) ? output : Path.Combine(RunBaseFolder, output));
        }
    }

    private async Task RunAsync(bool useWord)
    {
        if (IsRunning) return;
        IsRunning = true;
        _runCts = new CancellationTokenSource();
        OutputFiles.Clear();
        Log($"=== {(useWord ? "Word" : "Preview")} run — {DateTime.Now:T} ===");

        var program = Program;
        var options = new RunOptions
        {
            BaseFolder = RunBaseFolder,
            OutputFolder = ResolvedOutputFolder,
            ShowWord = ShowWordWindow,
            Cancellation = _runCts.Token
        };

        try
        {
            var result = await RunOnStaThread(() =>
            {
                using IDocumentWriter writer = useWord ? new WordComWriter() : new PreviewWriter();
                var interpreter = new Interpreter(writer, options, Log);
                return interpreter.Run(program);
            });

            foreach (var f in result.OutputFiles) OutputFiles.Add(f);
        }
        catch (Exception ex)
        {
            Log($"ERROR: {ex.Message}");
        }
        finally
        {
            IsRunning = false;
            _runCts = null;
        }
    }

    private static Task<RunResult> RunOnStaThread(Func<RunResult> work)
    {
        var tcs = new TaskCompletionSource<RunResult>();
        var thread = new Thread(() =>
        {
            try { tcs.SetResult(work()); }
            catch (Exception ex) { tcs.SetException(ex); }
        })
        {
            IsBackground = true
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return tcs.Task;
    }

    private void OpenOutputFolder()
    {
        var folder = ResolvedOutputFolder;
        Directory.CreateDirectory(folder);
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(folder) { UseShellExecute = true });
    }

    public void Log(string message)
    {
        void Append()
        {
            _log.AppendLine(message);
            OnPropertyChanged(nameof(LogText));
        }
        if (_dispatcher.CheckAccess()) Append();
        else _dispatcher.BeginInvoke(Append);
    }

    private static string SanitizeName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var clean = new string(name.Select(c => invalid.Contains(c) ? '_' : c).ToArray()).Trim();
        return clean.Length == 0 ? "program" : clean;
    }

    // ---------------------------------------------------------------- palette

    private static List<PaletteItem> BuildPalette()
    {
        // resource keys, not colors: the chips follow the active theme
        const string data = "BlockDataBrush";
        const string control = "BlockControlBrush";
        const string variables = "BlockVariablesBrush";
        const string document = "BlockDocumentBrush";
        const string folder = "BlockFolderBrush";

        return
        [
            new("Data sources", "open CSV file", "Loads a delimited text file as a named data source. " +
                "Say which line holds the column names, and the block lists the columns it found — " +
                "click one to use it in the input you are typing in.",
                data, () => new CsvSourceBlock { Name = "source1" }),
            new("Data sources", "open Excel file", "Loads an Excel worksheet as a named data source. " +
                "Say which row holds the column names, and the block lists the columns it found — " +
                "click one to use it in the input you are typing in.",
                data, () => new ExcelSourceBlock { Name = "source1" }),
            new("Data sources", "open database query", "Runs a SQL query (SQL Server or SQLite) as a named data source. " +
                "Press ⟳ and the block lists the columns the query returns — " +
                "click one to use it in the input you are typing in.",
                data, () => new DatabaseSourceBlock { Name = "source1" }),
            new("Data sources", "filter data source", "New source with only the rows matching a condition — " +
                "use it inside a loop to get the records related to the current one.",
                data, () => new FilterSourceBlock { Name = "filtered1" }),

            new("Control", "for each record", "Repeats its contents once per record of a data source.",
                control, () => new ForEachBlock()),
            new("Control", "if / else", "Runs its contents only when the condition is true.",
                control, () => new IfBlock { Condition = "" }),
            new("Control", "switch", "Picks one branch by value: fill it with 'case' blocks, " +
                "and whatever none of them matched runs under 'otherwise'.",
                control, () => new SwitchBlock()),
            new("Control", "case", "One branch of a switch. Answers to a value, or to several " +
                "separated by commas.",
                control, () => new CaseBlock()),

            new("Variables", "set variable", "Computes a value and stores it under a name.",
                variables, () => new SetVariableBlock()),

            new("Folders", "make folder", "Creates a folder; every document produced inside the block " +
                "is saved there. Put it in a loop for one folder per record.",
                folder, () => new MakeDirectoryBlock { FolderName = "Documents" }),
            new("Folders", "zip folder", "Same as 'make folder', but when the block ends the folder is " +
                "zipped up and (unless you keep it) removed — one archive per run, per customer, per month.",
                folder, () => new ZipDirectoryBlock { FolderName = "Archive" }),

            new("Document", "new document", "Starts a Word document; it is saved when the block ends. " +
                "Put it inside a loop to get one document per record.",
                document, () => new NewDocumentBlock()),
            // 'add paragraph' is deliberately not offered any more: writing a paragraph
            // here meant naming a Word style and hoping, when the block below lets you
            // author the real thing in Word. ParagraphBlock still exists, so programs
            // written before this keep loading and running unchanged.
            new("Document", "Word paragraphs", "Paragraphs authored directly in Word — " +
                "styles, colors, bullets, everything. {expressions} in the text are substituted at run time.",
                document, () => new WordFragmentBlock()),
            new("Document", "add table", "Adds a table filled from a data source.",
                document, () => new TableBlock()),
            new("Document", "page break", "Inserts a page break.",
                document, () => new PageBreakBlock()),
        ];
    }

    // ------------------------------------------------------------------ theme

    public RelayCommand ToggleThemeCommand { get; }

    /// <summary>Caption of the toolbar toggle — it names the theme you would switch to.</summary>
    public string ThemeToggleText => ThemeManager.IsDark ? "☀ Light" : "🌙 Dark";

    private void OnThemeChanged()
    {
        foreach (var item in Palette) item.RefreshBrush();
        OnPropertyChanged(nameof(ThemeToggleText));
    }

    // ------------------------------------------------------------------ INPC

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
