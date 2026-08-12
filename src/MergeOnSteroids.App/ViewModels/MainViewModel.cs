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
using MergeOnSteroids.Core.Fragments;
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

        ThemeManager.ThemeChanged += OnThemeChanged;

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
        IsDirty = true;
        if (e.PropertyName == nameof(SourceBlockBase.Name))
            RefreshSources();
    }

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

        return
        [
            new("Data sources", "open CSV file", "Loads a CSV file as a named data source.",
                data, () => new CsvSourceBlock { Name = "source1" }),
            new("Data sources", "open Excel file", "Loads an Excel worksheet as a named data source.",
                data, () => new ExcelSourceBlock { Name = "source1" }),
            new("Data sources", "open database query", "Runs a SQL query (SQL Server or SQLite) as a named data source.",
                data, () => new DatabaseSourceBlock { Name = "source1" }),
            new("Data sources", "filter data source", "New source with only the rows matching a condition — " +
                "use it inside a loop to get the records related to the current one.",
                data, () => new FilterSourceBlock { Name = "filtered1" }),

            new("Control", "for each record", "Repeats its contents once per record of a data source.",
                control, () => new ForEachBlock()),
            new("Control", "if / else", "Runs its contents only when the condition is true.",
                control, () => new IfBlock { Condition = "" }),

            new("Variables", "set variable", "Computes a value and stores it under a name.",
                variables, () => new SetVariableBlock()),

            new("Document", "new document", "Starts a Word document; it is saved when the block ends. " +
                "Put it inside a loop to get one document per record.",
                document, () => new NewDocumentBlock()),
            new("Document", "add paragraph", "Adds a paragraph. Use {expressions} to insert data.",
                document, () => new ParagraphBlock { TextTemplate = "Text with {placeholders}" }),
            new("Document", "Word paragraphs (rich)", "Formatted paragraphs authored directly in Word — " +
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
