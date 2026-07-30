using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;
using MailOnSteroids.App.Common;
using MailOnSteroids.Core;
using MailOnSteroids.Core.Blocks;
using MailOnSteroids.Core.Runtime;
using MailOnSteroids.Core.Samples;
using Microsoft.Win32;

namespace MailOnSteroids.App.ViewModels;

public sealed class MainViewModel : INotifyPropertyChanged
{
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
        $"Mail on Steroids — {(CurrentFilePath is null ? "unsaved program" : Path.GetFileName(CurrentFilePath))}{(IsDirty ? " *" : "")}";

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
            Filter = "Mail on Steroids program (*.mos.json)|*.mos.json|JSON files (*.json)|*.json|All files (*.*)|*.*"
        };
        if (dialog.ShowDialog() != true) return;
        try
        {
            LoadProgram(ProgramSerializer.Load(dialog.FileName), dialog.FileName);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Could not open program:\n{ex.Message}", "Mail on Steroids",
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
                Filter = "Mail on Steroids program (*.mos.json)|*.mos.json",
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
            MessageBox.Show($"Could not save program:\n{ex.Message}", "Mail on Steroids",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void LoadSample()
    {
        if (!ConfirmDiscard()) return;
        try
        {
            var folder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "MailOnSteroids", "Sample");
            var path = SampleFactory.CreateSample(folder);
            LoadProgram(ProgramSerializer.Load(path), path);
            Log($"Sample created in {folder}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Could not create sample:\n{ex.Message}", "Mail on Steroids",
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
            MessageBox.Show($"Could not open program:\n{ex.Message}", "Mail on Steroids",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public bool ConfirmDiscard()
    {
        if (!IsDirty) return true;
        var answer = MessageBox.Show(
            "The current program has unsaved changes. Discard them?", "Mail on Steroids",
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
        IsDirty = false;
        RefreshSources();
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
    /// The heart of the fragment experiment: open the fragment in a real Word
    /// window, let the user edit there, capture the result back into the block
    /// (formatted XML + Word-rendered preview image).
    /// </summary>
    private void EditFragment(WordFragmentBlock? block)
    {
        if (block is null) return;

        Core.Interop.WordFragmentEditSession? session;
        try
        {
            System.Windows.Input.Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            session = Core.Interop.WordFragmentEditSession.Start(
                block.HasContent ? block.FragmentXml : null);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Could not start Microsoft Word:\n{ex.Message}", "Mail on Steroids",
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
            {
                var capture = session.Capture();
                block.FragmentXml = capture.Xml;
                block.PlainText = capture.PlainText;
                block.PreviewPng = capture.PreviewPng.Length > 0 ? capture.PreviewPng : null;
                Log($"Fragment updated ({capture.PlainText.Split('\n').Length} paragraph(s)).");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Could not read the fragment back from Word (was the document closed?):\n{ex.Message}",
                "Mail on Steroids", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        finally
        {
            try { session.Dispose(); } catch { /* Word may already be gone */ }
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
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "MailOnSteroids");

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
        var data = Brush("#59C059");
        var control = Brush("#FFAB19");
        var variables = Brush("#FF8C1A");
        var document = Brush("#4C97FF");

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

    private static SolidColorBrush Brush(string hex)
    {
        var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
        brush.Freeze();
        return brush;
    }

    // ------------------------------------------------------------------ INPC

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
