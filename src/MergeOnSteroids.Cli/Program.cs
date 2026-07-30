using MergeOnSteroids.Core;
using MergeOnSteroids.Core.Runtime;
using MergeOnSteroids.Core.Samples;

namespace MergeOnSteroids.Cli;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            return Run(args);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("ERROR: " + ex.Message);
            return 1;
        }
    }

    private static int Run(string[] args)
    {
        if (args.Length == 0)
        {
            PrintUsage();
            return 1;
        }

        switch (args[0].ToLowerInvariant())
        {
            case "sample":
                {
                    var folder = args.Length > 1 ? args[1] : Path.Combine(Environment.CurrentDirectory, "sample");
                    var path = SampleFactory.CreateSample(folder);
                    Console.WriteLine($"Sample created: {path}");
                    Console.WriteLine($"Try: mos run \"{path}\"");
                    return 0;
                }

            case "richsample":
                {
                    var folder = args.Length > 1 ? args[1] : Path.Combine(Environment.CurrentDirectory, "rich-sample");
                    Console.WriteLine("Authoring a formatted fragment through Word…");
                    var path = SampleFactory.CreateRichSample(folder);
                    Console.WriteLine($"Rich sample created: {path}");
                    Console.WriteLine($"Try: mos run \"{path}\" --word");
                    return 0;
                }

            case "run":
                {
                    if (args.Length < 2)
                    {
                        PrintUsage();
                        return 1;
                    }
                    var programPath = Path.GetFullPath(args[1]);
                    var useWord = args.Contains("--word", StringComparer.OrdinalIgnoreCase);
                    var showWord = args.Contains("--show", StringComparer.OrdinalIgnoreCase);
                    var outIndex = Array.FindIndex(args, a => a.Equals("--out", StringComparison.OrdinalIgnoreCase));
                    var outOverride = outIndex >= 0 && outIndex + 1 < args.Length ? args[outIndex + 1] : null;

                    var program = ProgramSerializer.Load(programPath);
                    var baseFolder = Path.GetDirectoryName(programPath)!;
                    var outputFolder = outOverride
                        ?? (Path.IsPathRooted(program.OutputFolder)
                            ? program.OutputFolder
                            : Path.Combine(baseFolder, program.OutputFolder));

                    var options = new RunOptions
                    {
                        BaseFolder = baseFolder,
                        OutputFolder = Path.GetFullPath(outputFolder),
                        ShowWord = showWord
                    };

                    using IDocumentWriter writer = useWord ? new WordComWriter() : new PreviewWriter();
                    var interpreter = new Interpreter(writer, options, Console.WriteLine);
                    var result = interpreter.Run(program);
                    return result.Succeeded ? 0 : 2;
                }

            default:
                PrintUsage();
                return 1;
        }
    }

    private static void PrintUsage()
    {
        Console.WriteLine("""
            Merge on Steroids — command line runner

            Usage:
              mos sample [folder]                    Create sample data + program
              mos richsample [folder]                Create sample using a Word-authored rich fragment (needs Word)
              mos run <program.mos.json> [options]   Execute a program

            Options for run:
              --word        Generate real .docx files through Microsoft Word (default: text preview)
              --show        Make the Word window visible during generation
              --out <dir>   Override the program's output folder
            """);
    }
}
