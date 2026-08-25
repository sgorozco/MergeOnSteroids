using System.Text;
using MergeOnSteroids.Core.Blocks;
using MergeOnSteroids.Core.Fragments;
using MergeOnSteroids.Core.Interop;

namespace MergeOnSteroids.Core.Samples;

/// <summary>
/// Generates a self-contained demo: sample CSV data plus a program that produces
/// one account-statement letter per customer, with a filtered order table,
/// computed totals, and conditional paragraphs.
/// </summary>
public static class SampleFactory
{
    /// <summary>Writes sample data + program into <paramref name="folder"/> and returns the program path.</summary>
    public static string CreateSample(string folder)
    {
        Directory.CreateDirectory(folder);
        WriteCustomersCsv(Path.Combine(folder, "customers.csv"));
        WriteOrdersCsv(Path.Combine(folder, "orders.csv"));

        var program = BuildSampleProgram();
        var path = Path.Combine(folder, "statement-letters" + ProgramSerializer.FileExtension);
        ProgramSerializer.Save(program, path);
        return path;
    }

    public static ProgramModel BuildSampleProgram()
    {
        var program = new ProgramModel
        {
            Name = "Customer statement letters",
            OutputFolder = "output"
        };

        program.Blocks.Add(new CsvSourceBlock { Name = "customers", FilePath = "customers.csv" });
        program.Blocks.Add(new CsvSourceBlock { Name = "orders", FilePath = "orders.csv" });

        var loop = new ForEachBlock { Alias = "c", SourceName = "customers" };
        program.Blocks.Add(loop);

        loop.Children.Add(new FilterSourceBlock
        {
            Name = "custOrders",
            SourceName = "orders",
            Condition = "CustomerId = c.Id"
        });

        var doc = new NewDocumentBlock { FileNameTemplate = "Statement_{c.Company}" };
        loop.Children.Add(doc);

        doc.Children.Add(new ParagraphBlock { TextTemplate = "ACME Distribution Inc.", Style = ParagraphStyles.Title });
        doc.Children.Add(new ParagraphBlock
        {
            TextTemplate = "Account statement — {c.Company}",
            Style = ParagraphStyles.Heading1
        });
        doc.Children.Add(new ParagraphBlock { TextTemplate = "Generated on {FORMAT(TODAY(), \"D\")}" });
        doc.Children.Add(new ParagraphBlock { TextTemplate = "" });
        doc.Children.Add(new ParagraphBlock { TextTemplate = "Dear {c.Contact}," });

        var hasOrders = new IfBlock { Condition = "COUNT(custOrders) > 0" };
        doc.Children.Add(hasOrders);

        hasOrders.Children.Add(new ParagraphBlock
        {
            TextTemplate = "Here is a summary of your {COUNT(custOrders)} order(s) this period:"
        });
        hasOrders.Children.Add(new TableBlock
        {
            SourceName = "custOrders",
            ColumnsSpec = "Date: FORMAT(Date, \"d\") | Product: Product | Qty: Qty | " +
                          "Unit price: FORMAT(UnitPrice, \"C\") | Total: FORMAT(Qty * UnitPrice, \"C\")"
        });
        hasOrders.Children.Add(new ParagraphBlock
        {
            TextTemplate = "Grand total: {FORMAT(SUM(custOrders, Qty * UnitPrice), \"C\")}",
            Bold = true
        });
        hasOrders.Else.Add(new ParagraphBlock
        {
            TextTemplate = "You have no orders this period — we would love to hear from you!",
            Italic = true
        });

        var vip = new IfBlock { Condition = "SUM(custOrders, Qty * UnitPrice) > 10000" };
        doc.Children.Add(vip);
        vip.Children.Add(new ParagraphBlock
        {
            TextTemplate = "As one of our preferred customers, you qualify for a dedicated account manager. " +
                           "Expect a call from us soon.",
            Style = ParagraphStyles.Quote
        });

        doc.Children.Add(new ParagraphBlock { TextTemplate = "" });
        doc.Children.Add(new ParagraphBlock { TextTemplate = "Kind regards," });
        doc.Children.Add(new ParagraphBlock { TextTemplate = "The ACME team — {c.City} office", Italic = true });

        return program;
    }

    /// <summary>
    /// Rich-fragment demo: authors a formatted Word fragment through Word itself
    /// (styles, bold, color, placeholders), captures it, and builds a program that
    /// stamps one letter per customer from it. Requires Microsoft Word.
    /// </summary>
    public static string CreateRichSample(string folder)
    {
        Directory.CreateDirectory(folder);
        WriteCustomersCsv(Path.Combine(folder, "customers.csv"));
        WriteOrdersCsv(Path.Combine(folder, "orders.csv"));

        var relativeFragment = $"{FragmentFiles.FolderName}/carta-saldo.docx";
        var fragmentPath = Path.Combine(folder, relativeFragment.Replace('/', Path.DirectorySeparatorChar));
        var capture = AuthorRichFragment(fragmentPath);
        FragmentFiles.WritePreview(fragmentPath, capture.PreviewPng);

        var program = new ProgramModel { Name = "Rich fragment letters", OutputFolder = "output" };
        program.Blocks.Add(new CsvSourceBlock { Name = "customers", FilePath = "customers.csv" });

        var loop = new ForEachBlock { Alias = "c", SourceName = "customers" };
        program.Blocks.Add(loop);

        var doc = new NewDocumentBlock { FileNameTemplate = "Carta_{c.Company}" };
        loop.Children.Add(doc);

        doc.Children.Add(new WordFragmentBlock
        {
            FragmentFile = relativeFragment,
            PlainText = capture.PlainText,
            PreviewImage = capture.PreviewPng.Length > 0 ? capture.PreviewPng : null
        });
        doc.Children.Add(new ParagraphBlock { TextTemplate = "" });
        doc.Children.Add(new ParagraphBlock
        {
            TextTemplate = "— generated by Merge on Steroids on {FORMAT(TODAY(), \"d\")}",
            Italic = true
        });

        var path = Path.Combine(folder, "rich-letters" + ProgramSerializer.FileExtension);
        ProgramSerializer.Save(program, path);
        return path;
    }

    private static FragmentCapture AuthorRichFragment(string docxPath)
    {
        // the sample generator is a one-shot, so it owns its own Word for the duration
        using var word = new WordApplication();
        using var session = WordFragmentEditSession.Start(word, null, visible: false);
        dynamic doc = session.Document;

        AddStyledParagraph(doc, "Estado de cuenta — {c.Company}", -3 /* Heading 2 */);
        AddStyledParagraph(doc, "Estimado {c.Contact}:", -1 /* Normal */);

        // Mixed inline formatting: the balance placeholder gets bold + dark red,
        // proving formatting survives capture, insertion, and substitution.
        var text = "Le informamos que su saldo actual es de {FORMAT(c.Balance, \"C\")}, " +
                   "registrado en nuestra oficina de {UPPER(c.City)}.";
        var start = AddStyledParagraph(doc, text, -1);
        var marker = "{FORMAT(c.Balance, \"C\")}";
        var idx = text.IndexOf(marker, StringComparison.Ordinal);
        dynamic highlight = doc.Range(start + idx, start + idx + marker.Length);
        highlight.Font.Bold = 1;
        highlight.Font.Color = 0x000080; // dark red (BGR)

        AddStyledParagraph(doc, "Gracias por su preferencia.", -1);

        return session.SaveAs(docxPath);
    }

    /// <summary>Appends a styled paragraph; returns the start position of its text.</summary>
    private static int AddStyledParagraph(dynamic doc, string text, int styleId)
    {
        dynamic range = doc.Bookmarks["\\endofdoc"].Range;
        int start = (int)range.Start;
        range.Text = text;
        range.Style = styleId;
        range.InsertParagraphAfter();
        return start;
    }

    private static void WriteCustomersCsv(string path)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Id,Company,Contact,City,Balance");
        sb.AppendLine("1,Contoso Ltd,Maria Lopez,Monterrey,15200.50");
        sb.AppendLine("2,Fabrikam SA,Juan Perez,Guadalajara,830.00");
        sb.AppendLine("3,Wingtip Toys,Ana Torres,CDMX,0");
        sb.AppendLine("4,Tailspin MX,Luis Romero,Monterrey,42100.75");
        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
    }

    private static void WriteOrdersCsv(string path)
    {
        var sb = new StringBuilder();
        sb.AppendLine("OrderId,CustomerId,Date,Product,Qty,UnitPrice");
        sb.AppendLine("1001,1,2026-06-03,Steel brackets,120,18.50");
        sb.AppendLine("1002,1,2026-06-17,Hex bolts M8,500,0.85");
        sb.AppendLine("1003,2,2026-06-21,Paint 20L,12,64.00");
        sb.AppendLine("1004,4,2026-07-02,Conveyor belt,3,2999.99");
        sb.AppendLine("1005,4,2026-07-08,Roller set,40,155.25");
        sb.AppendLine("1006,1,2026-07-15,Angle grinder,6,899.00");
        sb.AppendLine("1007,4,2026-07-20,Safety kit,25,210.00");
        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
    }
}
