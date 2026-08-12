using System.Globalization;
using System.Text;
using MergeOnSteroids.Core.Data;

namespace MergeOnSteroids.Core.Expressions;

public sealed class MosExpressionException(string message) : Exception(message);

/// <summary>
/// Everything an expression can see while evaluating: loop records (by alias or bare
/// field name), variables, and named data sources (for aggregates like SUM/COUNT).
/// </summary>
public interface IEvalContext
{
    bool TryGetVariable(string name, out object? value);
    bool TryGetBareField(string field, out object? value);
    bool TryGetAliasField(string alias, string field, out object? value);
    bool IsAlias(string name);
    DataTableLite? ResolveSource(string name);
    /// <summary>Push a row as the innermost bare-field scope (used by aggregates/filters). Dispose to pop.</summary>
    IDisposable PushRowScope(DataTableLite table, int rowIndex);
}

public abstract class ExprNode
{
    public abstract object? Eval(IEvalContext ctx);
}

// ===========================================================================
// Compiler (tokenizer + recursive descent parser) with per-string caching
// ===========================================================================

public sealed class ExpressionCompiler
{
    private readonly Dictionary<string, ExprNode> _cache = new(StringComparer.Ordinal);

    public ExprNode Compile(string expression)
    {
        if (string.IsNullOrWhiteSpace(expression))
            throw new MosExpressionException("Empty expression.");
        if (_cache.TryGetValue(expression, out var cached)) return cached;
        // Word autocorrect turns straight quotes into curly ones while typing
        // inside a document — accept both.
        var normalized = expression
            .Replace('“', '"').Replace('”', '"')
            .Replace('‘', '\'').Replace('’', '\'');
        var node = new Parser(normalized).ParseFull();
        _cache[expression] = node;
        return node;
    }

    // ------------------------------------------------------------ tokenizer

    // QuotedIdent is a [bracketed] name: always a field, never a keyword — so a
    // column really called "Null" or "And" can still be read as [Null], [And].
    private enum TokKind { Ident, QuotedIdent, Number, String, Symbol, End }

    private readonly record struct Token(TokKind Kind, string Text, int Pos);

    private sealed class Parser(string source)
    {
        private readonly string _src = source;
        private int _pos;
        private Token _current;

        public ExprNode ParseFull()
        {
            Next();
            var node = ParseOr();
            Expect(TokKind.End, "end of expression");
            return node;
        }

        // -------------------------------------------------------- grammar

        private ExprNode ParseOr()
        {
            var left = ParseAnd();
            while (IsKeyword("OR"))
            {
                Next();
                var right = ParseAnd();
                left = new LogicalNode(left, right, isAnd: false);
            }
            return left;
        }

        private ExprNode ParseAnd()
        {
            var left = ParseNot();
            while (IsKeyword("AND"))
            {
                Next();
                var right = ParseNot();
                left = new LogicalNode(left, right, isAnd: true);
            }
            return left;
        }

        private ExprNode ParseNot()
        {
            if (IsKeyword("NOT"))
            {
                Next();
                return new NotNode(ParseNot());
            }
            return ParseComparison();
        }

        private ExprNode ParseComparison()
        {
            var left = ParseAdditive();
            if (_current.Kind == TokKind.Symbol &&
                _current.Text is "=" or "==" or "!=" or "<>" or "<" or "<=" or ">" or ">=")
            {
                var op = _current.Text;
                Next();
                var right = ParseAdditive();
                return new ComparisonNode(left, right, op);
            }
            return left;
        }

        private ExprNode ParseAdditive()
        {
            var left = ParseMultiplicative();
            while (_current.Kind == TokKind.Symbol && _current.Text is "+" or "-" or "&")
            {
                var op = _current.Text;
                Next();
                var right = ParseMultiplicative();
                left = new ArithmeticNode(left, right, op);
            }
            return left;
        }

        private ExprNode ParseMultiplicative()
        {
            var left = ParseUnary();
            while (_current.Kind == TokKind.Symbol && _current.Text is "*" or "/" or "%")
            {
                var op = _current.Text;
                Next();
                var right = ParseUnary();
                left = new ArithmeticNode(left, right, op);
            }
            return left;
        }

        private ExprNode ParseUnary()
        {
            if (_current.Kind == TokKind.Symbol && _current.Text == "-")
            {
                Next();
                return new NegateNode(ParseUnary());
            }
            return ParsePrimary();
        }

        private ExprNode ParsePrimary()
        {
            switch (_current.Kind)
            {
                case TokKind.Number:
                    {
                        var value = decimal.Parse(_current.Text, CultureInfo.InvariantCulture);
                        Next();
                        return new ConstNode(value);
                    }
                case TokKind.String:
                    {
                        var value = _current.Text;
                        Next();
                        return new ConstNode(value);
                    }
                case TokKind.Ident when IsKeyword("TRUE"): Next(); return new ConstNode(true);
                case TokKind.Ident when IsKeyword("FALSE"): Next(); return new ConstNode(false);
                case TokKind.Ident when IsKeyword("NULL"): Next(); return new ConstNode(null);
                case TokKind.Ident or TokKind.QuotedIdent:
                    {
                        var name = _current.Text;
                        var wasQuoted = _current.Kind == TokKind.QuotedIdent;
                        Next();
                        if (!wasQuoted && _current.Kind == TokKind.Symbol && _current.Text == "(")
                            return ParseCall(name);

                        var segments = new List<string> { name };
                        while (_current.Kind == TokKind.Symbol && _current.Text == ".")
                        {
                            Next();
                            if (_current.Kind is not (TokKind.Ident or TokKind.QuotedIdent))
                                throw Error($"Expected a field name after '{string.Join(".", segments)}.'");
                            segments.Add(_current.Text);
                            Next();
                        }
                        return new PathNode(segments);
                    }
                case TokKind.Symbol when _current.Text == "(":
                    {
                        Next();
                        var inner = ParseOr();
                        ExpectSymbol(")");
                        return inner;
                    }
                default:
                    throw Error($"Unexpected token '{_current.Text}'.");
            }
        }

        private ExprNode ParseCall(string name)
        {
            ExpectSymbol("(");
            var args = new List<ExprNode>();
            if (!(_current.Kind == TokKind.Symbol && _current.Text == ")"))
            {
                while (true)
                {
                    args.Add(ParseOr());
                    if (_current.Kind == TokKind.Symbol && _current.Text == ",") { Next(); continue; }
                    break;
                }
            }
            ExpectSymbol(")");
            return new FuncNode(name, args);
        }

        // ------------------------------------------------------- token feed

        private bool IsKeyword(string kw) =>
            _current.Kind == TokKind.Ident && _current.Text.Equals(kw, StringComparison.OrdinalIgnoreCase);

        private void Expect(TokKind kind, string what)
        {
            if (_current.Kind != kind)
                throw Error($"Expected {what} but found '{_current.Text}'.");
        }

        private void ExpectSymbol(string sym)
        {
            if (_current.Kind != TokKind.Symbol || _current.Text != sym)
                throw Error($"Expected '{sym}' but found '{_current.Text}'.");
            Next();
        }

        private MosExpressionException Error(string message) =>
            new($"{message} (position {_current.Pos + 1} in \"{_src}\")");

        private void Next()
        {
            while (_pos < _src.Length && char.IsWhiteSpace(_src[_pos])) _pos++;
            if (_pos >= _src.Length)
            {
                _current = new Token(TokKind.End, "", _pos);
                return;
            }

            var start = _pos;
            var ch = _src[_pos];

            // Bracket-quoted identifier: [Order Date]
            if (ch == '[')
            {
                var end = _src.IndexOf(']', _pos + 1);
                if (end < 0) throw new MosExpressionException($"Unterminated '[' at position {start + 1} in \"{_src}\"");
                _current = new Token(TokKind.QuotedIdent, _src.Substring(_pos + 1, end - _pos - 1).Trim(), start);
                _pos = end + 1;
                return;
            }

            if (char.IsLetter(ch) || ch == '_')
            {
                while (_pos < _src.Length && (char.IsLetterOrDigit(_src[_pos]) || _src[_pos] == '_')) _pos++;
                _current = new Token(TokKind.Ident, _src[start.._pos], start);
                return;
            }

            if (char.IsDigit(ch))
            {
                while (_pos < _src.Length && (char.IsDigit(_src[_pos]) || _src[_pos] == '.')) _pos++;
                _current = new Token(TokKind.Number, _src[start.._pos], start);
                return;
            }

            if (ch is '"' or '\'')
            {
                var quote = ch;
                var sb = new StringBuilder();
                _pos++;
                while (true)
                {
                    if (_pos >= _src.Length)
                        throw new MosExpressionException($"Unterminated string at position {start + 1} in \"{_src}\"");
                    var c = _src[_pos];
                    if (c == quote)
                    {
                        if (_pos + 1 < _src.Length && _src[_pos + 1] == quote) { sb.Append(quote); _pos += 2; continue; }
                        _pos++;
                        break;
                    }
                    sb.Append(c);
                    _pos++;
                }
                _current = new Token(TokKind.String, sb.ToString(), start);
                return;
            }

            // multi-char operators first
            foreach (var op in (string[])["==", "!=", "<>", "<=", ">="])
            {
                if (_src.AsSpan(_pos).StartsWith(op))
                {
                    _current = new Token(TokKind.Symbol, op, start);
                    _pos += op.Length;
                    return;
                }
            }

            if ("+-*/%&=<>(),.".Contains(ch))
            {
                _current = new Token(TokKind.Symbol, ch.ToString(), start);
                _pos++;
                return;
            }

            throw new MosExpressionException($"Unexpected character '{ch}' at position {start + 1} in \"{_src}\"");
        }
    }
}

// ===========================================================================
// AST nodes
// ===========================================================================

public sealed class ConstNode(object? value) : ExprNode
{
    public override object? Eval(IEvalContext ctx) => value;
}

public sealed class NegateNode(ExprNode inner) : ExprNode
{
    public override object? Eval(IEvalContext ctx)
    {
        var v = inner.Eval(ctx);
        var n = Values.ToNumber(v) ?? throw new MosExpressionException($"Cannot negate '{Values.Describe(v)}'.");
        return -n;
    }
}

public sealed class NotNode(ExprNode inner) : ExprNode
{
    public override object? Eval(IEvalContext ctx) => !Values.Truthy(inner.Eval(ctx));
}

public sealed class LogicalNode(ExprNode left, ExprNode right, bool isAnd) : ExprNode
{
    public override object? Eval(IEvalContext ctx)
    {
        var l = Values.Truthy(left.Eval(ctx));
        if (isAnd && !l) return false;
        if (!isAnd && l) return true;
        return Values.Truthy(right.Eval(ctx));
    }
}

public sealed class ComparisonNode(ExprNode left, ExprNode right, string op) : ExprNode
{
    public override object? Eval(IEvalContext ctx)
    {
        var l = left.Eval(ctx);
        var r = right.Eval(ctx);
        switch (op)
        {
            case "=" or "==": return Values.AreEqual(l, r);
            case "!=" or "<>": return !Values.AreEqual(l, r);
        }
        var cmp = Values.Compare(l, r);
        if (cmp is null) return false; // null compared with anything (except =/<>) is false
        return op switch
        {
            "<" => cmp < 0,
            "<=" => cmp <= 0,
            ">" => cmp > 0,
            ">=" => cmp >= 0,
            _ => throw new MosExpressionException($"Unknown comparison operator '{op}'.")
        };
    }
}

public sealed class ArithmeticNode(ExprNode left, ExprNode right, string op) : ExprNode
{
    public override object? Eval(IEvalContext ctx)
    {
        var l = left.Eval(ctx);
        var r = right.Eval(ctx);

        if (op == "&")
            return Values.ToDisplayString(l) + Values.ToDisplayString(r);

        if (op == "+" && (l is string || r is string))
            return Values.ToDisplayString(l) + Values.ToDisplayString(r);

        // date + number => add days
        if (op is "+" or "-" && l is DateTime dt && Values.ToNumber(r) is { } days)
            return op == "+" ? dt.AddDays((double)days) : dt.AddDays(-(double)days);
        if (op == "-" && l is DateTime a && r is DateTime b)
            return (decimal)(a - b).TotalDays;

        var ln = Values.ToNumber(l);
        var rn = Values.ToNumber(r);
        if (ln is null || rn is null)
            throw new MosExpressionException(
                $"Operator '{op}' needs numbers but got {Values.Describe(l)} and {Values.Describe(r)}.");

        return op switch
        {
            "+" => ln + rn,
            "-" => ln - rn,
            "*" => ln * rn,
            "/" => rn == 0 ? throw new MosExpressionException("Division by zero.") : ln / rn,
            "%" => rn == 0 ? throw new MosExpressionException("Division by zero.") : ln % rn,
            _ => throw new MosExpressionException($"Unknown operator '{op}'.")
        };
    }
}

/// <summary>Identifier path: variable, bare field, or alias.Field.</summary>
public sealed class PathNode(List<string> segments) : ExprNode
{
    public IReadOnlyList<string> Segments => segments;

    public override object? Eval(IEvalContext ctx)
    {
        if (segments.Count == 1)
        {
            var name = segments[0];
            if (ctx.TryGetBareField(name, out var fv)) return fv;
            if (ctx.TryGetVariable(name, out var vv)) return vv;
            if (ctx.IsAlias(name))
                throw new MosExpressionException(
                    $"'{name}' is a record alias — use '{name}.FieldName' to read a field.");
            throw new MosExpressionException(
                $"Unknown name '{name}'. It is not a field of any record in scope, nor a variable.");
        }

        if (segments.Count == 2)
        {
            var alias = segments[0];
            var field = segments[1];
            if (ctx.IsAlias(alias))
            {
                if (ctx.TryGetAliasField(alias, field, out var value)) return value;
                throw new MosExpressionException($"Record '{alias}' has no field named '{field}'.");
            }
            throw new MosExpressionException(
                $"Unknown record alias '{alias}' in '{alias}.{field}'.");
        }

        throw new MosExpressionException(
            $"Path '{string.Join(".", segments)}' is too deep — use alias.Field.");
    }
}

public sealed class FuncNode(string name, List<ExprNode> args) : ExprNode
{
    private static readonly string[] AggregateNames = ["COUNT", "SUM", "AVG", "MIN", "MAX", "FIRST"];

    public override object? Eval(IEvalContext ctx)
    {
        var fn = name.ToUpperInvariant();
        if (AggregateNames.Contains(fn))
            return EvalAggregate(fn, ctx);

        return fn switch
        {
            "UPPER" => Str(0, ctx).ToUpperInvariant(),
            "LOWER" => Str(0, ctx).ToLowerInvariant(),
            "TRIM" => Str(0, ctx).Trim(),
            "LEN" => (decimal)Str(0, ctx).Length,
            "LEFT" => Left(Str(0, ctx), Int(1, ctx)),
            "RIGHT" => Right(Str(0, ctx), Int(1, ctx)),
            "REPLACE" => Str(0, ctx).Replace(Str(1, ctx), Str(2, ctx), StringComparison.OrdinalIgnoreCase),
            "CONTAINS" => Str(0, ctx).Contains(Str(1, ctx), StringComparison.OrdinalIgnoreCase),
            "STARTSWITH" => Str(0, ctx).StartsWith(Str(1, ctx), StringComparison.OrdinalIgnoreCase),
            "ENDSWITH" => Str(0, ctx).EndsWith(Str(1, ctx), StringComparison.OrdinalIgnoreCase),
            "CONCAT" => string.Concat(args.Select(a => Values.ToDisplayString(a.Eval(ctx)))),
            "FORMAT" => Format(Arg(0, ctx), Str(1, ctx)),
            "ROUND" => Math.Round(Num(0, ctx), args.Count > 1 ? (int)Int(1, ctx) : 0, MidpointRounding.AwayFromZero),
            "ABS" => Math.Abs(Num(0, ctx)),
            "VALUE" => Values.ToNumber(Arg(0, ctx)) ?? throw new MosExpressionException(
                $"VALUE could not convert {Values.Describe(Arg(0, ctx))} to a number."),
            "IIF" => Values.Truthy(args[0].Eval(ctx)) ? Arg(1, ctx) : Arg(2, ctx),
            "ISNULL" or "COALESCE" => args.Select(a => a.Eval(ctx)).FirstOrDefault(v => v is not null),
            "TODAY" => DateTime.Today,
            "NOW" => DateTime.Now,
            "YEAR" => (decimal)Date(0, ctx).Year,
            "MONTH" => (decimal)Date(0, ctx).Month,
            "DAY" => (decimal)Date(0, ctx).Day,
            _ => throw new MosExpressionException($"Unknown function '{name}'.")
        };
    }

    private object? EvalAggregate(string fn, IEvalContext ctx)
    {
        if (args.Count == 0)
            throw new MosExpressionException($"{fn} needs a data source name as its first argument.");

        var sourceName = args[0] switch
        {
            PathNode { Segments.Count: 1 } p => p.Segments[0],
            ConstNode c when c.Eval(ctx) is string s => s,
            _ => throw new MosExpressionException(
                $"{fn}: first argument must be a data source name, e.g. {fn}(orders, ...).")
        };

        var table = ctx.ResolveSource(sourceName)
            ?? throw new MosExpressionException($"{fn}: unknown data source '{sourceName}'.");

        if (fn == "COUNT") return (decimal)table.Rows.Count;

        if (args.Count < 2)
            throw new MosExpressionException($"{fn} needs an expression, e.g. {fn}({sourceName}, Amount).");

        var expr = args[1];
        decimal sum = 0;
        var count = 0;
        object? min = null, max = null, first = null;

        for (var i = 0; i < table.Rows.Count; i++)
        {
            using var scope = ctx.PushRowScope(table, i);
            var v = expr.Eval(ctx);
            if (v is null) continue;
            if (count == 0) first = v;
            count++;
            if (fn is "SUM" or "AVG")
            {
                sum += Values.ToNumber(v) ?? throw new MosExpressionException(
                    $"{fn}: expression produced non-numeric value {Values.Describe(v)}.");
            }
            if (min is null || Values.Compare(v, min) < 0) min = v;
            if (max is null || Values.Compare(v, max) > 0) max = v;
        }

        return fn switch
        {
            "SUM" => sum,
            "AVG" => count == 0 ? null : sum / count,
            "MIN" => min,
            "MAX" => max,
            "FIRST" => first,
            _ => null
        };
    }

    private object? Arg(int i, IEvalContext ctx)
    {
        if (i >= args.Count)
            throw new MosExpressionException($"{name}: missing argument #{i + 1}.");
        return args[i].Eval(ctx);
    }

    private string Str(int i, IEvalContext ctx) => Values.ToDisplayString(Arg(i, ctx));
    private decimal Num(int i, IEvalContext ctx) => Values.ToNumber(Arg(i, ctx))
        ?? throw new MosExpressionException($"{name}: argument #{i + 1} must be a number.");
    private decimal Int(int i, IEvalContext ctx) => Math.Truncate(Num(i, ctx));
    private DateTime Date(int i, IEvalContext ctx) => Arg(i, ctx) as DateTime?
        ?? throw new MosExpressionException($"{name}: argument #{i + 1} must be a date.");

    private static string Left(string s, decimal n) => s[..Math.Min(s.Length, Math.Max(0, (int)n))];
    private static string Right(string s, decimal n) => s[^Math.Min(s.Length, Math.Max(0, (int)n))..];

    private static string Format(object? value, string format) => value switch
    {
        null => "",
        IFormattable f => f.ToString(format, CultureInfo.CurrentCulture),
        _ => value.ToString() ?? ""
    };
}

// ===========================================================================
// Value semantics
// ===========================================================================

public static class Values
{
    public static decimal? ToNumber(object? v) => v switch
    {
        null => null,
        decimal d => d,
        double d => (decimal)d,
        float f => (decimal)f,
        int i => i,
        long l => l,
        short s => s,
        byte b => b,
        bool b => b ? 1 : 0,
        string s when decimal.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out var d) => d,
        string s when decimal.TryParse(s, NumberStyles.Number, CultureInfo.CurrentCulture, out var d) => d,
        _ => null
    };

    public static bool Truthy(object? v) => v switch
    {
        null => false,
        bool b => b,
        string s => s.Length > 0,
        _ => ToNumber(v) is { } n ? n != 0 : true
    };

    public static bool AreEqual(object? a, object? b)
    {
        if (a is null && b is null) return true;
        if (a is null || b is null) return false;
        if (ToNumber(a) is { } na && ToNumber(b) is { } nb && (IsNumeric(a) || IsNumeric(b)))
            return na == nb;
        if (a is DateTime da && b is DateTime db) return da == db;
        if (a is bool ba && b is bool bb) return ba == bb;
        return string.Equals(ToDisplayString(a), ToDisplayString(b), StringComparison.OrdinalIgnoreCase);
    }

    public static int? Compare(object? a, object? b)
    {
        if (a is null || b is null) return null;
        if (ToNumber(a) is { } na && ToNumber(b) is { } nb && (IsNumeric(a) || IsNumeric(b)))
            return na.CompareTo(nb);
        if (a is DateTime da && b is DateTime db) return da.CompareTo(db);
        return string.Compare(ToDisplayString(a), ToDisplayString(b), StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsNumeric(object v) =>
        v is decimal or double or float or int or long or short or byte;

    public static string ToDisplayString(object? v) => v switch
    {
        null => "",
        string s => s,
        bool b => b ? "True" : "False",
        DateTime dt => dt.TimeOfDay == TimeSpan.Zero
            ? dt.ToString("d", CultureInfo.CurrentCulture)
            : dt.ToString("g", CultureInfo.CurrentCulture),
        decimal d => d.ToString("0.############", CultureInfo.CurrentCulture),
        IFormattable f => f.ToString(null, CultureInfo.CurrentCulture) ?? "",
        _ => v.ToString() ?? ""
    };

    public static string Describe(object? v) =>
        v is null ? "null" : $"'{ToDisplayString(v)}' ({v.GetType().Name})";
}

// ===========================================================================
// {expression} template interpolation
// ===========================================================================

public static class TemplateEngine
{
    /// <summary>
    /// Replaces {expression} placeholders. "{{" and "}}" are literal braces.
    /// Errors become inline {!message} markers and are reported via onError.
    /// </summary>
    public static string Interpolate(
        string template, ExpressionCompiler compiler, IEvalContext ctx, Action<string>? onError = null)
    {
        if (string.IsNullOrEmpty(template) || !template.Contains('{'))
            return template ?? "";

        var sb = new StringBuilder(template.Length + 32);
        for (var i = 0; i < template.Length; i++)
        {
            var ch = template[i];
            if (ch == '{')
            {
                if (i + 1 < template.Length && template[i + 1] == '{') { sb.Append('{'); i++; continue; }
                var end = FindClosingBrace(template, i + 1);
                if (end < 0) { sb.Append(template[i..]); break; }
                var expr = template[(i + 1)..end];
                try
                {
                    var value = compiler.Compile(expr).Eval(ctx);
                    sb.Append(Values.ToDisplayString(value));
                }
                catch (MosExpressionException ex)
                {
                    sb.Append("{!").Append(ex.Message).Append('}');
                    onError?.Invoke($"In '{{{expr}}}': {ex.Message}");
                }
                i = end;
            }
            else if (ch == '}')
            {
                if (i + 1 < template.Length && template[i + 1] == '}') { sb.Append('}'); i++; continue; }
                sb.Append('}');
            }
            else sb.Append(ch);
        }
        return sb.ToString();
    }

    /// <summary>
    /// All {expression} placeholders in a text, in order (inner expression without braces).
    /// Same scanning rules as Interpolate, including "{{" escapes and quote awareness.
    /// </summary>
    public static List<(string Expression, string RawPlaceholder)> ExtractPlaceholders(string text)
    {
        var result = new List<(string, string)>();
        if (string.IsNullOrEmpty(text)) return result;
        for (var i = 0; i < text.Length; i++)
        {
            if (text[i] != '{') continue;
            if (i + 1 < text.Length && text[i + 1] == '{') { i++; continue; }
            var end = FindClosingBrace(text, i + 1);
            if (end < 0) break;
            result.Add((text[(i + 1)..end], text[i..(end + 1)]));
            i = end;
        }
        return result;
    }

    private static int FindClosingBrace(string s, int start)
    {
        // Track the expected closing quote so '}' inside string literals is skipped.
        // Word smart quotes pair “ with ” and ‘ with ’.
        char? closer = null;
        for (var i = start; i < s.Length; i++)
        {
            var c = s[i];
            if (closer is { } q)
            {
                if (c == q) closer = null;
            }
            else
            {
                switch (c)
                {
                    case '"' or '\'': closer = c; break;
                    case '“': closer = '”'; break;
                    case '‘': closer = '’'; break;
                    case '}': return i;
                }
            }
        }
        return -1;
    }
}
