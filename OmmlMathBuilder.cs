using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace AIGFH;

/// <summary>Small LaTeX-to-OMML renderer for the formula structures emitted by AI web clients.</summary>
internal static class OmmlMathBuilder
{
    internal const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    public static string BuildEquationXml(string latex)
    {
        var inner = BuildInner((latex ?? String.Empty).Trim());
        if (String.IsNullOrWhiteSpace(inner)) return null;
        return "<m:oMath xmlns:m=\"" + MathNamespace + "\">" + inner + "</m:oMath>";
    }

    private static string BuildInner(string latex)
    {
        var environment = Regex.Match(latex, @"^\\begin\s*\{(?<name>[^{}]+)\}(?:\s*\{[^{}]*\})?(?<body>.*?)\\end\s*\{\k<name>\}\s*$", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        if (!environment.Success) return new Parser(latex).Parse();

        var rawName = environment.Groups["name"].Value.Replace(" ", String.Empty);
        var name = rawName.ToLowerInvariant();
        var body = environment.Groups["body"].Value.Trim();
        var rows = SplitRows(body).Where(row => !String.IsNullOrWhiteSpace(row)).ToList();
        if (name.Contains("matrix") || name == "array") return BuildMatrix(rows, rawName);
        if (name == "cases" || name == "dcases") return BuildCases(rows);
        if (name.StartsWith("align", StringComparison.Ordinal) || name.StartsWith("flalign", StringComparison.Ordinal) ||
            name == "aligned" || name == "split" || name.StartsWith("gather", StringComparison.Ordinal) ||
            name.StartsWith("multline", StringComparison.Ordinal))
            return BuildEquationArray(rows);
        if (name == "equation" || name == "equation*" || name == "displaymath" || name == "math")
            return new Parser(body).Parse();
        return new Parser(body.Replace("&", String.Empty)).Parse();
    }

    private static string BuildEquationArray(IList<string> rows)
    {
        var xml = new StringBuilder("<m:eqArr><m:eqArrPr><m:baseJc m:val=\"center\"/></m:eqArrPr>");
        foreach (var raw in rows)
        {
            var row = NormalizeRow(raw);
            xml.Append("<m:e>").Append(new Parser(row).Parse()).Append("</m:e>");
        }
        return xml.Append("</m:eqArr>").ToString();
    }

    private static string BuildCases(IList<string> rows)
    {
        var matrix = new StringBuilder("<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val=\"2\"/><m:mcJc m:val=\"left\"/></m:mcPr></m:mc></m:mcs></m:mPr>");
        foreach (var raw in rows)
        {
            var columns = SplitColumns(raw);
            matrix.Append("<m:mr>");
            for (var column = 0; column < 2; column++)
            {
                var value = column < columns.Count ? columns[column].Trim() : String.Empty;
                matrix.Append("<m:e>").Append(new Parser(value).Parse()).Append("</m:e>");
            }
            matrix.Append("</m:mr>");
        }
        matrix.Append("</m:m>");
        return "<m:d><m:dPr><m:begChr m:val=\"{\"/><m:endChr m:val=\"\"/><m:grow m:val=\"1\"/></m:dPr><m:e>" + matrix + "</m:e></m:d>";
    }

    private static string BuildMatrix(IList<string> rows, string name)
    {
        var normalizedName = (name ?? String.Empty).ToLowerInvariant();
        var parsedRows = rows.Select(SplitColumns).ToList();
        var columns = parsedRows.Count == 0 ? 1 : Math.Max(1, parsedRows.Max(row => row.Count));
        var matrix = new StringBuilder("<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val=\"")
            .Append(columns).Append("\"/><m:mcJc m:val=\"center\"/></m:mcPr></m:mc></m:mcs></m:mPr>");
        foreach (var row in parsedRows)
        {
            matrix.Append("<m:mr>");
            for (var column = 0; column < columns; column++)
                matrix.Append("<m:e>").Append(new Parser(column < row.Count ? row[column].Trim() : String.Empty).Parse()).Append("</m:e>");
            matrix.Append("</m:mr>");
        }
        matrix.Append("</m:m>");
        if (normalizedName == "pmatrix") return Delimiter("(", ")", matrix.ToString());
        if (name == "Bmatrix") return Delimiter("{", "}", matrix.ToString());
        if (normalizedName == "bmatrix") return Delimiter("[", "]", matrix.ToString());
        if (name == "Vmatrix") return Delimiter("‖", "‖", matrix.ToString());
        if (normalizedName == "vmatrix") return Delimiter("|", "|", matrix.ToString());
        return matrix.ToString();
    }

    private static string Delimiter(string begin, string end, string content)
    {
        return "<m:d><m:dPr><m:begChr m:val=\"" + Xml(begin) + "\"/><m:endChr m:val=\"" + Xml(end) + "\"/><m:grow m:val=\"1\"/></m:dPr><m:e>" + content + "</m:e></m:d>";
    }

    private static List<string> SplitRows(string value)
    {
        var rows = new List<string>();
        var current = new StringBuilder();
        var depth = 0;
        for (var index = 0; index < value.Length; index++)
        {
            var character = value[index];
            if (character == '{') depth++;
            else if (character == '}') depth = Math.Max(0, depth - 1);
            if (character == '\\' && index + 1 < value.Length && value[index + 1] == '\\' && depth == 0)
            {
                rows.Add(current.ToString());
                current.Clear();
                index++;
                continue;
            }
            current.Append(character);
        }
        rows.Add(current.ToString());
        return rows;
    }

    private static List<string> SplitColumns(string value)
    {
        var columns = new List<string>();
        var current = new StringBuilder();
        var depth = 0;
        foreach (var character in value)
        {
            if (character == '{') depth++;
            else if (character == '}') depth = Math.Max(0, depth - 1);
            if (character == '&' && depth == 0)
            {
                columns.Add(current.ToString());
                current.Clear();
            }
            else current.Append(character);
        }
        columns.Add(current.ToString());
        return columns;
    }

    private static string NormalizeRow(string value)
    {
        return Regex.Replace(value ?? String.Empty, @"\\(?:tag|label)\s*\{.*?\}", String.Empty).Replace("&", String.Empty).Trim();
    }

    private static string Xml(string value) { return WebUtility.HtmlEncode(value ?? String.Empty); }

    private sealed class Parser
    {
        private readonly string _source;
        private int _index;

        public Parser(string source) { _source = source ?? String.Empty; }

        public string Parse(char? stop = null)
        {
            var xml = new StringBuilder();
            while (_index < _source.Length)
            {
                if (stop.HasValue && _source[_index] == stop.Value) { _index++; break; }
                var atom = ParseAtom();
                if (String.IsNullOrEmpty(atom)) continue;
                string sub = null, sup = null;
                while (_index < _source.Length && (_source[_index] == '_' || _source[_index] == '^'))
                {
                    var marker = _source[_index++];
                    var argument = ParseScriptArgument();
                    if (marker == '_') sub = argument; else sup = argument;
                }
                if (sub != null && sup != null)
                    atom = "<m:sSubSup><m:e>" + atom + "</m:e><m:sub>" + sub + "</m:sub><m:sup>" + sup + "</m:sup></m:sSubSup>";
                else if (sub != null)
                    atom = "<m:sSub><m:e>" + atom + "</m:e><m:sub>" + sub + "</m:sub></m:sSub>";
                else if (sup != null)
                    atom = "<m:sSup><m:e>" + atom + "</m:e><m:sup>" + sup + "</m:sup></m:sSup>";
                xml.Append(atom);
            }
            return xml.ToString();
        }

        private string ParseAtom()
        {
            if (_index >= _source.Length) return String.Empty;
            var character = _source[_index++];
            if (character == '{') return Parse('}');
            if (character == '}') return String.Empty;
            if (character == '\\') return ParseCommand();
            if (character == '~') return Run(" ", true);
            if (Char.IsWhiteSpace(character))
            {
                while (_index < _source.Length && Char.IsWhiteSpace(_source[_index])) _index++;
                return Run(" ", true);
            }
            return Run(character.ToString(), false);
        }

        private string ParseCommand()
        {
            if (_index >= _source.Length) return Run("\\", true);
            if (_source[_index] == '\\') { _index++; return Run(" ", true); }
            var start = _index;
            if (Char.IsLetter(_source[_index])) while (_index < _source.Length && Char.IsLetter(_source[_index])) _index++;
            else _index++;
            var command = _source.Substring(start, _index - start);

            if (command == "frac" || command == "dfrac" || command == "tfrac" || command == "cfrac")
            {
                var numerator = ParseRequiredGroup();
                var denominator = ParseRequiredGroup();
                return "<m:f><m:fPr><m:type m:val=\"bar\"/></m:fPr><m:num>" + numerator + "</m:num><m:den>" + denominator + "</m:den></m:f>";
            }
            if (command == "binom")
            {
                var top = ParseRequiredGroup();
                var bottom = ParseRequiredGroup();
                var stack = "<m:f><m:fPr><m:type m:val=\"noBar\"/></m:fPr><m:num>" + top + "</m:num><m:den>" + bottom + "</m:den></m:f>";
                return Delimiter("(", ")", stack);
            }
            if (command == "genfrac")
            {
                var begin = ReadRequiredGroupText();
                var end = ReadRequiredGroupText();
                var thickness = ReadRequiredGroupText();
                ReadRequiredGroupText(); // display style; Office chooses it automatically
                var top = ParseRequiredGroup();
                var bottom = ParseRequiredGroup();
                var type = String.Equals(thickness.Trim(), "0pt", StringComparison.OrdinalIgnoreCase) ? "noBar" : "bar";
                var fraction = "<m:f><m:fPr><m:type m:val=\"" + type + "\"/></m:fPr><m:num>" + top + "</m:num><m:den>" + bottom + "</m:den></m:f>";
                return !String.IsNullOrEmpty(begin) || !String.IsNullOrEmpty(end) ? Delimiter(begin, end, fraction) : fraction;
            }
            if (command == "prescript")
            {
                var sub = ParseRequiredGroup();
                var sup = ParseRequiredGroup();
                var body = ParseRequiredGroup();
                return "<m:sPre><m:sPrePr/><m:sub>" + sub + "</m:sub><m:sup>" + sup + "</m:sup><m:e>" + body + "</m:e></m:sPre>";
            }
            if (command == "overset" || command == "underset" || command == "stackrel")
            {
                var limit = ParseRequiredGroup();
                var body = ParseRequiredGroup();
                return command != "underset"
                    ? "<m:limUpp><m:e>" + body + "</m:e><m:lim>" + limit + "</m:lim></m:limUpp>"
                    : "<m:limLow><m:e>" + body + "</m:e><m:lim>" + limit + "</m:lim></m:limLow>";
            }
            if (command == "xrightarrow" || command == "xleftarrow")
            {
                var below = ParseOptionalBracket();
                var above = ParseRequiredGroup();
                var arrow = Run(command == "xrightarrow" ? "→" : "←", false);
                if (!String.IsNullOrWhiteSpace(below))
                    arrow = "<m:limLow><m:e>" + arrow + "</m:e><m:lim>" + below + "</m:lim></m:limLow>";
                return "<m:limUpp><m:e>" + arrow + "</m:e><m:lim>" + above + "</m:lim></m:limUpp>";
            }
            if (command == "boxed")
            {
                var body = ParseRequiredGroup();
                return "<m:borderBox><m:borderBoxPr/><m:e>" + body + "</m:e></m:borderBox>";
            }
            if (command == "cancel" || command == "bcancel" || command == "xcancel")
            {
                var body = ParseRequiredGroup();
                var strike = command == "cancel" ? "<m:strikeBLTR m:val=\"1\"/>" :
                    command == "bcancel" ? "<m:strikeTLBR m:val=\"1\"/>" :
                    "<m:strikeBLTR m:val=\"1\"/><m:strikeTLBR m:val=\"1\"/>";
                return "<m:borderBox><m:borderBoxPr>" + strike + "</m:borderBoxPr><m:e>" + body + "</m:e></m:borderBox>";
            }
            if (command == "overbrace" || command == "underbrace" || command == "overrightarrow" ||
                command == "overleftarrow" || command == "underrightarrow" || command == "underleftarrow" ||
                command == "overparen" || command == "wideparen")
            {
                var body = ParseRequiredGroup();
                var under = command.StartsWith("under", StringComparison.Ordinal);
                var position = under ? "bot" : "top";
                var character = command.Contains("right") ? "→" : command.Contains("left") ? "←" :
                    command.Contains("paren") ? "⏜" : under ? "⏟" : "⏞";
                return "<m:groupChr><m:groupChrPr><m:chr m:val=\"" + Xml(character) + "\"/><m:pos m:val=\"" + position + "\"/></m:groupChrPr><m:e>" + body + "</m:e></m:groupChr>";
            }
            if (command == "substack")
            {
                var raw = ReadRequiredGroupText();
                return BuildEquationArray(SplitRows(raw));
            }
            if (command == "pmod" || command == "pod")
            {
                var body = ParseRequiredGroup();
                return Delimiter("(", ")", Run("mod ", true) + body);
            }
            if (command == "mod" || command == "bmod") return Run(" mod ", true);
            if (command == "color" || command == "textcolor")
            {
                ReadRequiredGroupText();
                return ParseRequiredGroup();
            }
            if (command == "phantom" || command == "hphantom" || command == "vphantom") return ParseRequiredGroup();
            if (command == "sqrt")
            {
                var degree = ParseOptionalBracket();
                var body = ParseRequiredGroup();
                var property = degree == null ? "<m:radPr><m:degHide m:val=\"1\"/></m:radPr><m:deg/>" : "<m:radPr/><m:deg>" + degree + "</m:deg>";
                return "<m:rad>" + property + "<m:e>" + body + "</m:e></m:rad>";
            }
            if (command == "begin")
            {
                var rawName = ReadRequiredGroupText().Replace(" ", String.Empty);
                var name = rawName.ToLowerInvariant();
                if (name == "array")
                {
                    SkipWhiteSpace();
                    if (_index < _source.Length && _source[_index] == '{') ReadRequiredGroupText();
                }
                var endToken = "\\end{" + rawName + "}";
                var endAt = _source.IndexOf(endToken, _index, StringComparison.OrdinalIgnoreCase);
                if (endAt >= 0)
                {
                    var body = _source.Substring(_index, endAt - _index);
                    _index = endAt + endToken.Length;
                    var rows = SplitRows(body).Where(row => !String.IsNullOrWhiteSpace(row)).ToList();
                    if (name == "cases" || name == "dcases") return BuildCases(rows);
                    if (name.Contains("matrix") || name == "array") return BuildMatrix(rows, rawName);
                    if (name.StartsWith("align", StringComparison.Ordinal) || name.StartsWith("flalign", StringComparison.Ordinal) ||
                        name == "aligned" || name == "split" || name.StartsWith("gather", StringComparison.Ordinal) ||
                        name.StartsWith("multline", StringComparison.Ordinal))
                        return BuildEquationArray(rows);
                    return new Parser(body.Replace("&", String.Empty)).Parse();
                }
                return Run("\\begin{" + name + "}", true);
            }
            if (command == "text" || command == "textrm" || command == "mathrm" || command == "mathbf" || command == "mathit" || command == "mathsf" || command == "mathtt" || command == "operatorname")
            {
                var raw = ReadRequiredGroupText();
                if (command == "mathbf" && Regex.IsMatch(raw, @"[_^\\{}]")) return new Parser(raw).Parse();
                return Run(raw, command != "mathit", command == "mathbf");
            }
            if (command == "boldsymbol" || command == "bm" || command == "pmb" || command == "mathnormal")
            {
                var raw = ReadRequiredGroupText();
                return new Parser(raw).Parse();
            }
            if (command == "mathbb" || command == "mathcal" || command == "mathfrak")
            {
                var raw = ReadRequiredGroupText();
                if (command == "mathbb") raw = ToDoubleStruck(raw);
                return Run(raw, true);
            }
            if (command == "overline" || command == "bar" || command == "underline")
            {
                var body = ParseRequiredGroup();
                var position = command == "underline" ? "bot" : "top";
                return "<m:bar><m:barPr><m:pos m:val=\"" + position + "\"/></m:barPr><m:e>" + body + "</m:e></m:bar>";
            }
            if (command == "vec" || command == "hat" || command == "widehat" || command == "tilde" || command == "widetilde" || command == "breve" || command == "check" || command == "acute" || command == "grave" || command == "dot" || command == "ddot")
            {
                var accent = command == "vec" ? "→" : command == "dot" ? "̇" : command == "ddot" ? "̈" : command == "tilde" || command == "widetilde" ? "̃" : command == "breve" ? "̆" : command == "check" ? "̌" : command == "acute" ? "́" : command == "grave" ? "̀" : "̂";
                var body = ParseRequiredGroup();
                return "<m:acc><m:accPr><m:chr m:val=\"" + Xml(accent) + "\"/></m:accPr><m:e>" + body + "</m:e></m:acc>";
            }
            if (command == "left") return ParseStretchDelimiter();
            if (command == "right") { ReadDelimiter(); return String.Empty; }
            if (command == "middle") return Run(ReadDelimiter(), false);
            if (command == "displaystyle" || command == "textstyle" || command == "scriptstyle" || command == "scriptscriptstyle") return String.Empty;
            if (command == "quad") return Run(" ", true);
            if (command == "qquad") return Run("  ", true);
            if (command == " ") return Run(" ", true);
            if (command == "," || command == ";" || command == ":") return Run(" ", true);
            if (command == "!") return String.Empty;

            string symbol;
            if (Symbols.TryGetValue(command, out symbol)) return Run(symbol, Functions.Contains(command));
            return Run("\\" + command, true);
        }

        private string ParseStretchDelimiter()
        {
            var begin = ReadDelimiter();
            var endMarker = FindRightDelimiter(_index);
            if (endMarker < 0) return Run(begin, true);
            var bodyText = _source.Substring(_index, endMarker - _index);
            _index = endMarker + "\\right".Length;
            var end = ReadDelimiter();
            return Delimiter(begin, end, new Parser(bodyText).Parse());
        }

        private int FindRightDelimiter(int start)
        {
            var depth = 0;
            for (var index = start; index < _source.Length - 5; index++)
            {
                if (_source.Substring(index).StartsWith("\\left", StringComparison.Ordinal)) { depth++; index += 4; }
                else if (_source.Substring(index).StartsWith("\\right", StringComparison.Ordinal))
                {
                    if (depth == 0) return index;
                    depth--; index += 5;
                }
            }
            return -1;
        }

        private string ReadDelimiter()
        {
            SkipWhiteSpace();
            if (_index >= _source.Length) return String.Empty;
            if (_source[_index] != '\\')
            {
                var value = _source[_index++].ToString();
                return value == "." ? String.Empty : value;
            }
            _index++;
            if (_index < _source.Length && !Char.IsLetter(_source[_index]))
            {
                var literal = _source[_index++].ToString();
                return literal == "." ? String.Empty : literal;
            }
            var start = _index;
            while (_index < _source.Length && Char.IsLetter(_source[_index])) _index++;
            var command = _source.Substring(start, _index - start);
            string mapped;
            return Delimiters.TryGetValue(command, out mapped) ? mapped : command;
        }

        private string ParseScriptArgument()
        {
            SkipWhiteSpace();
            if (_index < _source.Length && _source[_index] == '{') { _index++; return Parse('}'); }
            return ParseAtom();
        }

        private string ParseRequiredGroup()
        {
            var text = ReadRequiredGroupText();
            return new Parser(text).Parse();
        }

        private string ReadRequiredGroupText()
        {
            SkipWhiteSpace();
            if (_index >= _source.Length || _source[_index] != '{')
            {
                if (_index >= _source.Length) return String.Empty;
                if (_source[_index] == '\\')
                {
                    var commandStart = _index++;
                    if (_index < _source.Length && Char.IsLetter(_source[_index]))
                        while (_index < _source.Length && Char.IsLetter(_source[_index])) _index++;
                    else if (_index < _source.Length) _index++;
                    return _source.Substring(commandStart, _index - commandStart);
                }
                return _source[_index++].ToString();
            }
            _index++;
            var start = _index;
            var depth = 1;
            while (_index < _source.Length && depth > 0)
            {
                if (_source[_index] == '{') depth++;
                else if (_source[_index] == '}') depth--;
                _index++;
            }
            var length = Math.Max(0, _index - start - 1);
            return _source.Substring(start, length);
        }

        private string ParseOptionalBracket()
        {
            SkipWhiteSpace();
            if (_index >= _source.Length || _source[_index] != '[') return null;
            _index++;
            var start = _index;
            while (_index < _source.Length && _source[_index] != ']') _index++;
            var text = _source.Substring(start, _index - start);
            if (_index < _source.Length) _index++;
            return new Parser(text).Parse();
        }

        private void SkipWhiteSpace() { while (_index < _source.Length && Char.IsWhiteSpace(_source[_index])) _index++; }

        private static string Run(string text, bool plain, bool bold = false)
        {
            var properties = plain || bold ? "<m:rPr>" + (plain ? "<m:sty m:val=\"p\"/>" : String.Empty) + (bold ? "<m:scr m:val=\"roman\"/><m:sty m:val=\"b\"/>" : String.Empty) + "</m:rPr>" : String.Empty;
            return "<m:r>" + properties + "<m:t xml:space=\"preserve\">" + Xml(text) + "</m:t></m:r>";
        }

        private static string ToDoubleStruck(string value)
        {
            if (String.IsNullOrEmpty(value)) return value;
            var builder = new StringBuilder();
            foreach (var character in value)
            {
                if (character >= 'A' && character <= 'Z')
                {
                    var special = character == 'C' ? "ℂ" : character == 'H' ? "ℍ" : character == 'N' ? "ℕ" :
                        character == 'P' ? "ℙ" : character == 'Q' ? "ℚ" : character == 'R' ? "ℝ" : character == 'Z' ? "ℤ" : null;
                    builder.Append(special ?? Char.ConvertFromUtf32(0x1D538 + character - 'A'));
                }
                else if (character >= 'a' && character <= 'z') builder.Append(Char.ConvertFromUtf32(0x1D552 + character - 'a'));
                else if (character >= '0' && character <= '9') builder.Append(Char.ConvertFromUtf32(0x1D7D8 + character - '0'));
                else builder.Append(character);
            }
            return builder.ToString();
        }
    }

    private static readonly HashSet<string> Functions = new HashSet<string>(StringComparer.Ordinal)
    { "sin", "cos", "tan", "cot", "sec", "csc", "sinh", "cosh", "tanh", "coth", "arcsin", "arccos", "arctan", "arcsec", "arccsc", "arccot", "lim", "limsup", "liminf", "log", "lg", "lb", "ln", "exp", "min", "max", "sup", "inf", "det", "gcd", "lcm", "rank", "tr", "diag", "ker", "dim", "span", "sgn", "arg", "Pr", "Var", "Cov" };

    private static readonly Dictionary<string, string> Delimiters = new Dictionary<string, string>(StringComparer.Ordinal)
    { { "{", "{" }, { "}", "}" }, { "lbrace", "{" }, { "rbrace", "}" }, { "langle", "〈" }, { "rangle", "〉" }, { "lvert", "|" }, { "rvert", "|" }, { "lVert", "‖" }, { "rVert", "‖" }, { "lfloor", "⌊" }, { "rfloor", "⌋" }, { "lceil", "⌈" }, { "rceil", "⌉" } };

    private static readonly Dictionary<string, string> Symbols = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        { "alpha", "α" }, { "beta", "β" }, { "gamma", "γ" }, { "delta", "δ" }, { "epsilon", "ε" }, { "varepsilon", "ε" },
        { "zeta", "ζ" }, { "eta", "η" }, { "theta", "θ" }, { "vartheta", "ϑ" }, { "iota", "ι" }, { "kappa", "κ" }, { "lambda", "λ" },
        { "mu", "μ" }, { "nu", "ν" }, { "xi", "ξ" }, { "pi", "π" }, { "varpi", "ϖ" }, { "rho", "ρ" }, { "varrho", "ϱ" },
        { "sigma", "σ" }, { "varsigma", "ς" }, { "tau", "τ" }, { "upsilon", "υ" }, { "phi", "φ" }, { "varphi", "ϕ" }, { "chi", "χ" }, { "psi", "ψ" }, { "omega", "ω" },
        { "Gamma", "Γ" }, { "Delta", "Δ" }, { "Theta", "Θ" }, { "Lambda", "Λ" }, { "Xi", "Ξ" }, { "Pi", "Π" }, { "Sigma", "Σ" }, { "Upsilon", "Υ" }, { "Phi", "Φ" }, { "Psi", "Ψ" }, { "Omega", "Ω" },
        { "times", "×" }, { "cdot", "·" }, { "div", "÷" }, { "circ", "∘" }, { "pm", "±" }, { "mp", "∓" }, { "leqslant", "≤" }, { "geqslant", "≥" }, { "leq", "≤" }, { "geq", "≥" }, { "le", "≤" }, { "ge", "≥" }, { "lt", "<" }, { "gt", ">" }, { "neq", "≠" }, { "ne", "≠" },
        { "approx", "≈" }, { "sim", "∼" }, { "simeq", "≃" }, { "equiv", "≡" }, { "infty", "∞" }, { "in", "∈" }, { "notin", "∉" }, { "subset", "⊂" }, { "subseteq", "⊆" }, { "subsetneq", "⊊" }, { "nsubseteq", "⊈" },
        { "propto", "∝" }, { "cong", "≅" }, { "asymp", "≍" }, { "doteq", "≐" }, { "prec", "≺" }, { "succ", "≻" }, { "preceq", "≼" }, { "succeq", "≽" }, { "ll", "≪" }, { "gg", "≫" }, { "lesssim", "≲" }, { "gtrsim", "≳" }, { "lessapprox", "⪅" }, { "gtrapprox", "⪆" },
        { "supset", "⊃" }, { "supseteq", "⊇" }, { "supsetneq", "⊋" }, { "nsupseteq", "⊉" }, { "cup", "∪" }, { "cap", "∩" }, { "forall", "∀" }, { "exists", "∃" }, { "partial", "∂" }, { "nabla", "∇" },
        { "oplus", "⊕" }, { "ominus", "⊖" }, { "otimes", "⊗" }, { "oslash", "⊘" }, { "odot", "⊙" }, { "uplus", "⊎" }, { "sqcup", "⊔" }, { "sqcap", "⊓" }, { "setminus", "∖" }, { "smallsetminus", "∖" }, { "neg", "¬" }, { "lnot", "¬" }, { "land", "∧" }, { "wedge", "∧" }, { "lor", "∨" }, { "vee", "∨" },
        { "top", "⊤" }, { "bot", "⊥" }, { "vdash", "⊢" }, { "dashv", "⊣" }, { "models", "⊨" }, { "ni", "∋" }, { "nexists", "∄" }, { "sqsubseteq", "⊑" }, { "sqsupseteq", "⊒" },
        { "complement", "∁" }, { "mid", "∣" }, { "nmid", "∤" }, { "colon", ":" }, { "ell", "ℓ" }, { "hbar", "ℏ" }, { "aleph", "ℵ" }, { "Re", "ℜ" }, { "Im", "ℑ" }, { "wp", "℘" },
        { "dots", "…" }, { "ldots", "…" }, { "cdots", "⋯" }, { "vdots", "⋮" }, { "ddots", "⋱" },
        { "rightarrow", "→" }, { "leftarrow", "←" }, { "leftrightarrow", "↔" }, { "uparrow", "↑" }, { "downarrow", "↓" }, { "updownarrow", "↕" }, { "hookrightarrow", "↪" }, { "hookleftarrow", "↩" }, { "rightharpoonup", "⇀" }, { "leftharpoonup", "↼" }, { "rightleftharpoons", "⇌" }, { "longrightarrow", "⟶" }, { "longleftarrow", "⟵" }, { "longleftrightarrow", "⟷" },
        { "Uparrow", "⇑" }, { "Downarrow", "⇓" }, { "Updownarrow", "⇕" },
        { "Rightarrow", "⇒" }, { "Leftarrow", "⇐" }, { "Leftrightarrow", "⇔" }, { "Longrightarrow", "⟹" }, { "Longleftarrow", "⟸" }, { "Longleftrightarrow", "⟺" },
        { "implies", "⇒" }, { "iff", "⇔" }, { "mapsto", "↦" }, { "to", "→" },
        { "sum", "∑" }, { "prod", "∏" }, { "coprod", "∐" }, { "bigcup", "⋃" }, { "bigcap", "⋂" }, { "bigvee", "⋁" }, { "bigwedge", "⋀" }, { "int", "∫" }, { "iint", "∬" }, { "iiint", "∭" }, { "oint", "∮" }, { "oiint", "∯" }, { "oiiint", "∰" },
        { "sin", "sin" }, { "cos", "cos" }, { "tan", "tan" }, { "cot", "cot" }, { "sec", "sec" }, { "csc", "csc" }, { "sinh", "sinh" }, { "cosh", "cosh" }, { "tanh", "tanh" }, { "coth", "coth" },
        { "arcsin", "arcsin" }, { "arccos", "arccos" }, { "arctan", "arctan" }, { "arcsec", "arcsec" }, { "arccsc", "arccsc" }, { "arccot", "arccot" }, { "lim", "lim" }, { "limsup", "lim sup" }, { "liminf", "lim inf" }, { "log", "log" }, { "lg", "lg" }, { "lb", "lb" }, { "ln", "ln" }, { "exp", "exp" }, { "min", "min" }, { "max", "max" }, { "sup", "sup" }, { "inf", "inf" }, { "det", "det" }, { "gcd", "gcd" }, { "lcm", "lcm" }, { "rank", "rank" }, { "tr", "tr" }, { "diag", "diag" }, { "ker", "ker" }, { "dim", "dim" }, { "span", "span" }, { "sgn", "sgn" }, { "arg", "arg" }, { "Pr", "Pr" }, { "Var", "Var" }, { "Cov", "Cov" },
        { "therefore", "∴" }, { "because", "∵" }, { "perp", "⊥" }, { "parallel", "∥" }, { "triangle", "△" }, { "angle", "∠" }, { "measuredangle", "∡" }, { "degree", "°" }, { "emptyset", "∅" }, { "varnothing", "∅" },
        { "%", "%" }, { "#", "#" }, { "_", "_" }, { "&", "&" }, { "{", "{" }, { "}", "}" }
    };
}
