using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace AIGFH;

internal static class OmmlToLatexConverter
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // These names are provided directly by standard LaTeX/amsmath.  Names that are
    // commonly used in Chinese school and university material but are not built in
    // are emitted through \operatorname so that the exported text remains valid TeX.
    private static readonly HashSet<string> StandardFunctions = new HashSet<string>(StringComparer.Ordinal)
    {
        "arccos", "arcsin", "arctan", "arg", "cos", "cosh", "cot", "coth", "csc",
        "deg", "det", "dim", "exp", "gcd", "hom", "inf", "ker", "lg", "lim",
        "liminf", "limsup", "ln", "log", "max", "min", "Pr", "sec", "sin", "sinh",
        "sup", "tan", "tanh"
    };

    private static readonly HashSet<string> NamedOperators = new HashSet<string>(StringComparer.Ordinal)
    {
        "arccot", "arcsec", "arccsc", "sech", "csch", "arsinh", "arcosh", "artanh",
        "lcm", "rank", "tr", "sgn", "diag", "span", "proj", "adj", "nullity",
        "grad", "curl", "erf", "Var", "Cov", "E"
    };

    internal static string Convert(string openXml)
    {
        if (String.IsNullOrWhiteSpace(openXml)) return String.Empty;
        try
        {
            var document = XDocument.Parse(openXml, LoadOptions.PreserveWhitespace);
            var maths = document.Descendants(M + "oMath").ToArray();
            return maths.Length == 0 ? String.Empty : Clean(String.Join(" ", maths.Select(ConvertChildren)));
        }
        catch { return String.Empty; }
    }

    internal static string ConvertLinear(string source)
    {
        var value = (source ?? String.Empty).Trim();
        if (value.Length == 0) return String.Empty;
        value = NormalizeLinearCharacters(value);
        value = ConvertLinearStructures(value);
        value = ConvertLinearFractions(value);
        value = ReplaceBalancedFunction(value, '√', body =>
        {
            var separator = FindTopLevel(body, '&');
            return separator < 0
                ? "\\sqrt{" + ConvertLinear(body) + "}"
                : "\\sqrt[" + ConvertLinear(body.Substring(0, separator)) + "]{" +
                  ConvertLinear(body.Substring(separator + 1)) + "}";
        });
        value = Regex.Replace(value, @"([_^])\(([^()]*)\)", match => match.Groups[1].Value + "{" + ConvertLinear(match.Groups[2].Value) + "}");
        value = Regex.Replace(value,
            @"(?<!\\)\b(arccos|arcsin|arctan|arccot|arcsec|arccsc|arsinh|arcosh|artanh|sinh|cosh|tanh|coth|sech|csch|limsup|liminf|sin|cos|tan|cot|sec|csc|ln|lg|log|exp|lim|min|max|sup|inf|det|gcd|lcm|rank|tr|ker|dim|arg|hom|deg|sgn|diag|span|proj|adj|nullity|grad|curl|erf|Var|Cov|Pr)\b",
            match => FunctionCommand(match.Groups[1].Value));
        return Clean(MapSymbols(value));
    }

    private static string ConvertNode(XElement node)
    {
        var name = node.Name.LocalName;
        switch (name)
        {
            case "r": return ConvertRun(node);
            case "t": return MapSymbols(node.Value);
            case "f":
            {
                var numerator = ConvertPart(node, "num");
                var denominator = ConvertPart(node, "den");
                var fractionType = node.Descendants(M + "type").Select(item => AttributeValue(item, "val")).FirstOrDefault();
                return fractionType == "noBar" ? "\\binom{" + numerator + "}{" + denominator + "}" :
                    "\\frac{" + numerator + "}{" + denominator + "}";
            }
            case "rad":
            {
                var body = ConvertPart(node, "e");
                var degree = ConvertPart(node, "deg");
                var hidden = node.Descendants(M + "degHide").Any(item => AttributeValue(item, "val") == "1");
                return hidden || String.IsNullOrWhiteSpace(degree) ? "\\sqrt{" + body + "}" : "\\sqrt[" + degree + "]{" + body + "}";
            }
            case "sSub": return ConvertPart(node, "e") + "_{" + ConvertPart(node, "sub") + "}";
            case "sSup": return ConvertPart(node, "e") + "^{" + ConvertPart(node, "sup") + "}";
            case "sSubSup": return ConvertPart(node, "e") + "_{" + ConvertPart(node, "sub") + "}^{" + ConvertPart(node, "sup") + "}";
            case "sPre": return "_{" + ConvertPart(node, "sub") + "}^{" + ConvertPart(node, "sup") + "}" + ConvertPart(node, "e");
            case "nary":
            {
                var symbol = node.Descendants(M + "chr").Select(item => AttributeValue(item, "val")).FirstOrDefault();
                var command = NaryCommand(symbol);
                var lower = ConvertPart(node, "sub");
                var upper = ConvertPart(node, "sup");
                if (!String.IsNullOrWhiteSpace(lower)) command += "_{" + lower + "}";
                if (!String.IsNullOrWhiteSpace(upper)) command += "^{" + upper + "}";
                return command + " " + ConvertPart(node, "e");
            }
            case "func":
            {
                var function = ConvertPart(node, "fName").Trim();
                if (Regex.IsMatch(function, @"^[A-Za-z]+$")) function = FunctionCommand(function);
                return function + " " + ConvertPart(node, "e");
            }
            case "d":
            {
                var begin = node.Descendants(M + "begChr").Select(item => AttributeValue(item, "val")).FirstOrDefault() ?? "(";
                var end = node.Descendants(M + "endChr").Select(item => AttributeValue(item, "val")).FirstOrDefault() ?? ")";
                begin = Delimiter(begin); end = Delimiter(end);
                var parts = node.Elements(M + "e").Select(ConvertChildren).ToArray();
                var separator = node.Descendants(M + "sepChr").Select(item => AttributeValue(item, "val")).FirstOrDefault();
                var middle = String.IsNullOrEmpty(separator) ? "|" : Delimiter(separator);
                if (parts.Length == 1 && parts[0].StartsWith("\\binom", StringComparison.Ordinal) &&
                    begin == "(" && end == ")") return parts[0];
                return "\\left" + begin + String.Join("\\middle" + middle, parts) + "\\right" + end;
            }
            case "limLow": return ConvertPart(node, "e") + "_{" + ConvertPart(node, "lim") + "}";
            case "limUpp": return ConvertPart(node, "e") + "^{" + ConvertPart(node, "lim") + "}";
            case "acc":
            {
                var mark = node.Descendants(M + "chr").Select(item => AttributeValue(item, "val")).FirstOrDefault();
                var command = AccentCommand(mark);
                return "\\" + command + "{" + ConvertPart(node, "e") + "}";
            }
            case "bar":
            {
                var position = node.Descendants(M + "pos").Select(item => AttributeValue(item, "val")).FirstOrDefault();
                return "\\" + (position == "bot" ? "underline" : "overline") + "{" + ConvertPart(node, "e") + "}";
            }
            case "groupChr":
            {
                var position = node.Descendants(M + "pos").Select(item => AttributeValue(item, "val")).FirstOrDefault();
                var mark = node.Descendants(M + "chr").Select(item => AttributeValue(item, "val")).FirstOrDefault();
                var command = GroupCharacterCommand(mark, position);
                return "\\" + command + "{" + ConvertPart(node, "e") + "}";
            }
            case "borderBox":
            {
                var topLeftBottomRight = node.Descendants(M + "strikeTLBR").Any(item => AttributeValue(item, "val") != "0");
                var bottomLeftTopRight = node.Descendants(M + "strikeBLTR").Any(item => AttributeValue(item, "val") != "0");
                var command = topLeftBottomRight && bottomLeftTopRight ? "xcancel" : topLeftBottomRight ? "bcancel" : bottomLeftTopRight ? "cancel" : "boxed";
                return "\\" + command + "{" + ConvertPart(node, "e") + "}";
            }
            case "m":
            {
                var rows = node.Elements(M + "mr").Select(row => String.Join(" & ", row.Elements(M + "e").Select(ConvertChildren)));
                return "\\begin{matrix}" + String.Join(" \\\\ ", rows) + "\\end{matrix}";
            }
            case "eqArr":
            {
                var rows = node.Elements(M + "e").Select(ConvertChildren);
                return "\\begin{aligned}" + String.Join(" \\\\ ", rows) + "\\end{aligned}";
            }
            case "phant": return "\\phantom{" + ConvertPart(node, "e") + "}";
            default:
                if (name.EndsWith("Pr", StringComparison.Ordinal) || name == "ctrlPr" || name == "degHide" || name == "subHide" || name == "supHide" || name == "chr" || name == "pos") return String.Empty;
                return ConvertChildren(node);
        }
    }

    private static string ConvertChildren(XElement node)
    {
        return String.Concat(node.Elements().Select(ConvertNode));
    }

    private static string ConvertRun(XElement node)
    {
        var value = String.Concat(node.Descendants().Where(item => item.Name == M + "t" || item.Name == W + "t").Select(item => item.Value));
        var plain = node.Descendants(M + "sty").Any(item => AttributeValue(item, "val") == "p");
        if (!plain) return MapSymbols(value);
        if (String.IsNullOrWhiteSpace(value)) return value;
        var trimmed = value.Trim();
        if (StandardFunctions.Contains(trimmed) || NamedOperators.Contains(trimmed))
            return value.Replace(trimmed, FunctionCommand(trimmed) + " ");
        if (Regex.IsMatch(trimmed, @"^[A-Za-z]+$")) return "\\mathrm{" + trimmed + "}";
        if (Regex.IsMatch(trimmed, @"[\u3400-\u9fff]")) return "\\text{" + EscapeText(trimmed) + "}";
        return MapSymbols(value);
    }

    private static string EscapeText(string value)
    {
        return (value ?? String.Empty).Replace("\\", "\\textbackslash{}")
            .Replace("{", "\\{").Replace("}", "\\}").Replace("%", "\\%").Replace("#", "\\#")
            .Replace("&", "\\&").Replace("_", "\\_");
    }

    private static string NormalizeLinearCharacters(string value)
    {
        // FormKC helpfully normalizes full-width input, but it also destroys the
        // semantic distinction carried by mathematical alphanumerics and scripts.
        // Convert those two families first, then normalize the remaining text.
        value = ConvertDoubleStruckSymbols(value);
        value = ConvertUnicodeScripts(value);
        var protectedSymbols = new Dictionary<string, string>
        {
            ["ϑ"]="\uE300", ["ϰ"]="\uE301", ["ϖ"]="\uE302", ["ϱ"]="\uE303", ["ϕ"]="\uE304",
            ["∬"]="\uE305", ["∭"]="\uE306", ["∯"]="\uE307", ["∰"]="\uE308",
            ["ℏ"]="\uE309", ["ℵ"]="\uE30A", ["ℜ"]="\uE30B", ["ℑ"]="\uE30C", ["ℓ"]="\uE30D",
            ["ℬ"]="\uE30E", ["ℰ"]="\uE30F", ["ℱ"]="\uE310", ["ℋ"]="\uE311", ["ℐ"]="\uE312",
            ["ℒ"]="\uE313", ["ℳ"]="\uE314", ["ℛ"]="\uE315", ["…"]="\uE316", ["″"]="\uE317"
        };
        foreach (var pair in protectedSymbols) value = value.Replace(pair.Key, pair.Value);
        try { value = value.Normalize(NormalizationForm.FormKC); } catch { }
        foreach (var pair in protectedSymbols) value = value.Replace(pair.Value, pair.Key);
        return value;
    }

    private static string FunctionCommand(string name)
    {
        if (StandardFunctions.Contains(name)) return "\\" + name;
        return NamedOperators.Contains(name) ? "\\operatorname{" + name + "}" : "\\mathrm{" + name + "}";
    }

    private static string ConvertDoubleStruckSymbols(string value)
    {
        if (String.IsNullOrEmpty(value)) return value;
        var result = new StringBuilder(value.Length);
        for (var index = 0; index < value.Length; index++)
        {
            var codePoint = Char.ConvertToUtf32(value, index);
            var sourceLength = Char.IsSurrogatePair(value, index) ? 2 : 1;
            char letter = '\0';

            // Mathematical Double-Struck capitals keep code-point slots for the
            // seven historical BMP characters, so the alphabetic offset is stable.
            if (codePoint >= 0x1D538 && codePoint <= 0x1D551)
                letter = (char)('A' + codePoint - 0x1D538);
            else if (codePoint >= 0x1D552 && codePoint <= 0x1D56B)
                letter = (char)('a' + codePoint - 0x1D552);
            else if (codePoint >= 0x1D7D8 && codePoint <= 0x1D7E1)
                letter = (char)('0' + codePoint - 0x1D7D8);
            else
            {
                switch (codePoint)
                {
                    case 0x2102: letter = 'C'; break; // ℂ
                    case 0x210D: letter = 'H'; break; // ℍ
                    case 0x2115: letter = 'N'; break; // ℕ
                    case 0x2119: letter = 'P'; break; // ℙ
                    case 0x211A: letter = 'Q'; break; // ℚ
                    case 0x211D: letter = 'R'; break; // ℝ
                    case 0x2124: letter = 'Z'; break; // ℤ
                }
            }

            if (letter != '\0') result.Append("\\mathbb{").Append(letter).Append('}');
            else result.Append(value, index, sourceLength);
            if (sourceLength == 2) index++;
        }
        return result.ToString();
    }

    private static string ConvertUnicodeScripts(string value)
    {
        if (String.IsNullOrEmpty(value)) return value;
        var result = new StringBuilder(value.Length);
        for (var index = 0; index < value.Length;)
        {
            string mapped;
            if (TryMapSuperscript(value[index], out mapped))
            {
                result.Append("^{");
                do
                {
                    result.Append(mapped);
                    index++;
                }
                while (index < value.Length && TryMapSuperscript(value[index], out mapped));
                result.Append('}');
                continue;
            }
            if (TryMapSubscript(value[index], out mapped))
            {
                result.Append("_{");
                do
                {
                    result.Append(mapped);
                    index++;
                }
                while (index < value.Length && TryMapSubscript(value[index], out mapped));
                result.Append('}');
                continue;
            }
            result.Append(value[index++]);
        }
        return result.ToString();
    }

    private static bool TryMapSuperscript(char value, out string mapped)
    {
        switch (value)
        {
            case '⁰': mapped = "0"; return true; case '¹': mapped = "1"; return true;
            case '²': mapped = "2"; return true; case '³': mapped = "3"; return true;
            case '⁴': mapped = "4"; return true; case '⁵': mapped = "5"; return true;
            case '⁶': mapped = "6"; return true; case '⁷': mapped = "7"; return true;
            case '⁸': mapped = "8"; return true; case '⁹': mapped = "9"; return true;
            case '⁺': mapped = "+"; return true; case '⁻': mapped = "-"; return true;
            case '⁼': mapped = "="; return true; case '⁽': mapped = "("; return true;
            case '⁾': mapped = ")"; return true;
            case 'ᴬ': mapped = "A"; return true; case 'ᴮ': mapped = "B"; return true;
            case 'ᴰ': mapped = "D"; return true; case 'ᴱ': mapped = "E"; return true;
            case 'ᴳ': mapped = "G"; return true; case 'ᴴ': mapped = "H"; return true;
            case 'ᴵ': mapped = "I"; return true; case 'ᴶ': mapped = "J"; return true;
            case 'ᴷ': mapped = "K"; return true; case 'ᴸ': mapped = "L"; return true;
            case 'ᴹ': mapped = "M"; return true; case 'ᴺ': mapped = "N"; return true;
            case 'ᴼ': mapped = "O"; return true; case 'ᴾ': mapped = "P"; return true;
            case 'ᴿ': mapped = "R"; return true; case 'ᵀ': mapped = "T"; return true;
            case 'ᵁ': mapped = "U"; return true; case 'ⱽ': mapped = "V"; return true;
            case 'ᵂ': mapped = "W"; return true;
            case 'ᵃ': mapped = "a"; return true; case 'ᵇ': mapped = "b"; return true;
            case 'ᶜ': mapped = "c"; return true; case 'ᵈ': mapped = "d"; return true;
            case 'ᵉ': mapped = "e"; return true; case 'ᶠ': mapped = "f"; return true;
            case 'ᵍ': mapped = "g"; return true; case 'ʰ': mapped = "h"; return true;
            case 'ⁱ': mapped = "i"; return true; case 'ʲ': mapped = "j"; return true;
            case 'ᵏ': mapped = "k"; return true; case 'ˡ': mapped = "l"; return true;
            case 'ᵐ': mapped = "m"; return true; case 'ⁿ': mapped = "n"; return true;
            case 'ᵒ': mapped = "o"; return true; case 'ᵖ': mapped = "p"; return true;
            case 'ʳ': mapped = "r"; return true; case 'ˢ': mapped = "s"; return true;
            case 'ᵗ': mapped = "t"; return true; case 'ᵘ': mapped = "u"; return true;
            case 'ᵛ': mapped = "v"; return true; case 'ʷ': mapped = "w"; return true;
            case 'ˣ': mapped = "x"; return true; case 'ʸ': mapped = "y"; return true;
            case 'ᶻ': mapped = "z"; return true;
            case 'ᵅ': mapped = "\\alpha "; return true; case 'ᵝ': mapped = "\\beta "; return true;
            case 'ᵞ': mapped = "\\gamma "; return true; case 'ᵟ': mapped = "\\delta "; return true;
            case 'ᵋ': mapped = "\\epsilon "; return true; case 'ᶿ': mapped = "\\theta "; return true;
            case 'ᶥ': mapped = "\\iota "; return true; case 'ᶲ': mapped = "\\phi "; return true;
            case 'ᵡ': mapped = "\\chi "; return true;
            default: mapped = null; return false;
        }
    }

    private static bool TryMapSubscript(char value, out string mapped)
    {
        switch (value)
        {
            case '₀': mapped = "0"; return true; case '₁': mapped = "1"; return true;
            case '₂': mapped = "2"; return true; case '₃': mapped = "3"; return true;
            case '₄': mapped = "4"; return true; case '₅': mapped = "5"; return true;
            case '₆': mapped = "6"; return true; case '₇': mapped = "7"; return true;
            case '₈': mapped = "8"; return true; case '₉': mapped = "9"; return true;
            case '₊': mapped = "+"; return true; case '₋': mapped = "-"; return true;
            case '₌': mapped = "="; return true; case '₍': mapped = "("; return true;
            case '₎': mapped = ")"; return true;
            case 'ₐ': mapped = "a"; return true; case 'ₑ': mapped = "e"; return true;
            case 'ₕ': mapped = "h"; return true; case 'ᵢ': mapped = "i"; return true;
            case 'ⱼ': mapped = "j"; return true; case 'ₖ': mapped = "k"; return true;
            case 'ₗ': mapped = "l"; return true; case 'ₘ': mapped = "m"; return true;
            case 'ₙ': mapped = "n"; return true; case 'ₒ': mapped = "o"; return true;
            case 'ₚ': mapped = "p"; return true; case 'ᵣ': mapped = "r"; return true;
            case 'ₛ': mapped = "s"; return true; case 'ₜ': mapped = "t"; return true;
            case 'ᵤ': mapped = "u"; return true; case 'ᵥ': mapped = "v"; return true;
            case 'ₓ': mapped = "x"; return true;
            case 'ᵦ': mapped = "\\beta "; return true; case 'ᵧ': mapped = "\\gamma "; return true;
            case 'ᵨ': mapped = "\\rho "; return true; case 'ᵩ': mapped = "\\phi "; return true;
            case 'ᵪ': mapped = "\\chi "; return true;
            default: mapped = null; return false;
        }
    }

    private static string ConvertPart(XElement node, string localName)
    {
        var part = node.Elements(M + localName).FirstOrDefault();
        return part == null ? String.Empty : ConvertChildren(part);
    }

    private static string AttributeValue(XElement node, string localName)
    {
        var attribute = node.Attributes().FirstOrDefault(item => item.Name.LocalName == localName);
        return attribute == null ? null : attribute.Value;
    }

    private static string NaryCommand(string value)
    {
        switch (value)
        {
            case "∑": return "\\sum";
            case "∏": return "\\prod";
            case "∬": return "\\iint";
            case "∭": return "\\iiint";
            case "∮": return "\\oint";
            case "∯": return "\\oiint";
            case "∰": return "\\oiiint";
            case "∐": return "\\coprod";
            case "⋃": return "\\bigcup";
            case "⋂": return "\\bigcap";
            default: return String.IsNullOrWhiteSpace(value) ? "\\int" : MapSymbols(value).Trim();
        }
    }

    private static string AccentCommand(string value)
    {
        switch (value)
        {
            case "→": return "vec";
            case "~": case "̃": return "tilde";
            case "̇": return "dot";
            case "̈": return "ddot";
            case "¯": case "̅": return "bar";
            case "ˇ": case "̌": return "check";
            case "´": case "́": return "acute";
            case "`": case "̀": return "grave";
            case "˘": case "̆": return "breve";
            default: return "hat";
        }
    }

    private static string GroupCharacterCommand(string value, string position)
    {
        if (value == "⏞" || value == "︷") return "overbrace";
        if (value == "⏟" || value == "︸") return "underbrace";
        if (value == "→") return position == "bot" ? "underrightarrow" : "overrightarrow";
        if (value == "←") return position == "bot" ? "underleftarrow" : "overleftarrow";
        return position == "bot" ? "underbrace" : "overbrace";
    }

    private static string Delimiter(string value)
    {
        if (String.IsNullOrEmpty(value) || value == ".") return ".";
        switch (value)
        {
            case "{": return "\\{";
            case "}": return "\\}";
            case "〈": case "⟨": return "\\langle";
            case "〉": case "⟩": return "\\rangle";
            case "‖": return "\\Vert";
            case "|": return "\\vert";
            case "⌊": return "\\lfloor";
            case "⌋": return "\\rfloor";
            case "⌈": return "\\lceil";
            case "⌉": return "\\rceil";
            default: return value;
        }
    }

    private static string MapSymbols(string value)
    {
        value = ConvertDoubleStruckSymbols(value ?? String.Empty);
        value = ConvertUnicodeScripts(value);
        var map = new Dictionary<string, string>
        {
            ["α"]="\\alpha", ["β"]="\\beta", ["γ"]="\\gamma", ["δ"]="\\delta", ["ε"]="\\epsilon", ["ϵ"]="\\varepsilon", ["ζ"]="\\zeta", ["η"]="\\eta", ["θ"]="\\theta", ["ϑ"]="\\vartheta", ["ι"]="\\iota", ["κ"]="\\kappa", ["ϰ"]="\\varkappa", ["λ"]="\\lambda", ["μ"]="\\mu", ["ν"]="\\nu", ["ξ"]="\\xi", ["π"]="\\pi", ["ϖ"]="\\varpi", ["ρ"]="\\rho", ["ϱ"]="\\varrho", ["σ"]="\\sigma", ["ς"]="\\varsigma", ["τ"]="\\tau", ["φ"]="\\varphi", ["ϕ"]="\\phi", ["χ"]="\\chi", ["ψ"]="\\psi", ["ω"]="\\omega",
            ["Γ"]="\\Gamma", ["Δ"]="\\Delta", ["Θ"]="\\Theta", ["Λ"]="\\Lambda", ["Ξ"]="\\Xi", ["Π"]="\\Pi", ["Σ"]="\\Sigma", ["Φ"]="\\Phi", ["Ψ"]="\\Psi", ["Ω"]="\\Omega",
            ["≤"]="\\leq", ["≦"]="\\leq", ["≥"]="\\geq", ["≧"]="\\geq", ["≠"]="\\neq", ["≈"]="\\approx", ["≉"]="\\not\\approx", ["≃"]="\\simeq", ["≅"]="\\cong", ["≍"]="\\asymp", ["≡"]="\\equiv", ["≢"]="\\not\\equiv", ["≐"]="\\doteq", ["≜"]="\\triangleq", ["≔"]="\\mathrel{:=}", ["∼"]="\\sim", ["∝"]="\\propto", ["≪"]="\\ll", ["≫"]="\\gg", ["≲"]="\\lesssim", ["≳"]="\\gtrsim", ["≺"]="\\prec", ["≻"]="\\succ", ["≼"]="\\preceq", ["≽"]="\\succeq", ["≮"]="\\nless", ["≯"]="\\ngtr", ["≰"]="\\nleq", ["≱"]="\\ngeq", ["≶"]="\\lessgtr", ["≷"]="\\gtrless",
            ["∞"]="\\infty", ["±"]="\\pm", ["∓"]="\\mp", ["−"]="-", ["×"]="\\times", ["·"]="\\cdot", ["⋅"]="\\cdot", ["∙"]="\\bullet", ["÷"]="\\div", ["∕"]="/", ["⁄"]="/", ["∗"]="\\ast", ["⋆"]="\\star", ["∘"]="\\circ", ["≀"]="\\wr",
            ["∈"]="\\in", ["∉"]="\\notin", ["∋"]="\\ni", ["∌"]="\\not\\ni", ["⊂"]="\\subset", ["⊆"]="\\subseteq", ["⊊"]="\\subsetneq", ["⊄"]="\\nsubset", ["⊈"]="\\nsubseteq", ["⊃"]="\\supset", ["⊇"]="\\supseteq", ["⊋"]="\\supsetneq", ["⊅"]="\\nsupset", ["⊉"]="\\nsupseteq", ["∪"]="\\cup", ["∩"]="\\cap", ["⊎"]="\\uplus", ["⊔"]="\\sqcup", ["⊓"]="\\sqcap", ["∖"]="\\setminus", ["∁"]="\\complement", ["∅"]="\\varnothing",
            ["¬"]="\\neg", ["∧"]="\\land", ["∨"]="\\lor", ["⊤"]="\\top", ["⊥"]="\\perp", ["⊢"]="\\vdash", ["⊣"]="\\dashv", ["⊨"]="\\models", ["⊬"]="\\nvdash", ["⊭"]="\\nvDash", ["∀"]="\\forall", ["∃"]="\\exists", ["∄"]="\\nexists", ["∴"]="\\therefore", ["∵"]="\\because",
            ["∣"]="\\mid", ["∤"]="\\nmid", ["∥"]="\\parallel", ["∦"]="\\nparallel", ["∂"]="\\partial", ["∇"]="\\nabla", ["√"]="\\sqrt", ["∠"]="\\angle", ["∡"]="\\measuredangle", ["△"]="\\triangle",
            ["⊕"]="\\oplus", ["⊖"]="\\ominus", ["⊗"]="\\otimes", ["⊘"]="\\oslash", ["⊙"]="\\odot", ["⊛"]="\\circledast", ["⊚"]="\\circledcirc", ["⊝"]="\\circleddash", ["⊞"]="\\boxplus", ["⊟"]="\\boxminus", ["⊠"]="\\boxtimes", ["⊡"]="\\boxdot", ["⊲"]="\\triangleleft", ["⊳"]="\\triangleright", ["⊴"]="\\trianglelefteq", ["⊵"]="\\trianglerighteq",
            ["∑"]="\\sum", ["∏"]="\\prod", ["∐"]="\\coprod", ["∫"]="\\int", ["∬"]="\\iint", ["∭"]="\\iiint", ["∮"]="\\oint", ["∯"]="\\oiint", ["∰"]="\\oiiint", ["⋃"]="\\bigcup", ["⋂"]="\\bigcap", ["⨆"]="\\bigsqcup", ["⨁"]="\\bigoplus", ["⨂"]="\\bigotimes", ["⨀"]="\\bigodot", ["⨄"]="\\biguplus",
            ["→"]="\\to", ["←"]="\\leftarrow", ["↔"]="\\leftrightarrow", ["⇒"]="\\Rightarrow", ["⇐"]="\\Leftarrow", ["⇔"]="\\Leftrightarrow", ["⟶"]="\\longrightarrow", ["⟵"]="\\longleftarrow", ["⟷"]="\\longleftrightarrow", ["⟹"]="\\Longrightarrow", ["⟸"]="\\Longleftarrow", ["⟺"]="\\Longleftrightarrow", ["↦"]="\\mapsto", ["↪"]="\\hookrightarrow", ["↩"]="\\hookleftarrow", ["↠"]="\\twoheadrightarrow", ["↞"]="\\twoheadleftarrow", ["↝"]="\\leadsto", ["⇀"]="\\rightharpoonup", ["⇁"]="\\rightharpoondown", ["↼"]="\\leftharpoonup", ["↽"]="\\leftharpoondown", ["⇌"]="\\rightleftharpoons", ["⇋"]="\\leftrightharpoons", ["↑"]="\\uparrow", ["↓"]="\\downarrow", ["↕"]="\\updownarrow", ["⇑"]="\\Uparrow", ["⇓"]="\\Downarrow", ["⇕"]="\\Updownarrow", ["↗"]="\\nearrow", ["↘"]="\\searrow", ["↙"]="\\swarrow", ["↖"]="\\nwarrow",
            ["ℏ"]="\\hbar", ["ℵ"]="\\aleph", ["ℜ"]="\\Re", ["ℑ"]="\\Im", ["℘"]="\\wp", ["ℓ"]="\\ell", ["ℬ"]="\\mathcal{B}", ["ℰ"]="\\mathcal{E}", ["ℱ"]="\\mathcal{F}", ["ℋ"]="\\mathcal{H}", ["ℐ"]="\\mathcal{I}", ["ℒ"]="\\mathcal{L}", ["ℳ"]="\\mathcal{M}", ["ℛ"]="\\mathcal{R}",
            ["⌊"]="\\lfloor", ["⌋"]="\\rfloor", ["⌈"]="\\lceil", ["⌉"]="\\rceil", ["‖"]="\\Vert", ["…"]="\\ldots", ["⋯"]="\\cdots", ["⋮"]="\\vdots", ["⋱"]="\\ddots", ["⋄"]="\\diamond", ["□"]="\\square", ["∎"]="\\blacksquare", ["′"]="^{\\prime}", ["″"]="^{\\prime\\prime}", ["°"]="^{\\circ}"
        };
        foreach (var pair in map) value = value.Replace(pair.Key, pair.Value + " ");
        return value;
    }

    private static string ConvertLinearStructures(string source)
    {
        var value = source.Replace("├", "\\left").Replace("┤", "\\right");
        value = ReplaceBalancedStructure(value, "\\matrix", body => ConvertGrid(body, "matrix"));
        value = ReplaceBalancedStructure(value, "\\eqarray", body => ConvertGrid(body, "aligned"));
        value = ReplaceBalancedStructure(value, "\\cases", body => ConvertGrid(body, "cases"));
        value = ReplaceBalancedStructure(value, "■", body => ConvertGrid(body, "matrix"));
        value = ReplaceBalancedStructure(value, "█", body => ConvertGrid(body, "aligned"));
        return value;
    }

    private static string ReplaceBalancedStructure(string source, string marker, Func<string, string> replacement)
    {
        var value = source;
        var searchFrom = 0;
        while (searchFrom < value.Length)
        {
            var index = value.IndexOf(marker, searchFrom, StringComparison.Ordinal);
            if (index < 0) break;
            var open = index + marker.Length;
            while (open < value.Length && Char.IsWhiteSpace(value[open])) open++;
            if (open >= value.Length || value[open] != '(')
            {
                searchFrom = open;
                continue;
            }
            var close = MatchingClose(value, open);
            if (close < 0) break;
            var converted = replacement(value.Substring(open + 1, close - open - 1));
            value = value.Substring(0, index) + converted + value.Substring(close + 1);
            searchFrom = index + converted.Length;
        }
        return value;
    }

    private static string ConvertGrid(string body, string environment)
    {
        var rows = SplitTopLevel(body, '@')
            .Select(row => String.Join(" & ", SplitTopLevel(row, '&').Select(ConvertLinear)))
            .ToArray();
        return "\\begin{" + environment + "}" + String.Join(" \\\\ ", rows) + "\\end{" + environment + "}";
    }

    private static IEnumerable<string> SplitTopLevel(string value, char separator)
    {
        var start = 0;
        var round = 0;
        var square = 0;
        var brace = 0;
        for (var index = 0; index < value.Length; index++)
        {
            switch (value[index])
            {
                case '(': round++; break;
                case ')': if (round > 0) round--; break;
                case '[': square++; break;
                case ']': if (square > 0) square--; break;
                case '{': brace++; break;
                case '}': if (brace > 0) brace--; break;
            }
            if (value[index] != separator || round != 0 || square != 0 || brace != 0) continue;
            yield return value.Substring(start, index - start);
            start = index + 1;
        }
        yield return value.Substring(start);
    }

    private static int FindTopLevel(string value, char marker)
    {
        var round = 0;
        var square = 0;
        var brace = 0;
        for (var index = 0; index < value.Length; index++)
        {
            switch (value[index])
            {
                case '(': round++; break;
                case ')': if (round > 0) round--; break;
                case '[': square++; break;
                case ']': if (square > 0) square--; break;
                case '{': brace++; break;
                case '}': if (brace > 0) brace--; break;
            }
            if (value[index] == marker && round == 0 && square == 0 && brace == 0) return index;
        }
        return -1;
    }

    private static string ConvertLinearFractions(string source)
    {
        var value = source;
        for (var pass = 0; pass < 32; pass++)
        {
            var slash = FindFractionSlash(value);
            if (slash < 0) break;
            var leftStart = OperandStart(value, slash - 1);
            var rightEnd = OperandEnd(value, slash + 1);
            if (leftStart >= slash || rightEnd <= slash + 1) break;
            var numerator = StripOuter(value.Substring(leftStart, slash - leftStart).Trim());
            var denominator = StripOuter(value.Substring(slash + 1, rightEnd - slash - 1).Trim());
            if (numerator.Length == 0 || denominator.Length == 0) break;
            var replacement = "\\frac{" + ConvertLinear(numerator) + "}{" + ConvertLinear(denominator) + "}";
            value = value.Substring(0, leftStart) + replacement + value.Substring(rightEnd);
        }
        return value;
    }

    private static int FindFractionSlash(string value)
    {
        for (var index = value.Length - 1; index >= 0; index--)
            if (value[index] == '/' && (index == 0 || value[index - 1] != '\\')) return index;
        return -1;
    }

    private static int OperandStart(string value, int index)
    {
        while (index >= 0 && Char.IsWhiteSpace(value[index])) index--;
        if (index < 0) return 0;
        if (value[index] == ')') return MatchingOpen(value, index);
        while (index >= 0 && IsLinearOperandCharacter(value[index])) index--;
        return index + 1;
    }

    private static int OperandEnd(string value, int index)
    {
        while (index < value.Length && Char.IsWhiteSpace(value[index])) index++;
        if (index >= value.Length) return value.Length;
        if (value[index] == '(')
        {
            var close = MatchingClose(value, index);
            return close < 0 ? value.Length : close + 1;
        }
        while (index < value.Length && IsLinearOperandCharacter(value[index])) index++;
        return index;
    }

    private static bool IsLinearOperandCharacter(char value)
    {
        return Char.IsLetterOrDigit(value) && value < 0x370 ||
               value >= 0x370 && value <= 0x3ff ||
               value == '_' || value == '^' || value == '\'' || value == '∞';
    }

    private static int MatchingOpen(string value, int close)
    {
        var depth = 0;
        for (var index = close; index >= 0; index--)
        {
            if (value[index] == ')') depth++;
            else if (value[index] == '(' && --depth == 0) return index;
        }
        return close;
    }

    private static int MatchingClose(string value, int open)
    {
        var depth = 0;
        for (var index = open; index < value.Length; index++)
        {
            if (value[index] == '(') depth++;
            else if (value[index] == ')' && --depth == 0) return index;
        }
        return -1;
    }

    private static string StripOuter(string value)
    {
        return value.Length >= 2 && value[0] == '(' && MatchingClose(value, 0) == value.Length - 1 ? value.Substring(1, value.Length - 2) : value;
    }

    private static string ReplaceBalancedFunction(string source, char marker, Func<string, string> replacement)
    {
        var value = source;
        for (var index = value.IndexOf(marker); index >= 0; index = value.IndexOf(marker, index + 1))
        {
            var open = index + 1;
            while (open < value.Length && Char.IsWhiteSpace(value[open])) open++;
            if (open >= value.Length || value[open] != '(') continue;
            var close = MatchingClose(value, open);
            if (close < 0) continue;
            var result = replacement(value.Substring(open + 1, close - open - 1));
            value = value.Substring(0, index) + result + value.Substring(close + 1);
            index += result.Length - 1;
        }
        return value;
    }

    private static string Clean(string value)
    {
        value = Regex.Replace(value ?? String.Empty, @"[ \t]+", " ").Trim();
        value = Regex.Replace(value, @"\s+([_^{},)])", "$1");
        value = Regex.Replace(value, @"([({])\s+", "$1");
        return value;
    }
}
