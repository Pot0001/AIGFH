using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace AIGFH;

/// <summary>
/// Converts the Markdown emitted by the common AI web clients into a self-contained
/// HTML fragment which Word can import as native rich document content.
/// </summary>
internal static class AiWebFormatter
{
    private static readonly Regex FenceStart = new Regex(@"^\s*(?<fence>`{3,}|~{3,})\s*(?<lang>[^\s`]*)\s*$", RegexOptions.Compiled);
    private static readonly Regex Heading = new Regex(@"^\s{0,3}(?<marks>#{1,6})[ \t]+(?<text>.*?)[ \t]*#*\s*$", RegexOptions.Compiled);
    private static readonly Regex ListItem = new Regex(@"^(?<indent>[ \t]*)(?<marker>[-+*]|\d+[.)])[ \t]+(?<text>.+)$", RegexOptions.Compiled);
    private static readonly Regex Quote = new Regex(@"^\s{0,3}>[ \t]?(?<text>.*)$", RegexOptions.Compiled);
    private static readonly Regex Setext = new Regex(@"^\s*(?<mark>=+|-+)\s*$", RegexOptions.Compiled);
    private static readonly Regex FormulaBlockStart = new Regex(@"^\s*(?<start>\$\$|\\\[)\s*(?<tail>.*)$", RegexOptions.Compiled);

    public static string ToHtml(string markdown, bool preserveHeadings)
    {
        var value = (markdown ?? String.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Replace("&nbsp;", " ")
            .Replace("<br />", "\n")
            .Replace("<br/>", "\n")
            .Replace("<br>", "\n")
            .Trim('\n', '\a');

        var lines = value.Split(new[] { '\n' }, StringSplitOptions.None).ToList();
        var body = RenderBlocks(lines, preserveHeadings);
        return "<!DOCTYPE html><html><head><meta charset=\"utf-8\">" +
               "<style>" +
               "body{font-family:'Microsoft YaHei','DengXian','Aptos',sans-serif;font-size:11pt;color:#1f2937;line-height:1.55;}" +
               "p{margin:0 0 8pt 0;}" +
               "h1,h2,h3,h4,h5,h6{font-family:'Microsoft YaHei','DengXian','Aptos Display',sans-serif;color:#111827;font-weight:700;page-break-after:avoid;}" +
               "h1{font-size:20pt;margin:18pt 0 10pt 0;}h2{font-size:16pt;margin:16pt 0 8pt 0;}h3{font-size:14pt;margin:14pt 0 7pt 0;}" +
               "h4{font-size:12pt;margin:12pt 0 6pt 0;}h5,h6{font-size:11pt;margin:10pt 0 5pt 0;}" +
               "a{color:#0969da;text-decoration:underline;}blockquote{color:#57606a;border-left:3pt solid #d0d7de;margin:8pt 0 10pt 0;padding:2pt 0 2pt 10pt;}" +
               "pre{font-family:Consolas,'Courier New',monospace;font-size:9.5pt;line-height:1.35;background:#f6f8fa;border:1pt solid #d0d7de;padding:9pt;margin:6pt 0 10pt 0;white-space:pre-wrap;}" +
               "code{font-family:Consolas,'Courier New',monospace;font-size:9.5pt;background:#eff1f3;color:#b42318;padding:1pt 3pt;}" +
               "table{border-collapse:collapse;width:100%;margin:8pt 0 12pt 0;}th,td{border:1pt solid #d0d7de;padding:6pt 7pt;vertical-align:top;}th{background:#f3f4f6;font-weight:700;color:#111827;}" +
               "</style></head><body>" + body + "</body></html>";
    }

    private static string RenderBlocks(IList<string> lines, bool preserveHeadings)
    {
        var html = new StringBuilder();
        var index = 0;
        while (index < lines.Count)
        {
            if (String.IsNullOrWhiteSpace(lines[index])) { index++; continue; }

            var fence = FenceStart.Match(lines[index]);
            if (fence.Success)
            {
                RenderFence(lines, ref index, fence, html);
                continue;
            }

            var formula = FormulaBlockStart.Match(lines[index]);
            if (formula.Success)
            {
                RenderFormulaBlock(lines, ref index, formula, html);
                continue;
            }

            if (IsTableAt(lines, index))
            {
                RenderTable(lines, ref index, html);
                continue;
            }

            var heading = Heading.Match(lines[index]);
            if (heading.Success)
            {
                var level = Math.Min(6, heading.Groups["marks"].Value.Length);
                if (preserveHeadings)
                    html.Append("<h").Append(level).Append('>').Append(RenderInline(heading.Groups["text"].Value)).Append("</h").Append(level).Append('>');
                else
                    AppendParagraph(html, heading.Groups["text"].Value);
                index++;
                continue;
            }

            if (index + 1 < lines.Count && !String.IsNullOrWhiteSpace(lines[index]) && Setext.IsMatch(lines[index + 1]))
            {
                var mark = Setext.Match(lines[index + 1]).Groups["mark"].Value;
                var level = mark[0] == '=' ? 1 : 2;
                if (preserveHeadings)
                    html.Append("<h").Append(level).Append('>').Append(RenderInline(lines[index].Trim())).Append("</h").Append(level).Append('>');
                else
                    AppendParagraph(html, lines[index].Trim());
                index += 2;
                continue;
            }

            if (IsHorizontalRule(lines[index]))
            {
                html.Append("<hr style=\"border:0;border-top:1pt solid #d0d7de;margin:12pt 0;\">");
                index++;
                continue;
            }

            if (Quote.IsMatch(lines[index]))
            {
                var quoted = new List<string>();
                while (index < lines.Count)
                {
                    var quote = Quote.Match(lines[index]);
                    if (!quote.Success) break;
                    quoted.Add(quote.Groups["text"].Value);
                    index++;
                }
                html.Append("<blockquote style=\"border-left:3pt solid #d0d7de;color:#57606a;margin:8pt 0 10pt 0;padding-left:10pt;\">")
                    .Append(RenderBlocks(quoted, preserveHeadings)).Append("</blockquote>");
                continue;
            }

            if (ListItem.IsMatch(lines[index]))
            {
                RenderList(lines, ref index, html);
                continue;
            }

            var paragraph = new List<string>();
            while (index < lines.Count && !String.IsNullOrWhiteSpace(lines[index]) && !IsBlockStart(lines, index))
            {
                paragraph.Add(lines[index].Trim());
                index++;
            }
            if (paragraph.Count == 0)
            {
                paragraph.Add(lines[index].Trim());
                index++;
            }
            var joined = JoinParagraphLines(paragraph);
            AppendParagraph(html, joined, true);
        }
        return html.ToString();
    }

    private static void RenderFence(IList<string> lines, ref int index, Match fence, StringBuilder html)
    {
        var marker = fence.Groups["fence"].Value;
        var language = fence.Groups["lang"].Value.Trim();
        var content = new List<string>();
        index++;
        while (index < lines.Count && !Regex.IsMatch(lines[index], @"^\s*" + Regex.Escape(marker.Substring(0, 1)) + "{" + marker.Length + @",}\s*$"))
        {
            content.Add(lines[index]);
            index++;
        }
        if (index < lines.Count) index++;

        if (language.Equals("latex", StringComparison.OrdinalIgnoreCase) ||
            language.Equals("tex", StringComparison.OrdinalIgnoreCase) ||
            language.Equals("math", StringComparison.OrdinalIgnoreCase))
        {
            html.Append("<p style=\"text-align:center;margin:10pt 0 12pt 0;\">")
                .Append(WebUtility.HtmlEncode("$$" + String.Join("\n", content) + "$$"))
                .Append("</p>");
            return;
        }

        if (!String.IsNullOrWhiteSpace(language))
            html.Append("<p style=\"font-family:Consolas,'Courier New',monospace;font-size:8.5pt;color:#57606a;background:#eef1f4;margin:7pt 0 0 0;padding:4pt 8pt;\">")
                .Append(WebUtility.HtmlEncode(language)).Append("</p>");
        html.Append("<pre style=\"font-family:Consolas,'Courier New',monospace;font-size:9.5pt;line-height:1.35;background:#f6f8fa;border:1pt solid #d0d7de;padding:9pt;margin:")
            .Append(String.IsNullOrWhiteSpace(language) ? "6pt" : "0")
            .Append(" 0 10pt 0;white-space:pre-wrap;\">")
            .Append(WebUtility.HtmlEncode(String.Join("\n", content))).Append("</pre>");
    }

    private static void RenderFormulaBlock(IList<string> lines, ref int index, Match start, StringBuilder html)
    {
        var opening = start.Groups["start"].Value;
        var closing = opening == "$$" ? "$$" : "\\]";
        var parts = new List<string>();
        var tail = start.Groups["tail"].Value;

        if (tail.Contains(closing))
        {
            parts.Add(tail.Substring(0, tail.IndexOf(closing, StringComparison.Ordinal)));
            index++;
        }
        else
        {
            if (tail.Length > 0) parts.Add(tail);
            index++;
            while (index < lines.Count)
            {
                var line = lines[index];
                var closeAt = line.IndexOf(closing, StringComparison.Ordinal);
                if (closeAt >= 0)
                {
                    if (closeAt > 0) parts.Add(line.Substring(0, closeAt));
                    index++;
                    break;
                }
                parts.Add(line);
                index++;
            }
        }
        var source = opening + String.Join("\n", parts) + closing;
        html.Append("<p style=\"text-align:center;margin:10pt 0 12pt 0;\">")
            .Append(WebUtility.HtmlEncode(source)).Append("</p>");
    }

    private static void RenderList(IList<string> lines, ref int index, StringBuilder html)
    {
        while (index < lines.Count)
        {
            var item = ListItem.Match(lines[index]);
            if (!item.Success) break;
            var indent = item.Groups["indent"].Value.Replace("\t", "    ").Length;
            var level = Math.Min(6, indent / 2);
            var marker = item.Groups["marker"].Value;
            var text = item.Groups["text"].Value.Trim();
            var task = Regex.Match(text, @"^\[(?<state>[ xX])\]\s*(?<body>.*)$");
            string bullet;
            if (task.Success)
            {
                bullet = task.Groups["state"].Value == " " ? "☐" : "☑";
                text = task.Groups["body"].Value;
            }
            else if (Char.IsDigit(marker[0])) bullet = Regex.Replace(marker, @"[)]$", ".");
            else bullet = new[] { "•", "◦", "▪" }[level % 3];

            var left = 18 + level * 18;
            html.Append("<p style=\"margin:2pt 0 4pt ").Append(left).Append("pt;text-indent:-13pt;line-height:1.5;\"><span style=\"font-weight:600;\">")
                .Append(WebUtility.HtmlEncode(bullet)).Append("</span>&nbsp;")
                .Append(RenderInline(text)).Append("</p>");
            index++;
        }
    }

    private static void RenderTable(IList<string> lines, ref int index, StringBuilder html)
    {
        var headers = SplitTableRow(lines[index]);
        var separators = SplitTableRow(lines[index + 1]);
        var aligns = separators.Select(GetTableAlignment).ToList();
        var rows = new List<List<string>>();
        index += 2;
        while (index < lines.Count && !String.IsNullOrWhiteSpace(lines[index]) && LooksLikeTableRow(lines[index]))
        {
            rows.Add(SplitTableRow(lines[index]));
            index++;
        }

        var columns = Math.Max(headers.Count, Math.Max(aligns.Count, rows.Count == 0 ? 0 : rows.Max(row => row.Count)));
        html.Append("<table border=\"1\" cellspacing=\"0\" cellpadding=\"0\" width=\"100%\" style=\"border-collapse:collapse;width:100%;margin:8pt 0 12pt 0;\"><thead><tr>");
        for (var column = 0; column < columns; column++)
        {
            var align = column < aligns.Count ? aligns[column] : "left";
            html.Append("<th bgcolor=\"#f3f4f6\" style=\"border:1pt solid #d0d7de;padding:6pt 7pt;text-align:").Append(align).Append(";font-weight:700;\">")
                .Append(RenderInline(column < headers.Count ? headers[column].Trim() : String.Empty)).Append("</th>");
        }
        html.Append("</tr></thead><tbody>");
        foreach (var row in rows)
        {
            html.Append("<tr>");
            for (var column = 0; column < columns; column++)
            {
                var align = column < aligns.Count ? aligns[column] : "left";
                html.Append("<td style=\"border:1pt solid #d0d7de;padding:6pt 7pt;text-align:").Append(align).Append(";vertical-align:top;\">")
                    .Append(RenderInline(column < row.Count ? row[column].Trim() : String.Empty)).Append("</td>");
            }
            html.Append("</tr>");
        }
        html.Append("</tbody></table>");
    }

    private static bool IsTableAt(IList<string> lines, int index)
    {
        if (index + 1 >= lines.Count || !LooksLikeTableRow(lines[index])) return false;
        var cells = SplitTableRow(lines[index + 1]);
        return cells.Count > 0 && cells.All(cell => Regex.IsMatch(cell.Trim(), @"^:?-{3,}:?$"));
    }

    private static bool LooksLikeTableRow(string line)
    {
        if (String.IsNullOrWhiteSpace(line)) return false;
        var trimmed = line.Trim();
        return trimmed.Contains("|") && !(trimmed.StartsWith("```", StringComparison.Ordinal) || trimmed.StartsWith("~~~", StringComparison.Ordinal));
    }

    private static List<string> SplitTableRow(string line)
    {
        var value = line.Trim();
        if (value.StartsWith("|", StringComparison.Ordinal)) value = value.Substring(1);
        if (value.EndsWith("|", StringComparison.Ordinal) && !value.EndsWith("\\|", StringComparison.Ordinal)) value = value.Substring(0, value.Length - 1);
        var result = new List<string>();
        var cell = new StringBuilder();
        var escaped = false;
        var inCode = false;
        foreach (var character in value)
        {
            if (escaped)
            {
                if (character != '|') cell.Append('\\');
                cell.Append(character);
                escaped = false;
                continue;
            }
            if (character == '\\') { escaped = true; continue; }
            if (character == '`') inCode = !inCode;
            if (character == '|' && !inCode)
            {
                result.Add(cell.ToString());
                cell.Clear();
            }
            else cell.Append(character);
        }
        if (escaped) cell.Append('\\');
        result.Add(cell.ToString());
        return result;
    }

    private static string GetTableAlignment(string separator)
    {
        var value = separator.Trim();
        if (value.StartsWith(":", StringComparison.Ordinal) && value.EndsWith(":", StringComparison.Ordinal)) return "center";
        if (value.EndsWith(":", StringComparison.Ordinal)) return "right";
        return "left";
    }

    private static bool IsBlockStart(IList<string> lines, int index)
    {
        var line = lines[index];
        if (FenceStart.IsMatch(line) || FormulaBlockStart.IsMatch(line) || Heading.IsMatch(line) || Quote.IsMatch(line) || ListItem.IsMatch(line) || IsHorizontalRule(line)) return true;
        if (IsTableAt(lines, index)) return true;
        return index + 1 < lines.Count && !String.IsNullOrWhiteSpace(line) && Setext.IsMatch(lines[index + 1]);
    }

    private static bool IsHorizontalRule(string line)
    {
        var compact = Regex.Replace(line ?? String.Empty, @"\s", String.Empty);
        if (compact.Length < 3) return false;
        return compact.All(character => character == '-') || compact.All(character => character == '*') || compact.All(character => character == '_');
    }

    private static string JoinParagraphLines(IList<string> lines)
    {
        var value = new StringBuilder();
        for (var index = 0; index < lines.Count; index++)
        {
            var line = lines[index];
            var hardBreak = line.EndsWith("  ", StringComparison.Ordinal) || line.EndsWith("\\", StringComparison.Ordinal);
            line = hardBreak ? line.TrimEnd(' ', '\\') : line;
            if (index > 0) value.Append(hardBreak ? "<AI_BR>" : " ");
            value.Append(line);
        }
        return value.ToString();
    }

    private static void AppendParagraph(StringBuilder html, string value, bool containsBreakTokens = false)
    {
        var rendered = RenderInline(value);
        if (containsBreakTokens) rendered = rendered.Replace("&lt;AI_BR&gt;", "<br>").Replace("<AI_BR>", "<br>");
        html.Append("<p style=\"margin:0 0 8pt 0;line-height:1.55;\">").Append(rendered).Append("</p>");
    }

    private static string RenderInline(string source)
    {
        if (String.IsNullOrEmpty(source)) return String.Empty;
        var protectedValues = new List<string>();
        Func<string, string> protect = value =>
        {
            var token = "\uE000" + protectedValues.Count.ToString() + "\uE001";
            protectedValues.Add(value);
            return token;
        };

        var value = source;
        value = Regex.Replace(value, @"(?<!\\)(\$\$.*?\$\$|\$(?:\\.|[^$\r\n])+?\$|\\\(.*?\\\)|\\\[.*?\\\])", match => protect(WebUtility.HtmlEncode(match.Value)));
        value = Regex.Replace(value, @"(?<!`)`([^`\r\n]+)`(?!`)", match => protect("<code style=\"font-family:Consolas,'Courier New',monospace;font-size:9.5pt;background:#eff1f3;color:#b42318;padding:1pt 3pt;\">" + WebUtility.HtmlEncode(match.Groups[1].Value) + "</code>"));
        value = Regex.Replace(value, @"!\[(?<alt>[^\]]*)\]\((?<url>[^\s)]+)(?:\s+[""'](?<title>.*?)[""'])?\)", match =>
            protect("<a href=\"" + HtmlAttribute(match.Groups["url"].Value) + "\" style=\"color:#0969da;text-decoration:underline;\">[图片] " + WebUtility.HtmlEncode(match.Groups["alt"].Value) + "</a>"));
        value = Regex.Replace(value, @"\[(?<label>[^\]]+)\]\((?<url>[^\s)]+)(?:\s+[""'](?<title>.*?)[""'])?\)", match =>
            protect("<a href=\"" + HtmlAttribute(match.Groups["url"].Value) + "\" style=\"color:#0969da;text-decoration:underline;\">" + RenderInline(match.Groups["label"].Value) + "</a>"));
        value = Regex.Replace(value, @"<(?<url>https?://[^>]+)>", match => protect("<a href=\"" + HtmlAttribute(match.Groups["url"].Value) + "\" style=\"color:#0969da;text-decoration:underline;\">" + WebUtility.HtmlEncode(match.Groups["url"].Value) + "</a>"));
        value = Regex.Replace(value, @"\\([\\`*_{}\[\]()#+.!>|~-])", match => protect(WebUtility.HtmlEncode(match.Groups[1].Value)));

        value = WebUtility.HtmlEncode(value);
        value = Regex.Replace(value, @"\*\*(.+?)\*\*|__(.+?)__", match => "<strong>" + (match.Groups[1].Success ? match.Groups[1].Value : match.Groups[2].Value) + "</strong>");
        value = Regex.Replace(value, @"~~(.+?)~~", "<del>$1</del>");
        value = Regex.Replace(value, @"(?<!\*)\*([^*\r\n]+)\*(?!\*)|(?<![\w])_([^_\r\n]+)_(?![\w])", match => "<em>" + (match.Groups[1].Success ? match.Groups[1].Value : match.Groups[2].Value) + "</em>");

        for (var index = protectedValues.Count - 1; index >= 0; index--)
            value = value.Replace("\uE000" + index.ToString() + "\uE001", protectedValues[index]);
        return value;
    }

    private static string HtmlAttribute(string value)
    {
        return WebUtility.HtmlEncode((value ?? String.Empty).Trim());
    }
}
