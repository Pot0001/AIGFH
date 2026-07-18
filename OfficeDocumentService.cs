using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace AIGFH;

public sealed class OfficeDocumentService
{
    private readonly object _application;

    public OfficeDocumentService(object application)
    {
        _application = application;
    }

    public string ActiveDocumentName
    {
        get
        {
            try { return ((dynamic)_application).ActiveDocument.Name as string ?? "当前文档"; }
            catch { return "未打开文档"; }
        }
    }

    public int ReplaceAll(string findText, string replaceText)
    {
        return ReplaceAll(findText, replaceText, false);
    }

    public int ReplaceAll(string findText, string replaceText, bool useWildcards)
    {
        if (String.IsNullOrEmpty(findText)) return 0;
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic range = document.Content;
        dynamic find = range.Find;
        find.ClearFormatting();
        find.Replacement.ClearFormatting();
        find.Text = findText;
        find.Replacement.Text = replaceText;
        find.Forward = true;
        find.Wrap = 1;
        find.Format = false;
        find.MatchCase = false;
        find.MatchWholeWord = false;
        find.MatchWildcards = useWildcards;
        if (ExecuteReplaceAll(find)) return 1;
        // WPS 的 Find.Execute 在部分版本中既不接受命名参数，也不返回可靠的
        // 替换结果。逐个修改命中的小范围，避免重写全文后破坏已有公式和对象。
        if (useWildcards) return 0;
        return ReplaceLiteralRanges(document, findText, replaceText) > 0 ? 1 : 0;
    }

    private static int ReplaceLiteralRanges(dynamic document, string findText, string replaceText)
    {
        var content = document.Content;
        var source = content.Text as string ?? String.Empty;
        var positions = new List<int>();
        for (var index = source.IndexOf(findText, StringComparison.Ordinal);
             index >= 0;
             index = source.IndexOf(findText, index + findText.Length, StringComparison.Ordinal))
            positions.Add(index);
        var start = (int)content.Start;
        var changed = 0;
        foreach (var position in positions.OrderByDescending(value => value))
        {
            try
            {
                document.Range(start + position, start + position + findText.Length).Text = replaceText ?? String.Empty;
                changed++;
            }
            catch { }
        }
        return changed;
    }

    public int BatchApplyRules(System.Collections.Generic.IEnumerable<string> sourceFiles, System.Collections.Generic.IEnumerable<TextReplacement> rules)
    {
        var applied = 0;
        dynamic documents = ((dynamic)_application).Documents;
        foreach (var sourcePath in sourceFiles)
        {
            var outputDirectory = Path.Combine(Path.GetDirectoryName(sourcePath), "OfflineOutput");
            Directory.CreateDirectory(outputDirectory);
            var targetPath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(sourcePath) + ".offline.docx");
            File.Copy(sourcePath, targetPath, true);
            dynamic document = null;
            try
            {
                document = documents.Open(targetPath);
                foreach (var rule in rules) ApplyRuleToDocument(document, rule);
                CollapseRepeatedSpacesInDocument(document);
                document.Save();
                applied++;
            }
            finally
            {
                if (document != null) try { document.Close(); } catch { }
            }
        }
        return applied;
    }

    public int BatchApplyFormulaColor(System.Collections.Generic.IEnumerable<string> sourceFiles, string hex, int delayMilliseconds)
    {
        var color = ParseWordColor(hex);
        var applied = 0;
        dynamic documents = ((dynamic)_application).Documents;
        foreach (var sourcePath in sourceFiles)
        {
            var outputDirectory = Path.Combine(Path.GetDirectoryName(sourcePath), "OfflineOutput"); Directory.CreateDirectory(outputDirectory);
            var targetPath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(sourcePath) + ".formula-color.docx"); File.Copy(sourcePath, targetPath, true);
            dynamic document = null;
            try { document = documents.Open(targetPath); ApplyFormulaColor(document, color); document.Save(); applied++; }
            finally { if (document != null) try { document.Close(); } catch { } }
            if (delayMilliseconds > 0) Thread.Sleep(Math.Min(delayMilliseconds, 10000));
        }
        return applied;
    }

    public int GeneratePowerPoint()
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        var pageCount = (int)document.ComputeStatistics(2); // wdStatisticPages
        if (pageCount < 1) throw new InvalidOperationException("当前文档没有可生成幻灯片的页面。");
        var powerPointType = Type.GetTypeFromProgID("PowerPoint.Application");
        if (powerPointType == null) powerPointType = Type.GetTypeFromProgID("KWPP.Application");
        if (powerPointType == null) throw new InvalidOperationException("未检测到 PowerPoint 或 WPS 演示组件。");
        dynamic powerPoint = Activator.CreateInstance(powerPointType);
        powerPoint.Visible = true;
        dynamic presentation = powerPoint.Presentations.Add();
        var slideWidth = (float)presentation.PageSetup.SlideWidth;
        var slideHeight = (float)presentation.PageSetup.SlideHeight;
        dynamic selection = ((dynamic)_application).Selection;
        var originalStart = (int)selection.Start;
        var originalEnd = (int)selection.End;
        var created = 0;
        try
        {
            for (var page = 1; page <= pageCount; page++)
            {
                dynamic startRange = document.GoTo(1, 1, page); // wdGoToPage, wdGoToAbsolute
                var start = (int)startRange.Start;
                var end = (int)document.Content.End - 1;
                if (page < pageCount)
                {
                    dynamic nextRange = document.GoTo(1, 1, page + 1);
                    end = Math.Max(start, (int)nextRange.Start - 1);
                }
                dynamic pageRange = document.Range(start, end);
                pageRange.Select();
                selection.CopyAsPicture();
                Thread.Sleep(100);

                dynamic slide = presentation.Slides.Add(++created, 12); // ppLayoutBlank
                dynamic shapeRange;
                try { shapeRange = slide.Shapes.PasteSpecial(2); } // ppPasteEnhancedMetafile
                catch { shapeRange = slide.Shapes.Paste(); }
                dynamic picture = shapeRange[1];
                picture.LockAspectRatio = -1;
                var scale = Math.Min(slideWidth / (float)picture.Width, slideHeight / (float)picture.Height);
                picture.Width = (float)picture.Width * scale;
                picture.Left = (slideWidth - (float)picture.Width) / 2f;
                picture.Top = (slideHeight - (float)picture.Height) / 2f;
            }
        }
        finally
        {
            try { document.Range(originalStart, originalEnd).Select(); } catch { }
        }
        return created;
    }

    private static bool IsHeading(string value) { return value.StartsWith("#", StringComparison.Ordinal) || value.Length <= 28 && (value.EndsWith("：", StringComparison.Ordinal) || value.EndsWith(":", StringComparison.Ordinal)); }

    private static void ApplyFormulaColor(dynamic document, int color)
    {
        for (var index = SafeCollectionCount(document.OMaths); index >= 1; index--)
            try { document.OMaths[index].Range.Font.Color = color; } catch { }
        for (var index = SafeCollectionCount(document.InlineShapes); index >= 1; index--)
        {
            try
            {
                dynamic shape = document.InlineShapes[index];
                var id = Convert.ToString(shape.OLEFormat.ProgID);
                if (IsEquationOleProgramId(id)) shape.Range.Font.Color = color;
            }
            catch { }
        }
    }

    private static int ParseWordColor(string hex)
    {
        hex = (hex ?? String.Empty).Trim().TrimStart('#');
        if (!Regex.IsMatch(hex, "^[0-9a-fA-F]{6}$")) throw new ArgumentException("颜色必须是 6 位十六进制 RGB 值。");
        var red = Convert.ToInt32(hex.Substring(0, 2), 16); var green = Convert.ToInt32(hex.Substring(2, 2), 16); var blue = Convert.ToInt32(hex.Substring(4, 2), 16);
        return (blue << 16) | (green << 8) | red;
    }

    public string GetSelectedText()
    {
        try
        {
            dynamic selection = ((dynamic)_application).Selection;
            var text = selection.Range.Text as string;
            if (!String.IsNullOrWhiteSpace(text)) return text.Trim();
        }
        catch { }
        try { return (((dynamic)_application).ActiveDocument.Content.Text as string ?? String.Empty).Trim(); }
        catch { return String.Empty; }
    }

    public void PasteClipboard()
    {
        dynamic selection = ((dynamic)_application).Selection;
        selection.Paste();
    }

    public int NormalizeAiOutput(string profile, NormalizationSettings settings)
    {
        return NormalizeUnifiedAiOutput(settings);
    }

    public int NormalizeUnifiedAiOutput(NormalizationSettings settings)
    {
        NormalizeUnifiedAiText(settings);
        return NormalizeUnifiedAiFormulas(settings);
    }

    public int NormalizeUnifiedAiText(NormalizationSettings settings)
    {
        if (settings == null) settings = NormalizationSettings.Load();
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic range = GetAiWorkingRange(document, settings.AutoPaste, settings.ProcessingScope);
        var text = (string)range.Text;
        if (String.IsNullOrWhiteSpace((text ?? String.Empty).Trim('\r', '\a'))) return 0;
        if (!settings.NormalizePlainTextWithoutMarkers && !HasTextNormalizationMarkers(text)) return 0;
        // 已有原生公式时只在原位置删除 Markdown 标记并设置段落样式，
        // 不再重写整段文字，因此标题和列表仍能规范且公式对象保持不变。
        if (RangeContainsOfficeMath(range))
        {
            var changed = NormalizeMarkdownInPlace(document, range);
            FinishAiWebLayout(document, (int)range.Start, (int)range.End);
            SelectScopeRange(range, settings.ProcessingScope);
            return changed > 0 ? 1 : 0;
        }
        var latexDocument = IsLatexDocumentSource(text);
        text = NormalizeWebAiFormulaText(text);
        text = Regex.Replace(text, @"\$\s{1,3}([0-9A-Za-z']{1,8})\s{1,3}\$", "$$$1$$");
        text = NormalizeLatexDocumentStructure(text);
        if (!latexDocument) text = NormalizeUnifiedAiStructure(text);
        var start = (int)range.Start;
        range.Text = text;
        // Word converts CRLF pairs to a single paragraph mark on assignment.
        // Reuse the adjusted COM range end instead of the source string length.
        range = document.Range(start, (int)range.End);
        range = RenderMarkdownRangeAsWeb(document, range, true);
        FinishAiWebLayout(document, (int)range.Start, (int)range.End);
        SelectScopeRange(range, settings.ProcessingScope);
        return 1;
    }

    public int NormalizeExamPaper(NormalizationSettings settings)
    {
        return NormalizeExamPaper(settings, 6, true);
    }

    public int NormalizeExamPaper(NormalizationSettings settings, int paperSize, bool landscape)
    {
        if (settings == null) settings = NormalizationSettings.Load();
        dynamic document = ((dynamic)_application).ActiveDocument;
        // 纸张、页边距和试卷结构属于整篇文档设置，不受选区开关影响。
        dynamic range = GetAiWorkingRange(document, settings.AutoPaste, "Document");
        var text = (string)range.Text;
        if (String.IsNullOrWhiteSpace((text ?? String.Empty).Trim('\r', '\a'))) return 0;
        if (RangeContainsOfficeMath(range))
        {
            ApplyExamPaperStyle(document, paperSize, landscape, settings);
            return 0;
        }
        var latexDocument = IsLatexDocumentSource(text);
        text = NormalizeWebAiFormulaText(text);
        text = Regex.Replace(text, @"\$\s{1,3}([0-9A-Za-z']{1,8})\s{1,3}\$", "$$$1$$");
        text = NormalizeLatexDocumentStructure(text);
        if (!latexDocument) text = NormalizeUnifiedAiStructure(text);
        text = NormalizeExamPaperSource(text);
        var pendingFormulas = ProtectAiFormulas(ref text);
        var start = (int)range.Start;
        range.Text = text;
        range = document.Range(start, (int)range.End);
        RenderMarkdownRangeAsWeb(document, range, true);
        PrepareExamChoiceParagraphs(document, settings.ChoiceColumns);
        var count = RestoreProtectedFormulas(document, pendingFormulas);
        FinishAiWebLayout(document);
        ApplyExamPaperStyle(document, paperSize, landscape, settings);
        return count;
    }

    private sealed class PendingAiFormula
    {
        public string Token;
        public string Formula;
        public bool Display;
    }

    private sealed class AiFormulaCandidate
    {
        public int Index;
        public int Length;
        public string Source;
        public string Formula;
        public bool Display;
    }

    private static List<PendingAiFormula> ProtectAiFormulas(ref string text)
    {
        var source = text ?? String.Empty;
        var matches = FindAiFormulaCandidates(source);
        var pending = new List<PendingAiFormula>();
        for (var index = matches.Count - 1; index >= 0; index--)
        {
            var match = matches[index];
            var token = "ZZZAIMATH" + index.ToString("D6") + "ZZZ";
            pending.Add(new PendingAiFormula { Token = token, Formula = match.Formula, Display = match.Display });
            source = source.Substring(0, match.Index) + token + source.Substring(match.Index + match.Length);
        }
        pending.Reverse();
        text = source;
        return pending;
    }

    private static int RestoreProtectedFormulas(dynamic document, List<PendingAiFormula> pending)
    {
        var converted = 0;
        foreach (var item in pending)
        {
            var located = FindLiteralFormulaRange(document, item.Token);
            if (located == null) continue;
            var formula = item.Display ? WrapLongFormulaForDisplay(item.Formula) : item.Formula;
            if (ReplaceWithOfficeMath(document, located.Item1, located.Item2, formula, item.Display)) converted++;
        }
        return converted;
    }

    public int NormalizeUnifiedAiFormulas(NormalizationSettings settings)
    {
        if (settings == null) settings = NormalizationSettings.Load();
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic range = GetAiWorkingRange(document, settings.AutoPaste, settings.ProcessingScope);
        var repaired = RepairExistingOfficeMath(document, range);
        var text = (string)range.Text;
        if (String.IsNullOrWhiteSpace((text ?? String.Empty).Trim('\r', '\a'))) return 0;
        var rangeStart = (int)range.Start;
        var rangeEnd = (int)range.End;
        var occupiedMath = (List<Tuple<int, int>>)GetOfficeMathSpans(document, rangeStart, rangeEnd);
        var matches = FindAiFormulaCandidates(text)
            .Where(match => !occupiedMath.Any(span => rangeStart + match.Index < span.Item2 &&
                                                     rangeStart + match.Index + match.Length > span.Item1))
            .ToList();
        if (matches.Count == 0) return repaired;
        var previousScreenUpdating = true;
        var previousPagination = true;
        var selectionStart = -1;
        var selectionEnd = -1;
        try
        {
            try { selectionStart = (int)document.Application.Selection.Start; selectionEnd = (int)document.Application.Selection.End; } catch { }
            try { previousScreenUpdating = (bool)document.Application.ScreenUpdating; document.Application.ScreenUpdating = false; } catch { }
            try { previousPagination = (bool)document.Application.Options.Pagination; document.Application.Options.Pagination = false; } catch { }
            var count = repaired;
            foreach (var match in matches.OrderByDescending(item => item.Index))
            {
                var formula = match.Formula.Trim();
                if (formula.Length == 0) continue;
                if (ReplaceWithOfficeMath(document, rangeStart + match.Index, match.Length,
                    match.Display ? WrapLongFormulaForDisplay(formula) : formula, match.Display)) count++;
            }
            FinishAiWebLayout(document, (int)range.Start, (int)range.End);
            return count;
        }
        finally
        {
            try { document.Application.Options.Pagination = previousPagination; } catch { }
            try { document.Application.ScreenUpdating = previousScreenUpdating; } catch { }
            if (String.Equals(settings.ProcessingScope, "Selection", StringComparison.OrdinalIgnoreCase))
            {
                SelectScopeRange(range, settings.ProcessingScope);
            }
            else if (selectionStart >= 0)
            {
                try
                {
                    var documentEnd = (int)document.Content.End;
                    document.Range(Math.Min(selectionStart, documentEnd), Math.Min(Math.Max(selectionStart, selectionEnd), documentEnd)).Select();
                }
                catch { }
            }
        }
    }

    private static bool RangeContainsOfficeMath(dynamic range)
    {
        try { return (int)range.OMaths.Count > 0; }
        catch { return false; }
    }

    private static void SelectScopeRange(dynamic range, string scope)
    {
        if (!String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase) || range == null) return;
        try { range.Select(); } catch { }
    }

    private static List<Tuple<int, int>> GetOfficeMathSpans(dynamic document, int start, int end)
    {
        var spans = new List<Tuple<int, int>>();
        var count = SafeCollectionCount(document.OMaths);
        for (var index = 1; index <= count; index++)
        {
            try
            {
                dynamic math = document.OMaths[index];
                var mathStart = (int)math.Range.Start;
                var mathEnd = (int)math.Range.End;
                if (mathStart < end && mathEnd > start) spans.Add(Tuple.Create(mathStart, mathEnd));
            }
            catch { }
        }
        return spans;
    }

    private static int RepairExistingOfficeMath(dynamic document, dynamic workingRange)
    {
        var repaired = 0;
        var start = 0;
        var end = 0;
        try { start = (int)workingRange.Start; end = (int)workingRange.End; }
        catch { return 0; }
        var mathCount = SafeCollectionCount(document.OMaths);
        for (var index = mathCount; index >= 1; index--)
        {
            dynamic math;
            try { math = document.OMaths[index]; } catch { continue; }
            if (!MathRangesOverlap(math, start, end)) continue;
            string visible;
            try { visible = math.Range.Text as string ?? String.Empty; } catch { continue; }
            if (!Regex.IsMatch(visible,
                    @"\\(?:dfrac|tfrac|frac|Biggl|Biggr|biggl|biggr|Biggm|biggm|Bigl|Bigr|bigl|bigr|Bigm|bigm|Bigg|bigg|Big|big|mleft|mright|left|right|middle|ln|log|sin|cos|tan)(?=\b|\d|\{|\s|\()") &&
                !HasUnmatchedClosingParenthesis(visible)) continue;
            var original = visible;
            try
            {
                try { math.Linearize(); } catch { }
                dynamic mathRange = math.Range;
                var linear = mathRange.Text as string ?? original;
                var normalized = RemoveUnmatchedClosingParentheses(ConvertAiLatexExpression(linear));
                if (String.IsNullOrWhiteSpace(normalized)) continue;
                mathRange.Text = normalized;
                if (!TryBuildMath(math))
                {
                    dynamic current = ResolveCreatedMath(document, mathRange, mathRange, (int)mathRange.Start, (int)mathRange.End);
                    if (current == null || !TryBuildMath(current)) throw new InvalidOperationException();
                }
                repaired++;
            }
            catch
            {
                try { math.Range.Text = original; TryBuildMath(math); } catch { }
            }
        }
        return repaired;
    }

    public IList<FormulaPreviewItem> GetFormulaPreview(NormalizationSettings settings)
    {
        if (settings == null) settings = NormalizationSettings.Load();
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic range = GetAiWorkingRange(document, settings.AutoPaste, settings.ProcessingScope);
        var source = range.Text as string ?? String.Empty;
        var baseStart = (int)range.Start;
        var result = new List<FormulaPreviewItem>();
        foreach (var match in FindAiFormulaCandidates(source))
        {
            var formula = match.Formula.Trim();
            if (formula.Length == 0) continue;
            result.Add(new FormulaPreviewItem
            {
                Start = baseStart + match.Index,
                Length = match.Length,
                Source = match.Source,
                Formula = formula,
                Preview = ConvertAiLatexExpression(formula),
                Display = match.Display
            });
        }
        return result;
    }

    public int ConvertFormulaPreview(IEnumerable<FormulaPreviewItem> selected)
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        var count = 0;
        foreach (var item in (selected ?? Enumerable.Empty<FormulaPreviewItem>()).OrderByDescending(value => value.Start))
        {
            if (item == null || item.Length <= 0 || String.IsNullOrWhiteSpace(item.Formula)) continue;
            try
            {
                dynamic current = document.Range(item.Start, item.Start + item.Length);
                var existing = current.Text as string ?? String.Empty;
                if (!String.Equals(existing, item.Source, StringComparison.Ordinal)) continue;
                if (ReplaceWithOfficeMath(document, item.Start, item.Length, item.Formula, item.Display)) count++;
            }
            catch { }
        }
        return count;
    }

    public string GetCompatibilityReport()
    {
        var lines = new List<string> { "AI规范化 运行检查", "时间：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") };
        dynamic application = _application;
        try { lines.Add("宿主：" + Convert.ToString(application.Name)); } catch { lines.Add("宿主：未知"); }
        try { lines.Add("宿主版本：" + Convert.ToString(application.Version)); } catch { lines.Add("宿主版本：未知"); }
        lines.Add("进程位数：" + (Environment.Is64BitProcess ? "64 位" : "32 位"));
        lines.Add(".NET：" + Environment.Version);
        try
        {
            dynamic document = application.ActiveDocument;
            lines.Add("活动文档：正常（" + Convert.ToString(document.Name) + "）");
            try { var count = (int)document.OMaths.Count; lines.Add("Office 原生公式接口：正常（" + count + " 个公式）"); }
            catch { lines.Add("Office 原生公式接口：当前宿主未完整公开"); }
            try { var count = (int)document.Tables.Count; lines.Add("表格接口：正常（" + count + " 个表格）"); }
            catch { lines.Add("表格接口：当前宿主未完整公开"); }
            try { var pages = (int)document.ComputeStatistics(2); lines.Add("分页接口：正常（" + pages + " 页）"); }
            catch { lines.Add("分页接口：当前宿主未完整公开"); }
        }
        catch { lines.Add("活动文档：请先打开文档后复查"); }
        lines.Add("演示组件：" + (Type.GetTypeFromProgID("PowerPoint.Application") != null || Type.GetTypeFromProgID("KWPP.Application") != null ? "已检测" : "未检测"));
        lines.Add("项目主页：" + UpdateChecker.RepositoryUrl);
        return String.Join("\r\n", lines);
    }

    private static bool LooksLikeExamPaper(string text)
    {
        var value = text ?? String.Empty;
        return (value.Contains("试卷") || value.Contains("考试")) &&
               (value.Contains("注意事项") || value.Contains("选择题")) &&
               Regex.IsMatch(value, @"(?:\A|[\r\n])\s*1[\.．、]\s*");
    }

    private static string NormalizeExamPaperSource(string text)
    {
        var lines = Regex.Split(text ?? String.Empty, @"\r\n|\r|\n").ToList();
        var output = new List<string>();
        for (var index = 0; index < lines.Count; index++)
        {
            var line = lines[index].TrimEnd();
            if (Regex.IsMatch(line, @"^A[\.．、]\s*") && !Regex.IsMatch(line, @"\sB[\.．、]\s*"))
            {
                var options = new List<string> { line.Trim() };
                var expected = 'B';
                while (index + 1 < lines.Count && expected <= 'D' && Regex.IsMatch(lines[index + 1], "^\\s*" + expected + @"[\.．、]\s*"))
                {
                    options.Add(lines[++index].Trim());
                    expected++;
                }
                if (output.Count > 0 && output[output.Count - 1].Length > 0) output.Add(String.Empty);
                output.Add(String.Join("　", options));
                output.Add(String.Empty);
                continue;
            }
            if (Regex.IsMatch(line, @"^A[\.．、]\s*") && output.Count > 0 && output[output.Count - 1].Length > 0) output.Add(String.Empty);
            output.Add(line);
            if (Regex.IsMatch(line, @"^A[\.．、].*\sD[\.．、]\s*")) output.Add(String.Empty);
        }
        return String.Join("\r", output);
    }

    private static void ApplyExamPaperStyle(dynamic document, int paperSize, bool landscape, NormalizationSettings settings)
    {
        try
        {
            var sectionCount = SafeCollectionCount(document.Sections);
            for (var sectionIndex = 1; sectionIndex <= sectionCount; sectionIndex++)
            {
                dynamic section;
                try { section = document.Sections[sectionIndex]; } catch { continue; }
                dynamic setup = section.PageSetup;
                try { setup.TextColumns.SetCount(landscape ? 2 : 1); } catch { }
                try { setup.TextColumns.Spacing = 26f; } catch { }
                try { setup.Orientation = landscape ? 1 : 0; } catch { }
                try { setup.PaperSize = paperSize == 7 ? 7 : 6; } catch { }
                try
                {
                    var shortSide = paperSize == 7 ? 595.28f : 841.89f;
                    var longSide = paperSize == 7 ? 841.89f : 1190.55f;
                    setup.PageWidth = landscape ? longSide : shortSide;
                    setup.PageHeight = landscape ? shortSide : longSide;
                }
                catch { }
                var margin = Math.Max(0.8f, Math.Min(3.5f, settings.ExamMarginCm)) * 28.3465f;
                try { setup.TopMargin = margin; } catch { }
                try { setup.BottomMargin = margin; } catch { }
                try { setup.LeftMargin = margin; } catch { }
                try { setup.RightMargin = margin; } catch { }
                try { setup.HeaderDistance = 18f; } catch { }
                try { setup.FooterDistance = 18f; } catch { }
            }
        }
        catch { }

        try
        {
            dynamic whole = document.Content;
            whole.Font.NameFarEast = settings.FontName;
            whole.Font.Name = "Times New Roman";
            whole.Font.Size = settings.FontSize;
            whole.Font.Color = 0;
            whole.ParagraphFormat.LineSpacingRule = 0;
            whole.ParagraphFormat.SpaceAfter = settings.ParagraphSpaceAfter;
        }
        catch { }

        try
        {
            var paragraphCount = SafeCollectionCount(document.Paragraphs);
            for (var paragraphIndex = 1; paragraphIndex <= paragraphCount; paragraphIndex++)
            {
                dynamic paragraph;
                try { paragraph = document.Paragraphs[paragraphIndex]; } catch { continue; }
                dynamic range = paragraph.Range;
                var raw = range.Text as string ?? String.Empty;
                var value = raw.Trim('\r', '\a', ' ', '\t');
                if (value.Length == 0) continue;
                range.ParagraphFormat.KeepTogether = -1;
                range.ParagraphFormat.WidowControl = -1;

                if ((value.Contains("普通高等学校招生全国统一考试") && value.Contains("试卷")) ||
                    (value.Contains("高考") && value.Contains("数学")))
                {
                    var secondLine = value.IndexOf("数学（", StringComparison.Ordinal);
                    if (secondLine > 0) document.Range((int)range.Start + secondLine, (int)range.Start + secondLine).InsertBefore("\v");
                    range.ParagraphFormat.Alignment = 1;
                    range.ParagraphFormat.SpaceBefore = 4f;
                    range.ParagraphFormat.SpaceAfter = 14f;
                    range.Font.NameFarEast = "黑体";
                    range.Font.Size = 17f;
                    range.Font.Bold = 1;
                }
                else if (value.StartsWith("本试卷共", StringComparison.Ordinal))
                {
                    range.ParagraphFormat.Alignment = 1;
                    range.ParagraphFormat.SpaceAfter = 7f;
                    range.Font.Size = 10.5f;
                }
                else if (value == "注意事项" || value.StartsWith("第一部分", StringComparison.Ordinal) || value.StartsWith("第二部分", StringComparison.Ordinal))
                {
                    range.Font.NameFarEast = "黑体";
                    range.Font.Bold = 1;
                    range.Font.Size = 11f;
                    range.ParagraphFormat.SpaceBefore = 5f;
                    range.ParagraphFormat.SpaceAfter = 4f;
                }
                else if (Regex.IsMatch(value, @"^[一二三四五六七八九十]+[、．.]"))
                {
                    range.Font.NameFarEast = "黑体";
                    range.Font.Bold = 1;
                    range.Font.Size = 11f;
                    range.ParagraphFormat.SpaceBefore = 6f;
                    range.ParagraphFormat.SpaceAfter = 4f;
                    range.ParagraphFormat.KeepWithNext = -1;
                }
                else if (Regex.IsMatch(value, @"^\d+[．.]\s*[（(]\s*\d+\s*分"))
                {
                    range.Font.NameFarEast = "黑体";
                    range.Font.Bold = 1;
                    range.Font.Size = 11f;
                    range.ParagraphFormat.SpaceBefore = 6f;
                    range.ParagraphFormat.SpaceAfter = 4f;
                    range.ParagraphFormat.KeepWithNext = -1;
                }
                else if (Regex.IsMatch(value, @"^\d+[\.．、]\s*"))
                {
                    range.ParagraphFormat.FirstLineIndent = 0f;
                    range.ParagraphFormat.LeftIndent = 0f;
                    range.ParagraphFormat.SpaceBefore = 2f;
                    range.ParagraphFormat.SpaceAfter = 3f;
                }
                else if (Regex.IsMatch(value, @"^A[\.．、]\s*"))
                {
                    try
                    {
                        range.ParagraphFormat.TabStops.ClearAll();
                        range.ParagraphFormat.TabStops.Add(122f);
                        range.ParagraphFormat.TabStops.Add(252f);
                        range.ParagraphFormat.TabStops.Add(382f);
                    }
                    catch { }
                    range.ParagraphFormat.LeftIndent = 0f;
                    range.ParagraphFormat.SpaceAfter = 5f;
                }
                else if (Regex.IsMatch(value, @"^[（(](?:\d+|[ivxIVX]+)[）)]\s*"))
                {
                    range.ParagraphFormat.LeftIndent = 10.5f;
                    range.ParagraphFormat.FirstLineIndent = -10.5f;
                    range.ParagraphFormat.SpaceAfter = 3f;
                }
                else if (Regex.IsMatch(value, @"^第\s*\d+\s*页[，,]\s*共\s*\d+\s*页"))
                {
                    range.ParagraphFormat.Alignment = 1;
                    range.Font.Size = 9f;
                }
            }
        }
        catch { }

        try
        {
            var tableCount = SafeCollectionCount(document.Tables);
            for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
            {
                dynamic table;
                try { table = document.Tables[tableIndex]; } catch { continue; }
                table.Borders.Enable = 0;
                table.AllowAutoFit = true;
                table.AutoFitBehavior(2);
                table.Range.Font.NameFarEast = "宋体";
                table.Range.Font.Name = "Times New Roman";
                table.Range.Font.Size = 10f;
                table.Range.Font.Bold = 0;
                table.Range.Font.Color = 0;
                table.Range.Shading.BackgroundPatternColor = 16777215;
                table.Range.ParagraphFormat.SpaceAfter = 0f;
            }
        }
        catch { }
    }

    private static void PrepareExamChoiceParagraphs(dynamic document, int columns)
    {
        try
        {
            var paragraphCount = SafeCollectionCount(document.Paragraphs);
            for (var paragraphIndex = 1; paragraphIndex <= paragraphCount; paragraphIndex++)
            {
                dynamic paragraph;
                try { paragraph = document.Paragraphs[paragraphIndex]; } catch { continue; }
                dynamic range = paragraph.Range;
                var raw = range.Text as string ?? String.Empty;
                var body = raw.TrimEnd('\r', '\a');
                if (!Regex.IsMatch(body, @"^\s*A[\.．、]\s*") || !Regex.IsMatch(body, @"\s+B[\.．、]\s*")) continue;
                body = Regex.Replace(body, @"[ \t\u3000]+(?=[BCD][\.．、]\s*)", "\t");
                if (columns <= 1) body = Regex.Replace(body, @"\t(?=[BCD][\.．、]\s*)", "\r");
                else if (columns == 2) body = Regex.Replace(body, @"\t(?=C[\.．、]\s*)", "\r");
                dynamic writable = range.Duplicate;
                writable.End = writable.End - 1;
                writable.Text = body;
            }
        }
        catch { }
    }

    private static bool HasTextNormalizationMarkers(string text)
    {
        var value = text ?? String.Empty;
        if (Regex.IsMatch(value, @"(?m)^\s{0,3}(?:#{1,6}\s+|```|~~~|[-+*]\s+|\d+[\.、]\s+|>\s+|\|.*\|\s*$)")) return true;
        if (Regex.IsMatch(value, @"(?is)<(?:h[1-6]|p|table|ul|ol|li|blockquote)\b")) return true;
        return Regex.IsMatch(value, @"\\(?:documentclass|usepackage|geometry|begin\{document\}|(?:subsubsection|subsection|section)\*?\{|begin\{(?:enumerate|itemize|description|tabular|tabularx|longtable|theorem|lemma|definition|proof)\})");
    }

    private static bool IsLatexDocumentSource(string text)
    {
        return Regex.IsMatch(text ?? String.Empty, @"\\documentclass(?:\[[^\]]*\])?\{[^}]+\}|\\begin\{document\}");
    }

    private static string NormalizeLatexDocumentStructure(string text)
    {
        var value = text ?? String.Empty;
        if (!Regex.IsMatch(value, @"\\(?:documentclass|begin\{document\}|(?:subsubsection|subsection|section)\*?\{|begin\{(?:center|enumerate|itemize|description|tabular|tabularx|longtable|theorem|lemma|proposition|corollary|definition|example|remark|proof)\}|(?:tiny|scriptsize|footnotesize|small|normalsize|large|Large|LARGE|huge|Huge)\s*\{)")) return value;
        // Word stores paragraph boundaries as CR-only characters, while pasted
        // TeX source commonly uses LF or CRLF.  Normalize all three forms before
        // applying line-oriented preamble, list and table rules.
        value = value.Replace("\r\n", "\n").Replace('\r', '\n');

        value = Regex.Replace(value, @"(?m)^\s*%(?!%).*$", String.Empty);
        value = Regex.Replace(value, @"(?m)^\s*\\(?:documentclass|usepackage|geometry)(?:\[[^\]]*\])?\{[^\r\n]*\}\s*$", String.Empty);
        value = Regex.Replace(value, @"(?m)^\s*\\(?:begin|end)\{document\}\s*$", String.Empty);
        // WPS may prepend its equation placeholder when TeX is pasted as plain text.
        // It is source noise only when a structural TeX command follows it.
        value = Regex.Replace(value, @"(?m)^\s*在此处键入公式。\s*(?=\\(?:begin|end)\{|vspace|hspace|(?:tiny|scriptsize|footnotesize|small|normalsize|large|Large|LARGE|huge|Huge)\b)", String.Empty);
        value = Regex.Replace(value, @"\\(?:vspace|hspace)\*?\s*\{[^{}]*\}", String.Empty);
        value = Regex.Replace(value, @"\\(?:tiny|scriptsize|footnotesize|small|normalsize|large|Large|LARGE|huge|Huge)\s*\{([^{}\r\n]*)\}", "$1");
        value = Regex.Replace(value, @"\\(?:tiny|scriptsize|footnotesize|small|normalsize|large|Large|LARGE|huge|Huge)\b\s*", String.Empty);
        value = value.Replace("\\star", "★");

        // Preserve the visual intent of common TeX title blocks instead of leaving
        // begin/end/spacing commands visible in WPS.
        value = Regex.Replace(value, @"(?ms)^\s*\\begin\{center\}\s*(?<body>.*?)\s*\\end\{center\}\s*", match =>
        {
            var lines = Regex.Split(match.Groups["body"].Value, @"\r?\n")
                .Select(line => Regex.Replace(line.Trim(), @"\\\\\s*$", String.Empty).Trim())
                .Where(line => line.Length > 0)
                .ToArray();
            if (lines.Length == 0) return String.Empty;
            var centered = new StringBuilder();
            for (var index = 0; index < lines.Length; index++)
            {
                var line = lines[index];
                if (Regex.IsMatch(line, @"(?:考试|试卷|测试|数学|语文|英语|物理|化学|生物)")) centered.Append("# ");
                centered.Append(line).Append("\n");
            }
            return centered.ToString();
        });
        value = Regex.Replace(value, @"\\(?:begin|end)\{center\}", String.Empty);
        value = Regex.Replace(value, @"\\(?:centering|raggedright|raggedleft)\b\s*", String.Empty);
        value = Regex.Replace(value, @"\\(?:newpage|clearpage|pagebreak)\b", "\n\n");
        value = Regex.Replace(value, @"(?m)^\s*\\noindent\s*(?<title>[^\n]+)(?=\n(?:\s*\n)*\s*\\section)", "# ${title}");
        value = Regex.Replace(value, @"\\noindent\b\s*", String.Empty);
        value = ReplaceLatexHeadings(value);
        value = Regex.Replace(value, @"\\(?:textbf|emph)\{([^{}\r\n]*)\}", "$1");

        var environments = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "theorem", "定理" }, { "lemma", "引理" }, { "proposition", "命题" }, { "corollary", "推论" },
            { "definition", "定义" }, { "example", "例" }, { "remark", "说明" }, { "proof", "证明" }
        };
        value = Regex.Replace(value, @"\\begin\{(?<name>theorem|lemma|proposition|corollary|definition|example|remark|proof)\}(?:\[(?<title>[^\]]+)\])?", match =>
        {
            var title = environments[match.Groups["name"].Value];
            var detail = match.Groups["title"].Value.Trim();
            return "\n### " + title + (detail.Length > 0 ? "（" + detail + "）" : String.Empty) + "\n";
        }, RegexOptions.IgnoreCase);
        value = Regex.Replace(value, @"\\end\{(?:theorem|lemma|proposition|corollary|definition|example|remark|proof)\}", "\n", RegexOptions.IgnoreCase);

        value = Regex.Replace(value,
            @"(?ms)^\s*\\begin\{(?:tabular|tabularx|longtable)\}(?:\{[^\r\n]*\}){0,2}\s*\r?\n(?<body>.*?)^\s*\\end\{(?:tabular|tabularx|longtable)\}\s*$",
            match => ConvertLatexTabularToMarkdown(match.Groups["body"].Value));

        var output = new StringBuilder();
        var enumerateDepth = 0;
        var counters = new List<int>();
        foreach (var rawLine in Regex.Split(value, @"\r?\n"))
        {
            var line = rawLine.Trim();
            if (Regex.IsMatch(line, @"^\\begin\{(?:enumerate|itemize|description)\}"))
            {
                enumerateDepth++;
                counters.Add(0);
                continue;
            }
            if (Regex.IsMatch(line, @"^\\end\{(?:enumerate|itemize|description)\}"))
            {
                if (enumerateDepth > 0) enumerateDepth--;
                if (counters.Count > 0) counters.RemoveAt(counters.Count - 1);
                continue;
            }

            var item = Regex.Match(line, @"^\\item(?:\[(?<label>[^\]]+)\])?\s*(?<body>.*)$");
            if (item.Success)
            {
                var label = item.Groups["label"].Value.Trim();
                string prefix;
                if (label.Length > 0) prefix = label + " ";
                else if (enumerateDepth <= 1)
                {
                    if (counters.Count == 0) counters.Add(0);
                    counters[counters.Count - 1]++;
                    prefix = counters[counters.Count - 1] + ". ";
                }
                else prefix = "- ";
                line = prefix + item.Groups["body"].Value;
            }
            line = Regex.Replace(line, @"\\\\\s*$", String.Empty);
            output.Append(line).Append("\r\n");
        }
        return Regex.Replace(output.ToString(), @"(?:\r\n){3,}", "\r\n\r\n").Trim();
    }

    private static string ReplaceLatexHeadings(string source)
    {
        var commands = new[]
        {
            new { Name = "\\subsubsection", Prefix = "#### " },
            new { Name = "\\subsection", Prefix = "### " },
            new { Name = "\\section", Prefix = "## " }
        };
        foreach (var command in commands)
        {
            var searchFrom = 0;
            while (searchFrom < source.Length)
            {
                var commandIndex = source.IndexOf(command.Name, searchFrom, StringComparison.Ordinal);
                if (commandIndex < 0) break;
                var cursor = commandIndex + command.Name.Length;
                if (cursor < source.Length && source[cursor] == '*') cursor++;
                while (cursor < source.Length && Char.IsWhiteSpace(source[cursor])) cursor++;
                if (!TryReadGroup(source, cursor, out var title, out var end))
                {
                    searchFrom = commandIndex + command.Name.Length;
                    continue;
                }
                var replacement = command.Prefix + title.Trim();
                source = source.Substring(0, commandIndex) + replacement + source.Substring(end);
                searchFrom = commandIndex + replacement.Length;
            }
        }
        return source;
    }

    private static string ConvertLatexTabularToMarkdown(string body)
    {
        var rows = new List<string[]>();
        foreach (var rawLine in Regex.Split(body ?? String.Empty, @"\r?\n"))
        {
            var line = rawLine.Trim();
            line = Regex.Replace(line, @"^\\(?:hline|toprule|midrule|bottomrule|cline\{[^}]+\})\s*", String.Empty);
            line = Regex.Replace(line, @"\\(?:hline|toprule|midrule|bottomrule|cline\{[^}]+\})\s*$", String.Empty);
            line = Regex.Replace(line, @"\\\\\s*$", String.Empty).Trim();
            if (line.Length == 0 || Regex.IsMatch(line, @"^\\(?:hline|toprule|midrule|bottomrule)$")) continue;
            rows.Add(line.Split('&').Select(cell =>
            {
                var value = cell.Trim();
                value = Regex.Replace(value, @"\\multicolumn\{\d+\}\{[^{}]*\}\{(?<text>[^{}]*)\}", "${text}");
                value = Regex.Replace(value, @"\\multirow\{\d+\}\{[^{}]*\}\{(?<text>[^{}]*)\}", "${text}");
                value = Regex.Replace(value, @"\\(?:textbf|emph|mathrm)\{([^{}]*)\}", "$1");
                return value;
            }).ToArray());
        }
        if (rows.Count == 0) return String.Empty;
        var columns = rows.Max(row => row.Length);
        var output = new StringBuilder();
        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var row = rows[rowIndex];
            output.Append("| ");
            for (var column = 0; column < columns; column++) output.Append(column < row.Length ? row[column] : String.Empty).Append(" | ");
            output.Append("\r\n");
            if (rowIndex == 0)
            {
                output.Append("| ");
                for (var column = 0; column < columns; column++) output.Append("--- | ");
                output.Append("\r\n");
            }
        }
        return output.ToString().TrimEnd();
    }

    private static string NormalizeUnifiedAiStructure(string text)
    {
        var value = text ?? String.Empty;
        value = Regex.Replace(value, @"(?<![#\r\n])(?=#{1,6}\s*)", "\r\r");
        value = Regex.Replace(value, @"(?:\A|(?<=[\r\n]))\s*公式[：:]\s*", String.Empty);
        value = Regex.Replace(value, @"(?<=[^\r\n])(?<!# )(?<!## )(?<!### )(?=[一二三四五六七八九十百]{1,3}、)", "\r\r");
        value = Regex.Replace(value, @"(?:\A|(?<=[\r\n]))(?<head>[一二三四五六七八九十百]{1,3}、[^\r\n]{1,28}?（[^）\r\n]{1,28}）)(?=\S)", "${head}\r");
        value = Regex.Replace(value, @"八、积化和差\s*&\s*和差化积积化和差(?=\\\()", "八、积化和差 & 和差化积\r\r### 积化和差\r");
        value = Regex.Replace(value, @"(?:\A|(?<=[\r\n]))(?<head>[一二三四五六七八九十百]{1,3}、[^\\\r\n]{1,28}?)(?=\\\()", "${head}\r");
        value = Regex.Replace(value, @"(?<=[\p{IsCJKUnifiedIdeographs}）)])(?=\d+[\.．、]\s*)", "\r");
        value = Regex.Replace(value, @"\\\)(?=[一二三四五六七八九十百]{1,3}、)", "\\)\r\r");
        value = Regex.Replace(value, @"\\\)(?=(?:积化和差|和差化积|正弦定理|余弦定理|三角形面积)\\\()", "\\)\r\r");
        value = Regex.Replace(value, @"(?<!^)(?<![\r\n])(?=(?:正弦定理|余弦定理|三角形面积)\\\()", "\r\r");
        value = Regex.Replace(value, @"(?:\A|(?<=[\r\n]))(?<sub>(?:积化和差|和差化积|正弦定理|余弦定理|三角形面积))(?=\\\()", "### ${sub}\r");
        value = Regex.Replace(value, @"(?:\A|(?<=[\r\n]))(?<number>\d+[\.．、])\s*", "### ${number} ");
        value = Regex.Replace(value, @"(?:\A|(?<=[\r\n]))(?!##\s)(?<section>[一二三四五六七八九十百]{1,3}、[^\r\n]+)(?=\r|\n|\z)", "## ${section}");
        return value.TrimStart('\r', '\n');
    }

    private dynamic GetAiWorkingRange(dynamic document, bool autoPaste, string scope)
    {
        dynamic selection = ((dynamic)_application).Selection;
        dynamic selected = selection.Range.Duplicate;
        var selectedText = selected.Text as string;
        if (String.Equals(scope, "Document", StringComparison.OrdinalIgnoreCase)) return document.Content;
        if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase))
        {
            if (String.IsNullOrWhiteSpace((selectedText ?? String.Empty).Trim('\r', '\a'))) throw new InvalidOperationException("请先选中需要规范的文本。");
            return selected;
        }
        if (!String.IsNullOrWhiteSpace((selectedText ?? String.Empty).Trim('\r', '\a'))) return selected;
        var existingText = document.Content.Text as string;
        if (!String.IsNullOrWhiteSpace((existingText ?? String.Empty).Trim('\r', '\a'))) return document.Content;
        if (!autoPaste) return document.Content;

        var insertionStart = (int)selected.Start;
        var beforeEnd = (int)document.Content.End;
        selection.Paste();
        var afterEnd = (int)document.Content.End;
        var added = Math.Max(0, afterEnd - beforeEnd);
        if (added == 0)
        {
            dynamic afterSelection = selection.Range;
            var afterStart = (int)afterSelection.Start;
            var afterSelectionEnd = (int)afterSelection.End;
            var pastedEnd = Math.Max(afterStart, afterSelectionEnd);
            if (pastedEnd > insertionStart) return document.Range(insertionStart, pastedEnd);
            return document.Content;
        }
        return document.Range(insertionStart, insertionStart + added);
    }

    private bool SelectionHasText()
    {
        try
        {
            dynamic range = ((dynamic)_application).Selection.Range;
            var text = range.Text as string;
            return !String.IsNullOrWhiteSpace((text ?? String.Empty).Trim('\r', '\a'));
        }
        catch { return false; }
    }

    public bool HasSelectedText()
    {
        return SelectionHasText();
    }

    public int ConvertCustomDelimitedFormulas(NormalizationSettings settings)
    {
        if (settings == null) settings = NormalizationSettings.Load();
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic content = document.Content;
        var source = (string)content.Text;
        var patterns = new[]
        {
            Tuple.Create(settings.InlinePrefix, settings.InlineSuffix, false),
            Tuple.Create(settings.DisplayPrefix, settings.DisplaySuffix, true),
            Tuple.Create(settings.MultiPrefix, settings.MultiSuffix, true)
        };
        var matches = new System.Collections.Generic.List<Tuple<int, int, string, bool>>();
        foreach (var pair in patterns)
        {
            if (String.IsNullOrEmpty(pair.Item1) || String.IsNullOrEmpty(pair.Item2)) continue;
            var start = 0;
            while (start < source.Length)
            {
                var left = source.IndexOf(pair.Item1, start, StringComparison.Ordinal);
                if (left < 0) break;
                var bodyStart = left + pair.Item1.Length;
                var right = source.IndexOf(pair.Item2, bodyStart, StringComparison.Ordinal);
                if (right < 0) break;
                matches.Add(Tuple.Create(left, right + pair.Item2.Length - left, source.Substring(bodyStart, right - bodyStart), pair.Item3));
                start = right + pair.Item2.Length;
            }
        }
        foreach (var item in matches.OrderByDescending(item => item.Item1)) ReplaceWithOfficeMath(document, content.Start + item.Item1, item.Item2, item.Item3, item.Item4);
        return matches.Count;
    }

    public int ConvertPrefixlessFormulas()
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic content = document.Content;
        var source = (string)content.Text;
        var matches = FindAiFormulaCandidates(source);
        var converted = 0;
        for (var index = matches.Count - 1; index >= 0; index--)
        {
            var match = matches[index];
            if (ReplaceWithOfficeMath(document, content.Start + match.Index, match.Length, match.Formula, match.Display)) converted++;
        }
        return converted;
    }


    public int ConvertLatexToOfficeMath()
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic content = document.Content;
        var source = (string)content.Text;
        var matches = FindAiFormulaCandidates(source);
        var converted = 0;
        for (var index = matches.Count - 1; index >= 0; index--)
        {
            var match = matches[index];
            var wordFormula = match.Display ? WrapLongFormulaForDisplay(match.Formula) : match.Formula;
            if (ReplaceWithOfficeMath(document, (int)content.Start + match.Index, match.Length, wordFormula, match.Display)) converted++;
        }
        return converted;
    }

    private static Tuple<int, int>[] LocateFormulaRanges(dynamic document, MatchCollection matches)
    {
        var result = new Tuple<int, int>[matches.Count];
        var contentStart = (int)document.Content.Start;
        var contentEnd = (int)document.Content.End;
        for (var directIndex = 0; directIndex < matches.Count; directIndex++)
        {
            var directMatch = matches[directIndex];
            var directStart = contentStart + directMatch.Index;
            var directEnd = Math.Min(contentEnd, directStart + directMatch.Length);
            try
            {
                dynamic directRange = document.Range(directStart, directEnd);
                var directText = directRange.Text as string ?? String.Empty;
                if (String.Equals(directText, directMatch.Value, StringComparison.Ordinal))
                    result[directIndex] = Tuple.Create(directStart, directMatch.Length);
            }
            catch { }
        }
        var cursor = (int)document.Content.Start;
        for (var index = 0; index < matches.Count; index++)
        {
            if (result[index] != null)
            {
                cursor = result[index].Item1 + result[index].Item2;
                continue;
            }
            var match = matches[index];
            string opening = null, closing = null;
            if (match.Groups["display"].Success) { opening = "$$"; closing = "$$"; }
            else if (match.Groups["dollar"].Success) { opening = "$"; closing = "$"; }
            else if (match.Groups["paren"].Success) { opening = "\\("; closing = "\\)"; }
            else if (match.Groups["bracket"].Success || match.Groups["block"].Success) { opening = "\\["; closing = "\\]"; }
            if (opening == null) continue;

            var openRange = FindNextTextRange(document, opening, cursor);
            if (openRange == null) continue;
            var closeRange = FindNextTextRange(document, closing, openRange.Item1 + openRange.Item2);
            if (closeRange == null) continue;
            var end = closeRange.Item1 + closeRange.Item2;
            result[index] = Tuple.Create(openRange.Item1, end - openRange.Item1);
            cursor = end;
        }
        return result;
    }

    private static Tuple<int, int> FindNextTextRange(dynamic document, string text, int afterStart)
    {
        try
        {
            var end = (int)document.Content.End;
            if (afterStart >= end) return null;
            dynamic range = document.Range(afterStart, end);
            dynamic find = range.Find;
            find.ClearFormatting();
            find.Text = text;
            find.Forward = true;
            find.Wrap = 0;
            find.Format = false;
            find.MatchWildcards = false;
            return (bool)find.Execute() ? Tuple.Create((int)range.Start, (int)(range.End - range.Start)) : null;
        }
        catch { return null; }
    }

    private static Tuple<int, int> FindLiteralFormulaRange(dynamic document, string text)
    {
        if (String.IsNullOrEmpty(text)) return null;
        try
        {
            dynamic range = document.Content.Duplicate;
            dynamic find = range.Find;
            find.ClearFormatting();
            find.Text = text.Replace("\r\n", "^p").Replace("\r", "^p").Replace("\n", "^p");
            find.Forward = false;
            find.Wrap = 0;
            find.Format = false;
            find.MatchWildcards = false;
            if ((bool)find.Execute()) return Tuple.Create((int)range.Start, (int)(range.End - range.Start));
        }
        catch { }
        return null;
    }

    private static Tuple<int, int> FindDelimitedFormulaRange(dynamic document, Match match)
    {
        string opening = null;
        string closing = null;
        if (match.Groups["display"].Success) { opening = "$$"; closing = "$$"; }
        else if (match.Groups["dollar"].Success) { opening = "$"; closing = "$"; }
        else if (match.Groups["paren"].Success) { opening = "\\("; closing = "\\)"; }
        else if (match.Groups["bracket"].Success || match.Groups["block"].Success) { opening = "\\["; closing = "\\]"; }
        else return null;

        var closeRange = FindLastTextRange(document, closing, null);
        if (closeRange == null) return null;
        var openRange = FindLastTextRange(document, opening, closeRange.Item1);
        if (openRange == null) return null;

        // $$ 与 $ 的起止标记相同；结束标记搜索到最后一个，起始标记只在其前方搜索。
        var end = closeRange.Item1 + closeRange.Item2;
        return end > openRange.Item1 ? Tuple.Create(openRange.Item1, end - openRange.Item1) : null;
    }

    private static Tuple<int, int> FindLastTextRange(dynamic document, string text, int? beforeEnd)
    {
        try
        {
            var start = (int)document.Content.Start;
            var end = beforeEnd ?? (int)document.Content.End;
            if (end <= start) return null;
            dynamic range = document.Range(start, end);
            dynamic find = range.Find;
            find.ClearFormatting();
            find.Text = text;
            find.Forward = false;
            find.Wrap = 0;
            find.Format = false;
            find.MatchWildcards = false;
            return (bool)find.Execute() ? Tuple.Create((int)range.Start, (int)(range.End - range.Start)) : null;
        }
        catch { return null; }
    }

    private static bool IsDisplayStyleFormula(string formula)
    {
        var value = formula ?? String.Empty;
        if (Regex.IsMatch(value, @"\\begin\s*\{(?:math|displaymath|equation\*?|align\*?|aligned|alignedat|alignat\*?|flalign\*?|gather\*?|gathered|multline\*?|multlined|split|cases|dcases|array|matrix|smallmatrix|pmatrix|bmatrix|Bmatrix|vmatrix|Vmatrix)\}", RegexOptions.IgnoreCase)) return true;
        if (value.IndexOf('\r') >= 0 || value.IndexOf('\n') >= 0) return true;
        return value.Length >= 90 && Regex.Matches(value, @"\\quad\b").Count >= 2;
    }

    private static string WrapLongFormulaForDisplay(string formula)
    {
        var value = (formula ?? String.Empty).Trim();
        if (Regex.IsMatch(value, @"\\begin\s*\{", RegexOptions.IgnoreCase)) return value;
        var rows = Regex.Split(value, @"\s*,\s*\\quad\s*|(?:\r\n|\r|\n)+")
            .Select(row => row.Trim().Trim(',').Trim())
            .Where(row => row.Length > 0)
            .ToArray();
        if (rows.Length < 2) return value;
        return "\\begin{aligned}" + String.Join("\\\\", rows) + "\\end{aligned}";
    }

    private static void RemoveCopiedFormulaDelimiters(dynamic document)
    {
        foreach (var marker in new[] { "\\[", "\\]" })
        {
            try
            {
                dynamic range = document.Content.Duplicate;
                dynamic find = range.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                find.Text = marker;
                find.Replacement.Text = String.Empty;
                find.Forward = true;
                find.Wrap = 1;
                find.Format = false;
                find.MatchWildcards = false;
                ExecuteReplaceAll(find);
            }
            catch { }
        }
    }

    public int ConvertDisplayLatexToOfficeMath()
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic content = document.Content;
        var source = (string)content.Text;
        var matches = Regex.Matches(source, @"\$\$(.+?)\$\$|\\\[(.+?)\\\]", RegexOptions.Singleline);
        for (var index = matches.Count - 1; index >= 0; index--)
        {
            var match = matches[index];
            ReplaceWithOfficeMath(document, content.Start + match.Index, match.Length, match.Groups[1].Success ? match.Groups[1].Value : match.Groups[2].Value, true);
        }
        return matches.Count;
    }

    public int NormalizeMarkdownText()
    {
        return RenderAiWebLayout(true);
    }

    public int NormalizeMarkdownText(bool removeHeadings)
    {
        return RenderAiWebLayout(!removeHeadings);
    }

    public int RenderAiWebLayout()
    {
        return RenderAiWebLayout(true);
    }

    public int RenderAiWebLayout(bool preserveHeadings)
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic selection = ((dynamic)_application).Selection;
        dynamic selected = selection.Range.Duplicate;
        var selectedText = selected.Text as string;
        dynamic target = !String.IsNullOrWhiteSpace((selectedText ?? String.Empty).Trim('\r', '\a')) ? selected : document.Content;
        var text = target.Text as string ?? String.Empty;
        if (String.IsNullOrWhiteSpace(text.Trim('\r', '\a'))) return 0;
        RenderMarkdownRangeAsWeb(document, target, preserveHeadings);
        var formulas = ConvertLatexToOfficeMath();
        FinishAiWebLayout(document, (int)target.Start, (int)target.End);
        return formulas + 1;
    }

    public int RestoreAiWebFormulaLayout()
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic selection = ((dynamic)_application).Selection;
        dynamic selected = selection.Range.Duplicate;
        var selectedText = selected.Text as string;
        dynamic target = !String.IsNullOrWhiteSpace((selectedText ?? String.Empty).Trim('\r', '\a')) ? selected : document.Content;
        var source = target.Text as string ?? String.Empty;
        if (String.IsNullOrWhiteSpace(source.Trim('\r', '\a'))) return 0;

        var normalized = NormalizeWebAiFormulaText(source);
        if (!String.Equals(source, normalized, StringComparison.Ordinal))
        {
            var start = (int)target.Start;
            target.Text = normalized;
            target = document.Range(start, start + normalized.Length);
        }
        var count = ConvertLatexToOfficeMath();
        FinishAiWebLayout(document, (int)target.Start, (int)target.End);
        return count;
    }

    private static string NormalizeWebAiFormulaText(string text)
    {
        var value = text ?? String.Empty;
        value = value.Replace("\\\\(", "\\(").Replace("\\\\)", "\\)").Replace("\\\\[", "\\[").Replace("\\\\]", "\\]");
        value = value.Replace("反斜杠quad", "\\quad").Replace("反斜杠qquad", "\\qquad");
        value = Regex.Replace(
            value,
            @"\\\\(?=(?:begin|end|frac|dfrac|tfrac|binom|sqrt|text|textcolor|color|mathrm|mathbf|mathit|mathnormal|mathbb|mathcal|mathfrak|boldsymbol|bm|operatorname|overset|underset|overbrace|underbrace|substack|cancel|bcancel|xcancel|boxed|left|right|sin|cos|tan|cot|sec|csc|arcsin|arccos|arctan|lim|limsup|liminf|log|ln|exp|Pr|alpha|beta|gamma|delta|epsilon|theta|lambda|mu|nu|xi|pi|rho|sigma|tau|phi|varphi|omega|sum|prod|int|quad|qquad|cdot|times|leqslant|geqslant|leq|geq|neq|approx|equiv|infty|pm|mp|div|in|notin|subset|supset|cup|cap|forall|exists|partial|nabla|Longrightarrow|Longleftarrow|Longleftrightarrow|Rightarrow|Leftarrow|Leftrightarrow|longrightarrow|longleftarrow|longleftrightarrow|rightarrow|leftarrow|leftrightarrow|implies|iff|mapsto|to|therefore|because|perp|parallel)\b)",
            "\\");

        // ChatGPT/GPT 网页复制有时会把 \[...\] 的反斜杠丢掉，并把等号
        // 复制成一整行 ===。先恢复为标准 LaTeX 定界符，随后 Markdown 与
        // Word 公式转换即可沿用同一条稳定路径。
        value = Regex.Replace(
            value,
            @"(?ms)^[ \t]*\[[ \t]*(?:\r\n|\r|\n)(?<body>.*?)(?:\r\n|\r|\n)[ \t]*\][ \t]*$",
            match => "\\[\r" + NormalizeGptDisplayFormula(match.Groups["body"].Value) + "\r\\]");
        return value;
    }

    private static string NormalizeGptDisplayFormula(string body)
    {
        var lines = Regex.Split(body ?? String.Empty, @"\r\n|\r|\n")
            .Select(line => line.Trim())
            .Where(line => line.Length > 0)
            .ToArray();
        if (lines.Length == 0) return String.Empty;

        var result = new StringBuilder();
        var equalityChain = false;
        foreach (var rawLine in lines)
        {
            if (Regex.IsMatch(rawLine, @"^={3,}$"))
            {
                if (result.Length > 0 && result[result.Length - 1] != '=') result.Append('=');
                continue;
            }

            var line = rawLine;
            var hashTerm = Regex.Match(line, @"^#+\s*(?<term>.+)$");
            if (hashTerm.Success)
            {
                line = hashTerm.Groups["term"].Value.Trim();
                if (result.Length > 0 && result[result.Length - 1] != '=') result.Append('=');
                equalityChain = true;
            }
            else if (result.Length > 0 && result[result.Length - 1] != '=')
            {
                if (equalityChain) result.Append('=');
                else result.Append(' ');
            }
            result.Append(line);
        }
        return Regex.Replace(result.ToString(), @"\s+", " ").Trim();
    }

    private static dynamic RenderMarkdownRangeAsWeb(dynamic document, dynamic sourceRange, bool preserveHeadings)
    {
        var source = sourceRange.Text as string ?? String.Empty;
        var html = AiWebFormatter.ToHtml(source, preserveHeadings);
        var tempPath = Path.Combine(Path.GetTempPath(), "ai-word-layout-" + Guid.NewGuid().ToString("N") + ".html");
        File.WriteAllText(tempPath, html, new UTF8Encoding(true));
        try
        {
            var start = (int)sourceRange.Start;
            sourceRange.Text = String.Empty;
            dynamic insertion = document.Range(start, start);
            var beforeEnd = (int)document.Content.End;
            try
            {
                try { insertion.InsertFile(tempPath, Type.Missing, false, false, false); }
                catch { insertion.InsertFile(tempPath); }
                var insertedLength = Math.Max(0, (int)document.Content.End - beforeEnd);
                if (insertedLength == 0) throw new InvalidOperationException("网页排版内容未写入文档。");
                dynamic inserted = document.Range(start, Math.Min((int)document.Content.End, start + insertedLength));
                try { inserted.ListFormat.RemoveNumbers(); } catch { }
                return inserted;
            }
            catch
            {
                // InsertFile 在旧版 WPS 中可能先写入一部分内容再抛出异常。
                // 先删除这段增量，再原样恢复文本，避免一键规范清空文档。
                try
                {
                    var delta = Math.Max(0, (int)document.Content.End - beforeEnd);
                    if (delta > 0) document.Range(start, Math.Min((int)document.Content.End, start + delta)).Delete();
                }
                catch { }
                dynamic restored = document.Range(start, start);
                restored.Text = source;
                return ApplyPlainMarkdownFormatting(document, document.Range(start, (int)restored.End));
            }
        }
        finally
        {
            try { File.Delete(tempPath); } catch { }
        }
    }

    private static dynamic ApplyPlainMarkdownFormatting(dynamic document, dynamic sourceRange)
    {
        if (RangeContainsOfficeMath(sourceRange))
        {
            NormalizeMarkdownInPlace(document, sourceRange);
            return sourceRange;
        }
        var start = (int)sourceRange.Start;
        var source = sourceRange.Text as string ?? String.Empty;
        source = Regex.Replace(source, @"\*\*(?<body>.+?)\*\*", "${body}");
        source = Regex.Replace(source, @"__(?<body>.+?)__", "${body}");
        sourceRange.Text = source;
        dynamic formatted = document.Range(start, (int)sourceRange.End);
        try
        {
            var paragraphCount = SafeCollectionCount(formatted.Paragraphs);
            for (var paragraphIndex = 1; paragraphIndex <= paragraphCount; paragraphIndex++)
            {
                dynamic paragraph;
                try { paragraph = formatted.Paragraphs[paragraphIndex]; } catch { continue; }
                dynamic paragraphRange = paragraph.Range;
                var raw = paragraphRange.Text as string ?? String.Empty;
                var body = raw.TrimEnd('\r', '\a');
                var heading = Regex.Match(body, @"^\s*(?<marks>#{1,6})\s+(?<text>.+)$");
                if (heading.Success)
                {
                    dynamic writable = paragraphRange.Duplicate;
                    try { writable.End = writable.End - 1; } catch { }
                    writable.Text = heading.Groups["text"].Value.Trim();
                    paragraphRange.Font.Bold = 1;
                    paragraphRange.Font.NameFarEast = "微软雅黑";
                    paragraphRange.Font.Size = Math.Max(12f, 19f - heading.Groups["marks"].Length * 1.5f);
                    paragraphRange.ParagraphFormat.SpaceBefore = 6f;
                    paragraphRange.ParagraphFormat.SpaceAfter = 4f;
                    continue;
                }
                var bullet = Regex.Match(body, @"^\s*[-+*]\s+(?<text>.+)$");
                if (bullet.Success)
                {
                    dynamic writable = paragraphRange.Duplicate;
                    try { writable.End = writable.End - 1; } catch { }
                    writable.Text = "• " + bullet.Groups["text"].Value;
                }
            }
        }
        catch { }
        return formatted;
    }

    private static int NormalizeMarkdownInPlace(dynamic document, dynamic sourceRange)
    {
        var changed = 0;
        var paragraphCount = SafeCollectionCount(sourceRange.Paragraphs);
        var fencedParagraphs = new HashSet<int>();
        var inFence = false;
        for (var scanIndex = 1; scanIndex <= paragraphCount; scanIndex++)
        {
            try
            {
                var scanText = Convert.ToString(sourceRange.Paragraphs[scanIndex].Range.Text).Trim('\r', '\n', '\a', ' ', '\t');
                if (scanText.StartsWith("```", StringComparison.Ordinal))
                {
                    fencedParagraphs.Add(scanIndex);
                    inFence = !inFence;
                }
                else if (inFence) fencedParagraphs.Add(scanIndex);
            }
            catch { }
        }
        // 倒序删除前缀，避免前面段落的长度变化影响后续坐标。
        for (var paragraphIndex = paragraphCount; paragraphIndex >= 1; paragraphIndex--)
        {
            if (fencedParagraphs.Contains(paragraphIndex)) continue;
            dynamic paragraph;
            try { paragraph = sourceRange.Paragraphs[paragraphIndex]; } catch { continue; }
            dynamic paragraphRange = paragraph.Range;
            var raw = paragraphRange.Text as string ?? String.Empty;
            var body = raw.TrimEnd('\r', '\a');
            changed += FormatInlineMarkdownMarkers(document, paragraphRange, body);
            raw = paragraphRange.Text as string ?? String.Empty;
            body = raw.TrimEnd('\r', '\a');
            var heading = Regex.Match(body, @"^(?<prefix>\s*#{1,6}\s+)(?<text>.+)$");
            if (heading.Success)
            {
                try
                {
                    var marks = Regex.Match(heading.Groups["prefix"].Value, @"#{1,6}").Value.Length;
                    var start = (int)paragraphRange.Start;
                    document.Range(start, start + heading.Groups["prefix"].Length).Delete();
                    dynamic updated = document.Range(start, start).Paragraphs[1].Range;
                    updated.Font.Bold = 1;
                    updated.Font.NameFarEast = "微软雅黑";
                    updated.Font.Size = Math.Max(12f, 19f - marks * 1.5f);
                    updated.ParagraphFormat.SpaceBefore = 6f;
                    updated.ParagraphFormat.SpaceAfter = 4f;
                    updated.ParagraphFormat.KeepWithNext = -1;
                    changed++;
                }
                catch { }
                continue;
            }

            var bullet = Regex.Match(body, @"^(?<prefix>\s*[-+*]\s+)(?<text>.+)$");
            if (bullet.Success)
            {
                try
                {
                    var start = (int)paragraphRange.Start;
                    document.Range(start, start + bullet.Groups["prefix"].Length).Text = "• ";
                    changed++;
                }
                catch { }
            }
            var quote = Regex.Match(body, @"^(?<prefix>\s*>\s+)(?<text>.+)$");
            if (quote.Success)
            {
                try
                {
                    var start = (int)paragraphRange.Start;
                    document.Range(start, start + quote.Groups["prefix"].Length).Delete();
                    document.Range(start, start).Paragraphs[1].Range.ParagraphFormat.LeftIndent = 18f;
                    changed++;
                }
                catch { }
            }
        }
        return changed;
    }

    private static int FormatInlineMarkdownMarkers(dynamic document, dynamic paragraphRange, string body)
    {
        var matches = Regex.Matches(body ?? String.Empty, @"(?<mark>\*\*|__)(?<text>.+?)\k<mark>");
        var changed = 0;
        for (var index = matches.Count - 1; index >= 0; index--)
        {
            var match = matches[index];
            try
            {
                var start = (int)paragraphRange.Start + match.Index;
                var markerLength = match.Groups["mark"].Length;
                document.Range(start + match.Length - markerLength, start + match.Length).Delete();
                document.Range(start, start + markerLength).Delete();
                dynamic content = document.Range(start, start + match.Groups["text"].Length);
                content.Font.Bold = 1;
                changed++;
            }
            catch { }
        }
        return changed;
    }

    private static void FinishAiWebLayout(dynamic document)
    {
        FinishAiWebLayout(document, (int)document.Content.Start, (int)document.Content.End);
    }

    private static void FinishAiWebLayout(dynamic document, int scopeStart, int scopeEnd)
    {
        try
        {
            var tableCount = SafeCollectionCount(document.Tables);
            for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
            {
                try
                {
                    dynamic table = document.Tables[tableIndex];
                    if ((int)table.Range.End <= scopeStart || (int)table.Range.Start >= scopeEnd) continue;
                    table.AllowAutoFit = true;
                    table.AutoFitBehavior(2);
                    table.Range.ParagraphFormat.SpaceAfter = 0f;
                    table.Range.ParagraphFormat.LineSpacingRule = 0;
                    table.Rows[1].Range.Font.Bold = 1;
                    table.Rows[1].Shading.BackgroundPatternColor = ParseWordColor("F3F4F6");
                    table.Borders.Enable = 1;
                    table.Borders.OutsideColor = ParseWordColor("D0D7DE");
                    table.Borders.InsideColor = ParseWordColor("D0D7DE");
                }
                catch { }
            }
        }
        catch { }

        try
        {
            dynamic equations = document.OMaths;
            var equationCount = (int)equations.Count;
            for (var equationIndex = 1; equationIndex <= equationCount; equationIndex++)
            {
                try
                {
                    dynamic math = equations[equationIndex];
                    if (!MathRangesOverlap(math, scopeStart, scopeEnd)) continue;
                    math.Range.Font.Name = "Cambria Math";
                    math.Range.Font.Position = 0;
                    dynamic paragraphRange = math.Range.Paragraphs[1].Range;
                    var paragraphText = ((string)paragraphRange.Text).Trim('\r', '\a', ' ', '\t');
                    var mathText = ((string)math.Range.Text).Trim('\r', '\a', ' ', '\t');
                    if (String.Equals(paragraphText, mathText, StringComparison.Ordinal))
                    {
                        paragraphRange.ParagraphFormat.Alignment = 1;
                        paragraphRange.ParagraphFormat.SpaceBefore = 6f;
                        paragraphRange.ParagraphFormat.SpaceAfter = 8f;
                        paragraphRange.ParagraphFormat.KeepTogether = -1;
                    }
                }
                catch { }
            }
        }
        catch { }
    }

    public int RestoreMarkdownTables()
    {
        var settings = NormalizationSettings.Load();
        return RestoreMarkdownTables(settings.ThreeLineTables, settings.AutoFitTables, false);
    }

    public int RestoreMarkdownTables(bool threeLine, bool autoFit, bool deepseekStyle)
    {
        return RestoreMarkdownTables(threeLine, autoFit, deepseekStyle, "Document");
    }

    public int RestoreMarkdownTables(bool threeLine, bool autoFit, bool deepseekStyle, string scope)
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic content = document.Content;
        if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase))
        {
            dynamic selected = ((dynamic)_application).Selection.Range;
            if ((int)selected.End <= (int)selected.Start) throw new InvalidOperationException("请先选中需要处理的内容。");
            content = selected.Duplicate;
        }
        var text = (string)content.Text;
        var blocks = Regex.Matches(text, @"(?m)^(?:\|.*\|\r?$\n?){2,}");
        for (var blockIndex = blocks.Count - 1; blockIndex >= 0; blockIndex--)
        {
            var block = blocks[blockIndex];
            var rows = block.Value.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                .Where(line => !Regex.IsMatch(line, @"^\|?\s*:?-{3,}"))
                .Select(line => line.Trim().Trim('|').Split('|').Select(cell => Regex.Replace(cell.Trim(), @"\s*<br\s*/?>\s*", "\r", RegexOptions.IgnoreCase)).ToArray()).ToList();
            if (rows.Count == 0) continue;
            var columns = rows.Max(row => row.Length);
            dynamic tableRange = document.Range(content.Start + block.Index, content.Start + block.Index + block.Length);
            tableRange.Text = String.Empty;
            try
            {
                dynamic table = document.Tables.Add(tableRange, rows.Count, columns);
                for (var row = 0; row < rows.Count; row++) for (var column = 0; column < rows[row].Length; column++) table.Cell(row + 1, column + 1).Range.Text = rows[row][column];
                if (autoFit) try { table.AutoFitBehavior(2); } catch { }
                if (threeLine) ApplyThreeLineStyle(table);
                if (deepseekStyle) try { table.Rows[1].Range.Font.Bold = 1; } catch { }
            }
            catch { tableRange.Text = String.Join("\t", rows.Select(row => String.Join("\t", row))) + Environment.NewLine; }
        }
        SelectScopeRange(content, scope);
        return blocks.Count;
    }

    public int ClearTableSpaces()
    {
        return ClearTableSpaces("Document");
    }

    public int ClearTableSpaces(string scope)
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        var scopeStart = (int)document.Content.Start;
        var scopeEnd = (int)document.Content.End;
        if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase))
        {
            dynamic selected = ((dynamic)_application).Selection.Range;
            scopeStart = (int)selected.Start;
            scopeEnd = (int)selected.End;
            if (scopeEnd <= scopeStart) throw new InvalidOperationException("请先选中需要处理的内容。");
        }
        return ClearTableSpaces(document, scopeStart, scopeEnd);
    }

    private static int ClearTableSpaces(dynamic document, int scopeStart, int scopeEnd)
    {
        var count = 0;
        var tableCount = SafeCollectionCount(document.Tables);
        for (var tableIndex = tableCount; tableIndex >= 1; tableIndex--)
        {
            dynamic table;
            try { table = document.Tables[tableIndex]; } catch { continue; }
            try { if ((int)table.Range.End <= scopeStart || (int)table.Range.Start >= scopeEnd) continue; } catch { }
            var cellCount = SafeCollectionCount(table.Range.Cells);
            for (var cellIndex = cellCount; cellIndex >= 1; cellIndex--)
            {
                try
                {
                    dynamic cell = table.Range.Cells[cellIndex];
                    if ((int)cell.Range.End <= scopeStart || (int)cell.Range.Start >= scopeEnd) continue;
                    var text = (string)cell.Range.Text;
                    text = text.TrimEnd('\r', '\a');
                    text = Regex.Replace(text, @"[ \t\u3000]+", " ").Trim();
                    dynamic writable = cell.Range.Duplicate;
                    writable.End = writable.End - 1;
                    writable.Text = text;
                    count++;
                }
                catch { }
            }
        }
        return count;
    }

    public int SmartCleanTables(bool autoFit, bool threeLine)
    {
        return SmartCleanTables(autoFit, threeLine, "Document");
    }

    public int SmartCleanTables(bool autoFit, bool threeLine, string scope)
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic trackingRange = null;
        var scopeStart = (int)document.Content.Start;
        var scopeEnd = (int)document.Content.End;
        if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase))
        {
            dynamic selected = ((dynamic)_application).Selection.Range;
            if ((int)selected.End <= (int)selected.Start) throw new InvalidOperationException("请先选中需要处理的内容。");
            trackingRange = selected.Duplicate;
            scopeStart = (int)trackingRange.Start;
            scopeEnd = (int)trackingRange.End;
        }
        var count = ClearTableSpaces(document, scopeStart, scopeEnd);
        if (trackingRange != null)
        {
            scopeStart = (int)trackingRange.Start;
            scopeEnd = (int)trackingRange.End;
        }
        var tableCount = SafeCollectionCount(document.Tables);
        for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
        {
            dynamic table;
            try { table = document.Tables[tableIndex]; } catch { continue; }
            try { if ((int)table.Range.End <= scopeStart || (int)table.Range.Start >= scopeEnd) continue; } catch { }
            try { table.AllowAutoFit = autoFit; if (autoFit) table.AutoFitBehavior(2); } catch { }
            if (threeLine) ApplyThreeLineStyle(table);
        }
        SelectScopeRange(trackingRange, scope);
        return count;
    }

    private static void ApplyThreeLineStyle(dynamic table)
    {
        try
        {
            table.Borders.Enable = 0;
            table.Borders[-1].LineStyle = 1;
            table.Borders[-3].LineStyle = 1;
            table.Rows[1].Borders[-3].LineStyle = 1;
            table.Rows[1].Range.Font.Bold = 1;
        }
        catch { }
    }

    public void ApplyPreset(string fontName, float fontSize, int alignment, bool firstLineIndent, bool paragraphSpacing, float lineSpacing = 1.5F, float spaceAfter = 6F)
    {
        ApplyPreset(fontName, fontSize, alignment, firstLineIndent, paragraphSpacing, lineSpacing, spaceAfter, "Document");
    }

    public void ApplyPreset(string fontName, float fontSize, int alignment, bool firstLineIndent, bool paragraphSpacing, float lineSpacing, float spaceAfter, string scope)
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic range = document.Content;
        if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase))
        {
            dynamic selected = ((dynamic)_application).Selection.Range;
            if ((int)selected.End <= (int)selected.Start) throw new InvalidOperationException("请先选中需要排版的内容。");
            range = selected;
        }
        if (!String.IsNullOrWhiteSpace(fontName)) range.Font.Name = fontName;
        if (fontSize > 0) range.Font.Size = fontSize;
        range.ParagraphFormat.Alignment = alignment;
        range.ParagraphFormat.FirstLineIndent = firstLineIndent ? 24f : 0f;
        range.ParagraphFormat.SpaceAfter = paragraphSpacing ? spaceAfter : 0f;
        range.ParagraphFormat.LineSpacingRule = 5;
        try { range.ParagraphFormat.LineSpacing = Math.Max(1F, lineSpacing) * fontSize; } catch { }
    }

    public void InsertAnswerArea(NormalizationSettings settings)
    {
        if (settings == null) settings = NormalizationSettings.Load();
        dynamic selection = ((dynamic)_application).Selection;
        var lines = Math.Max(1, Math.Min(20, settings.AnswerLineCount));
        var length = Math.Max(8, Math.Min(80, settings.AnswerLineLength));
        var builder = new StringBuilder("\r");
        for (var index = 0; index < lines; index++) builder.Append(new string('＿', length)).Append("\r");
        selection.Collapse(0);
        selection.TypeText(builder.ToString());
    }

    public int SetAnswerSectionsHidden(bool hidden)
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        var changed = 0;
        foreach (var item in CollectAnswerSectionRanges(document))
        {
            try { document.Range(item.Item1, item.Item2).Font.Hidden = hidden ? 1 : 0; changed++; } catch { }
        }
        return changed;
    }

    public Tuple<string, int> CreateStudentCopy()
    {
        dynamic source = ((dynamic)_application).ActiveDocument;
        dynamic documents = ((dynamic)_application).Documents;
        var sourceName = Convert.ToString(source.Name);
        var baseName = Path.GetFileNameWithoutExtension(String.IsNullOrWhiteSpace(sourceName) ? "新建文档" : sourceName);
        string folder;
        try { folder = Convert.ToString(source.Path); } catch { folder = String.Empty; }
        if (String.IsNullOrWhiteSpace(folder)) folder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var targetPath = GetUniqueOutputPath(folder, baseName + "-学生版", ".docx");
        dynamic target = null;
        try
        {
            try
            {
                source.SaveCopyAs(targetPath);
                target = documents.Open(targetPath);
            }
            catch
            {
                target = documents.Add();
                target.Content.FormattedText = source.Content.FormattedText;
                try { target.SaveAs2(targetPath, 16); }
                catch { target.SaveAs(targetPath, 16); }
            }
            var ranges = CollectAnswerSectionRanges(target);
            for (var index = ranges.Count - 1; index >= 0; index--)
            {
                try { target.Range(ranges[index].Item1, ranges[index].Item2).Delete(); } catch { }
            }
            try { target.Save(); } catch { }
            return Tuple.Create(targetPath, ranges.Count);
        }
        catch
        {
            if (target != null) try { target.Close(false); } catch { }
            throw;
        }
    }

    public string GetDocumentCheckReport()
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        var text = document.Content.Text as string ?? String.Empty;
        var formulaCandidates = FindAiFormulaCandidates(text);
        var rawCommands = formulaCandidates.Count;
        var markdown = Regex.Matches(text, @"(?m)^\s{0,3}(?:#{1,6}\s+|[-+*]\s+|>\s+)|(?:\*\*|__)[^\r\n]+?(?:\*\*|__)").Count;
        var unmatched = CountUnmatchedFormulaDelimiters(text);
        var answers = CollectAnswerSectionRanges(document).Count;
        var lines = new List<string>
        {
            "文档：" + ActiveDocumentName,
            "页数：" + SafeDocumentStatistic(document, 2),
            "可编辑公式：" + SafeCollectionCount(document.OMaths),
            "待转换 TeX：" + rawCommands,
            "表格：" + SafeCollectionCount(document.Tables),
            "答案或解析段落：" + answers,
            "待整理标记：" + markdown,
            "公式定界符问题：" + unmatched,
            String.Empty,
            unmatched == 0 && rawCommands == 0 && markdown == 0
                ? "检查完成，未发现明显的排版残留。"
                : "建议先规范文本和公式，再重新检查。"
        };
        return String.Join(Environment.NewLine, lines);
    }

    private static List<Tuple<int, int>> CollectAnswerSectionRanges(dynamic document)
    {
        var ranges = new List<Tuple<int, int>>();
        var inAnswer = false;
        var paragraphCount = SafeCollectionCount(document.Paragraphs);
        for (var paragraphIndex = 1; paragraphIndex <= paragraphCount; paragraphIndex++)
        {
            dynamic paragraph;
            try { paragraph = document.Paragraphs[paragraphIndex]; } catch { continue; }
            dynamic range = paragraph.Range;
            var text = ((string)range.Text ?? String.Empty).Trim();
            var answerHeading = Regex.IsMatch(text, @"^(?:[#*\s]*)(?:【|\[)?(?:参考答案|答案(?:与解析)?|参考解答|解析|解答|评分标准)(?:】|\])?[：:]?", RegexOptions.IgnoreCase);
            var nextQuestion = Regex.IsMatch(text, @"^(?:[#*\s]*)(?:\d+[\.．、]|[一二三四五六七八九十百]+、|第[一二三四五六七八九十百\d]+部分)") && !answerHeading;
            if (nextQuestion) inAnswer = false;
            if (answerHeading) inAnswer = true;
            if (!inAnswer) continue;
            try { ranges.Add(Tuple.Create((int)range.Start, (int)range.End)); } catch { }
        }
        return ranges;
    }

    private static string GetUniqueOutputPath(string folder, string baseName, string extension)
    {
        Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, baseName + extension);
        for (var index = 2; File.Exists(path); index++) path = Path.Combine(folder, baseName + " (" + index + ")" + extension);
        return path;
    }

    private static int SafeDocumentStatistic(dynamic document, int statistic)
    {
        try { return (int)document.ComputeStatistics(statistic); } catch { return 0; }
    }

    private static int CountUnmatchedFormulaDelimiters(string text)
    {
        var count = 0;
        if (Regex.Matches(text ?? String.Empty, @"(?<!\\)\$\$").Count % 2 != 0) count++;
        var withoutDouble = Regex.Replace(text ?? String.Empty, @"(?<!\\)\$\$", String.Empty);
        if (Regex.Matches(withoutDouble, @"(?<!\\)\$").Count % 2 != 0) count++;
        if (Regex.Matches(text ?? String.Empty, @"\\\(").Count != Regex.Matches(text ?? String.Empty, @"\\\)").Count) count++;
        if (Regex.Matches(text ?? String.Empty, @"\\\[").Count != Regex.Matches(text ?? String.Empty, @"\\\]").Count) count++;
        return count;
    }

    public int ConvertOfficeMathToMarkdown()
    {
        return ConvertOfficeMathToMarkdown("Auto");
    }

    public int ConvertOfficeMathToMarkdown(string scope)
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        var scopeStart = (int)document.Content.Start;
        var scopeEnd = (int)document.Content.End;
        dynamic trackingRange = null;
        if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase) ||
            String.Equals(scope, "Auto", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                dynamic selection = ((dynamic)_application).Selection;
                dynamic selected = selection.Range;
                if ((int)selected.End > (int)selected.Start)
                {
                    trackingRange = selected.Duplicate;
                    scopeStart = (int)trackingRange.Start;
                    scopeEnd = (int)trackingRange.End;
                }
                else if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("请先选中需要转换的内容。");
            }
            catch (InvalidOperationException) { throw; }
            catch { if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("请先选中需要转换的内容。"); }
        }
        var count = 0;
        for (var index = SafeCollectionCount(document.OMaths); index >= 1; index--)
        {
            try
            {
                dynamic math = document.OMaths[index];
                dynamic range = math.Range;
                if (!MathRangesOverlap(math, scopeStart, scopeEnd)) continue;
                var display = IsStandaloneOfficeMath(document, range);
                var text = String.Empty;
                try { text = OmmlToLatexConverter.Convert(Convert.ToString(range.WordOpenXML)); } catch { }
                if (!IsValidLatexConversion(text))
                {
                    var fallbackOriginal = range.Text as string ?? String.Empty;
                    try
                    {
                        math.Linearize();
                        range = math.Range;
                        text = OmmlToLatexConverter.ConvertLinear(range.Text as string ?? String.Empty);
                    }
                    catch { text = String.Empty; }
                    if (!IsValidLatexConversion(text))
                    {
                        try { range.Text = fallbackOriginal; TryBuildMath(math); } catch { }
                        continue;
                    }
                }
                var marker = display ? "$$" + text + "$$" : "$" + text + "$";
                var original = range.Text as string ?? String.Empty;
                var markerStart = (int)range.Start;
                var convertedToText = false;
                try
                {
                    // 先把标准 TeX 写入公式对象，再调用 Remove。Word 会把公式内部的
                    // 数学字母还原为普通 ASCII 文本，同时移除公式框；反过来先 Remove
                    // 会导致原线性公式残留并与 TeX 文本重复。
                    range.Text = marker;
                    math.Remove();
                    dynamic plain = document.Range(markerStart,
                        Math.Min((int)document.Content.End, markerStart + marker.Length));
                    convertedToText = String.Equals(plain.Text as string ?? String.Empty, marker, StringComparison.Ordinal);
                }
                catch { }
                if (!convertedToText)
                {
                    try
                    {
                        range.Text = marker;
                        math.ConvertToNormalText();
                        dynamic plain = document.Range(markerStart,
                            Math.Min((int)document.Content.End, markerStart + marker.Length));
                        convertedToText = String.Equals(plain.Text as string ?? String.Empty, marker, StringComparison.Ordinal);
                    }
                    catch { }
                }
                if (!convertedToText)
                {
                    try { range.Text = original; TryBuildMath(math); } catch { }
                    continue;
                }
                count++;
            }
            catch { }
        }
        SelectScopeRange(trackingRange, trackingRange == null ? "Document" : "Selection");
        return count;
    }

    private static bool IsStandaloneOfficeMath(dynamic document, dynamic mathRange)
    {
        try
        {
            dynamic paragraph = mathRange.Paragraphs[1].Range;
            var before = document.Range((int)paragraph.Start, (int)mathRange.Start).Text as string ?? String.Empty;
            var after = document.Range((int)mathRange.End, (int)paragraph.End).Text as string ?? String.Empty;
            return String.IsNullOrWhiteSpace(before.Trim('\r', '\n', '\a')) &&
                   String.IsNullOrWhiteSpace(after.Trim('\r', '\n', '\a'));
        }
        catch { return false; }
    }

    private static bool IsValidLatexConversion(string value)
    {
        if (String.IsNullOrWhiteSpace(value)) return false;
        if (Regex.IsMatch(value, @"[■█┴┬├┤√]")) return false;
        if (Regex.IsMatch(value, @"\\(?:matrix|eqarray|cases|sqrt)\s*\(", RegexOptions.IgnoreCase)) return false;
        var depth = 0;
        for (var index = 0; index < value.Length; index++)
        {
            if (value[index] == '\\') { index++; continue; }
            if (value[index] == '{') depth++;
            else if (value[index] == '}' && --depth < 0) return false;
        }
        return depth == 0;
    }

    public int ConvertSelectedTextToOfficeMath()
    {
        return ConvertTextToOfficeMath("Auto");
    }

    public int ConvertTextToOfficeMath(string scope)
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic selection = ((dynamic)_application).Selection;
        dynamic range = selection.Range.Duplicate;
        var sourceText = range.Text as string ?? String.Empty;
        var text = sourceText.Trim();
        if (String.Equals(scope, "Document", StringComparison.OrdinalIgnoreCase)) return ConvertLatexToOfficeMath();
        if (String.IsNullOrWhiteSpace(text))
        {
            if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("请先选中需要转换的 TeX。");
            return ConvertLatexToOfficeMath();
        }

        // 选中“正文 + 多个公式”时逐个转换，保留正文与各自的行内/独立公式属性。
        // 只有选区本身是裸 TeX 时，才将整个选区作为一个公式处理。
        var candidates = FindAiFormulaCandidates(sourceText);
        if (candidates.Count > 0)
        {
            var rangeStart = (int)range.Start;
            var converted = 0;
            foreach (var candidate in candidates.OrderByDescending(item => item.Index))
            {
                if (ReplaceWithOfficeMath(document, rangeStart + candidate.Index, candidate.Length,
                    candidate.Display ? WrapLongFormulaForDisplay(candidate.Formula) : candidate.Formula,
                    candidate.Display)) converted++;
            }
            SelectScopeRange(range, "Selection");
            return converted;
        }
        if (!IsSelectedBareTex(text)) return 0;
        text = Regex.Replace(text, @"^\s*(?:\$\$|\$|\\\[|\\\()", String.Empty);
        text = Regex.Replace(text, @"(?:\$\$|\$|\\\]|\\\))\s*$", String.Empty);
        var result = ReplaceWithOfficeMath(document, (int)range.Start, (int)(range.End - range.Start), text, false) ? 1 : 0;
        SelectScopeRange(range, "Selection");
        return result;
    }

    private static bool IsSelectedBareTex(string value)
    {
        var text = (value ?? String.Empty).Trim();
        if (text.Length == 0 || Regex.IsMatch(text, @"\r|\n")) return false;
        var outsideTextCommands = Regex.Replace(text, @"\\(?:text|textrm|textnormal|mathrm|operatorname)\s*\{[^{}]*\}", String.Empty);
        if (Regex.IsMatch(outsideTextCommands, @"[\u3400-\u9fff]")) return false;
        return IsReliableBareFormula(text, true, false);
    }

    public int FloatEquations()
    {
        return FloatEquations("Document");
    }

    public int FloatEquations(string scope)
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        var count = 0;
        var scopeStart = (int)document.Content.Start;
        var scopeEnd = (int)document.Content.End;
        if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase))
        {
            dynamic selected = ((dynamic)_application).Selection.Range;
            scopeStart = (int)selected.Start; scopeEnd = (int)selected.End;
            if (scopeEnd <= scopeStart) throw new InvalidOperationException("请先选中需要调整的公式。");
        }
        for (var index = SafeCollectionCount(document.OMaths); index >= 1; index--)
        {
            try { dynamic math = document.OMaths[index]; if (!MathRangesOverlap(math, scopeStart, scopeEnd)) continue; math.Range.Font.Position = 3; count++; }
            catch { }
        }
        return count;
    }

    public int ResetEquationPosition()
    {
        return ResetEquationPosition("Document");
    }

    public int ResetEquationPosition(string scope)
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        var count = 0;
        var scopeStart = (int)document.Content.Start;
        var scopeEnd = (int)document.Content.End;
        if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase))
        {
            dynamic selected = ((dynamic)_application).Selection.Range;
            scopeStart = (int)selected.Start; scopeEnd = (int)selected.End;
            if (scopeEnd <= scopeStart) throw new InvalidOperationException("请先选中需要调整的公式。");
        }
        for (var index = SafeCollectionCount(document.OMaths); index >= 1; index--)
        {
            try { dynamic math = document.OMaths[index]; if (!MathRangesOverlap(math, scopeStart, scopeEnd)) continue; math.Range.Font.Position = 0; count++; }
            catch { }
        }
        return count;
    }

    public int ConvertMathTypeToOfficeMath()
    {
        return ConvertMathTypeToOfficeMath("Auto");
    }

    public int ConvertMathTypeToOfficeMath(string scope)
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        var originalStart = -1;
        var originalEnd = -1;
        dynamic scopedRange = null;
        var scopeStart = (int)document.Content.Start;
        var scopeEnd = (int)document.Content.End;
        var selectionScope = false;
        try
        {
            dynamic selected = ((dynamic)_application).Selection.Range;
            originalStart = (int)selected.Start;
            originalEnd = (int)selected.End;
            if ((int)selected.End > (int)selected.Start)
            {
                if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase) ||
                    String.Equals(scope, "Auto", StringComparison.OrdinalIgnoreCase))
                {
                    selectionScope = true;
                    scopedRange = selected.Duplicate;
                    scopeStart = (int)scopedRange.Start;
                    scopeEnd = (int)scopedRange.End;
                }
            }
            else if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("请先选中需要转换的 MathType 公式。");
        }
        catch (InvalidOperationException) { throw; }
        catch { if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("请先选中需要转换的 MathType 公式。"); }

        try
        {
            // 浮动 MathType/OLE 公式先转换为嵌入式对象，再沿用同一条转换路径。
            for (var index = SafeCollectionCount(document.Shapes); index >= 1; index--)
            {
                try
                {
                    dynamic floating = document.Shapes[index];
                    var programId = Convert.ToString(floating.OLEFormat.ProgID);
                    if (!IsEquationOleProgramId(programId)) continue;
                    dynamic anchor = floating.Anchor;
                    if ((int)anchor.Start >= scopeEnd || (int)anchor.End <= scopeStart) continue;
                    floating.ConvertToInlineShape();
                }
                catch { }
            }

            if (scopedRange != null)
            {
                scopeStart = (int)scopedRange.Start;
                scopeEnd = (int)scopedRange.End;
            }

            var count = 0;
            for (var index = SafeCollectionCount(document.InlineShapes); index >= 1; index--)
            {
                dynamic shape;
                try { shape = document.InlineShapes[index]; } catch { continue; }
                string programId;
                try { programId = Convert.ToString(shape.OLEFormat.ProgID); } catch { continue; }
                if (!IsEquationOleProgramId(programId)) continue;

                dynamic range = shape.Range;
                if ((int)range.Start >= scopeEnd || (int)range.End <= scopeStart) continue;
                var beforeMath = SafeCollectionCount(document.OMaths);
                try
                {
                    range.Select();
                    if ((TryExecuteMso("MathConvertToOfficeMath") || TryExecuteMso("EquationConvertToProfessional")) &&
                        SafeCollectionCount(document.OMaths) > beforeMath)
                    {
                        count++;
                        continue;
                    }
                }
                catch { }

                var source = ReadEquationOleText(shape);
                if (String.IsNullOrWhiteSpace(source)) continue;
                try
                {
                    source = source.Trim().Trim('$');
                    if (ReplaceWithOfficeMath(document, (int)range.Start, (int)(range.End - range.Start), source, false)) count++;
                }
                catch { }
            }
            return count;
        }
        finally
        {
            if (selectionScope && scopedRange != null) SelectScopeRange(scopedRange, "Selection");
            else RestoreSelection(document, originalStart, originalEnd);
        }
    }

    public int ConvertOfficeMathToMathType()
    {
        return ConvertOfficeMathToMathType("Auto");
    }

    public int ConvertOfficeMathToMathType(string scope)
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        var originalStart = -1;
        var originalEnd = -1;
        dynamic target = document.Content;
        var selectionScope = false;
        try
        {
            dynamic selected = ((dynamic)_application).Selection.Range;
            originalStart = (int)selected.Start;
            originalEnd = (int)selected.End;
            if ((int)selected.End > (int)selected.Start &&
                (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase) ||
                 String.Equals(scope, "Auto", StringComparison.OrdinalIgnoreCase)))
            {
                selectionScope = true;
                target = selected.Duplicate;
            }
            else if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("请先选中需要转换的 Office 公式。");
        }
        catch (InvalidOperationException) { throw; }
        catch { if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("请先选中需要转换的 Office 公式。"); }

        var targetMath = GetOfficeMathSpans(document, (int)target.Start, (int)target.End).Count;
        if (targetMath == 0) return -2;
        try
        {
            var beforeMath = SafeCollectionCount(document.OMaths);
            var beforeShapes = SafeCollectionCount(document.InlineShapes) + SafeCollectionCount(document.Shapes);
            target.Select();
            if (!TryExecuteMso("MathConvertToMathType") && !TryRunMathTypeConvertCommand()) return -1;
            var remaining = SafeCollectionCount(document.OMaths);
            var shapes = SafeCollectionCount(document.InlineShapes) + SafeCollectionCount(document.Shapes);
            return Math.Max(beforeMath - remaining, Math.Max(0, shapes - beforeShapes));
        }
        finally
        {
            if (selectionScope) SelectScopeRange(target, "Selection");
            else RestoreSelection(document, originalStart, originalEnd);
        }
    }

    private static void RestoreSelection(dynamic document, int start, int end)
    {
        if (start < 0) return;
        try
        {
            var documentEnd = (int)document.Content.End;
            start = Math.Min(Math.Max((int)document.Content.Start, start), documentEnd);
            end = Math.Min(Math.Max(start, end), documentEnd);
            document.Range(start, end).Select();
        }
        catch { }
    }

    private bool TryExecuteMso(string controlId)
    {
        try
        {
            ((dynamic)_application).CommandBars.ExecuteMso(controlId);
            return true;
        }
        catch { return false; }
    }

    private bool TryRunMathTypeConvertCommand()
    {
        var macros = new[]
        {
            "MathTypeCommands.UILib.MTCommand_ConvertEquations",
            "MathTypeCommands2016.UILib.MTCommand_ConvertEquations",
            "MTCommand_ConvertEquations"
        };
        foreach (var macro in macros)
        {
            try { ((dynamic)_application).Run(macro); return true; }
            catch { }
        }
        return false;
    }

    private static int SafeCollectionCount(dynamic collection)
    {
        try { return (int)collection.Count; }
        catch { return 0; }
    }

    private static bool IsEquationOleProgramId(string value)
    {
        return !String.IsNullOrWhiteSpace(value) &&
               (value.IndexOf("Equation", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("MathType", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("DSMT", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    private static string ReadEquationOleText(dynamic shape)
    {
        string value;
        try { value = Convert.ToString(shape.AlternativeText); if (!String.IsNullOrWhiteSpace(value)) return value; } catch { }
        try { value = Convert.ToString(shape.Title); if (!String.IsNullOrWhiteSpace(value)) return value; } catch { }
        dynamic equation;
        try { equation = shape.OLEFormat.Object; } catch { return String.Empty; }
        try { value = Convert.ToString(equation.LaTeX); if (!String.IsNullOrWhiteSpace(value)) return value; } catch { }
        try { value = Convert.ToString(equation.Latex); if (!String.IsNullOrWhiteSpace(value)) return value; } catch { }
        try { value = Convert.ToString(equation.Text); if (!String.IsNullOrWhiteSpace(value)) return value; } catch { }
        try { value = Convert.ToString(equation.EquationString); if (!String.IsNullOrWhiteSpace(value)) return value; } catch { }
        return String.Empty;
    }

    private static string ToLatexText(string value)
    {
        var replacements = new[] { new[] { "×", "\\times" }, new[] { "≤", "\\leq" }, new[] { "≥", "\\geq" }, new[] { "≠", "\\neq" }, new[] { "∞", "\\infty" }, new[] { "±", "\\pm" }, new[] { "π", "\\pi" }, new[] { "α", "\\alpha" }, new[] { "β", "\\beta" }, new[] { "γ", "\\gamma" }, new[] { "θ", "\\theta" }, new[] { "∑", "\\sum" }, new[] { "∏", "\\prod" }, new[] { "∫", "\\int" } };
        foreach (var item in replacements) value = value.Replace(item[0], item[1]);
        return value;
    }

    private static string ToLinearOfficeMath(string latex)
    {
        var value = latex.Trim();
        value = Regex.Replace(value, @"\\frac\s*\{([^{}]+)\}\s*\{([^{}]+)\}", "($1)/($2)");
        value = Regex.Replace(value, @"\\sqrt\s*\{([^{}]+)\}", "sqrt($1)");
        value = value.Replace("\\times", "×").Replace("\\cdot", "·").Replace("\\leq", "≤").Replace("\\geq", "≥").Replace("\\neq", "≠");
        value = value.Replace("\\alpha", "α").Replace("\\beta", "β").Replace("\\gamma", "γ").Replace("\\theta", "θ").Replace("\\pi", "π");
        value = value.Replace("{", "(").Replace("}", ")");
        return value;
    }

    private static string FirstSuccessfulGroup(Match match, params string[] names)
    {
        foreach (var name in names) if (match.Groups[name].Success) return match.Groups[name].Value;
        return match.Value;
    }

    private static MatchCollection FindAiLatexMatches(string source)
    {
        return Regex.Matches(
            source,
            @"\$\$(?<display>.+?)\$\$|(?<!\\)(?<!\$)\$(?!\$)(?<dollar>[^\r\n]+?)(?<!\\)\$(?!\$)|\\{1,2}\((?<paren>[^\r\n]+?)\\{1,2}\)|(?<!\\)\((?<gptparen>\\(?:frac|dfrac|tfrac|cfrac|sqrt|vec|overline|widehat|sin|cos|tan|cot|sec|csc|alpha|beta|gamma|delta|theta|lambda|mu|pi|sum|prod|int|lim|det|Pr)[^\r\n]*?)\)|\\{1,2}\[(?<bracket>.+?)\\{1,2}\]|```(?:latex|tex|math)?[ \t]*\r?\n(?<fence>.+?)\r?\n```|(?<environment>\\begin\s*\{(?<envname>math|displaymath|equation\*?|align\*?|aligned|alignedat|alignat\*?|flalign\*?|gather\*?|gathered|multline\*?|multlined|split|cases|dcases|array|matrix|smallmatrix|pmatrix|bmatrix|Bmatrix|vmatrix|Vmatrix)\}.*?\\end\s*\{\k<envname>\})|(?:\A|(?<=\r))[ \t]*\\?\[[ \t]*\r(?<block>.+?)\r[ \t]*\\?\][ \t]*(?=\r|\z)|(?:\A|(?<=\r))[ \t]*(?<line>\\(?:begin|sin|cos|tan|cot|sec|csc|frac|dfrac|tfrac|cfrac|sqrt|vec|overline|widehat|text|operatorname|mathbb|alpha|beta|gamma|delta|theta|lambda|mu|pi|sum|prod|int|iint|iiint|oint|lim|det|Pr)(?=\b|\d|\{|\s)[^\r\n]*)[ \t]*(?=\r|\z)",
            RegexOptions.Singleline);
    }

    private static List<AiFormulaCandidate> FindAiFormulaCandidates(string source)
    {
        source = source ?? String.Empty;
        var result = new List<AiFormulaCandidate>();
        var occupied = new List<Tuple<int, int>>();
        foreach (Match match in FindAiLatexMatches(source))
        {
            var formula = FirstSuccessfulGroup(match, "display", "dollar", "paren", "gptparen", "bracket", "fence", "environment", "block", "line").Trim();
            if (formula.Length == 0) continue;
            var explicitInline = match.Groups["dollar"].Success || match.Groups["paren"].Success || match.Groups["gptparen"].Success;
            var display = match.Groups["display"].Success || match.Groups["bracket"].Success ||
                          match.Groups["fence"].Success || match.Groups["environment"].Success ||
                          match.Groups["block"].Success || match.Groups["line"].Success ||
                          (!explicitInline && IsDisplayStyleFormula(formula));
            result.Add(new AiFormulaCandidate
            {
                Index = match.Index,
                Length = match.Length,
                Source = match.Value,
                Formula = formula,
                Display = display
            });
            occupied.Add(Tuple.Create(match.Index, match.Index + match.Length));
        }

        // AI 页面偶尔会丢失公式的起始定界符，例如
        // “同理 a<-\dfrac13\)。”。强 TeX 命令仍能可靠表明这是公式，
        // 因此只补捉包含这类命令的局部片段，不把整段普通文字改成公式。
        var strongCommands = Regex.Matches(source,
            @"\\(?:dfrac|tfrac|cfrac|frac|sqrt|binom|sum|prod|coprod|bigcup|bigcap|int|iint|iiint|oint|oiint|lim|sin|cos|tan|cot|sec|csc|sinh|cosh|log|ln|exp|det|rank|ker|Pr|vec|overrightarrow|xrightarrow|xleftarrow|overline|widehat|mathbb|alpha|beta|gamma|delta|epsilon|varepsilon|theta|lambda|mu|pi|varphi|phi|triangle|emptyset|varnothing|underline|operatorname)(?=\b|\d|\{|\s)",
            RegexOptions.IgnoreCase);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (Match command in strongCommands)
        {
            if (occupied.Any(span => command.Index >= span.Item1 && command.Index < span.Item2)) continue;
            var start = command.Index;
            while (start > 0 && IsBareFormulaCharacter(source[start - 1], false)) start--;
            var end = command.Index + command.Length;
            while (end < source.Length && IsBareFormulaCharacter(source[end], true)) end++;
            while (start < end && (source[start] == ' ' || source[start] == '\t')) start++;
            while (end > start && (source[end - 1] == ' ' || source[end - 1] == '\t')) end--;
            while (end > start && ".,;:".IndexOf(source[end - 1]) >= 0) end--;
            if (end <= start) continue;
            var raw = source.Substring(start, end - start);
            var formula = NormalizeBareFormulaBoundary(raw);
            if (String.IsNullOrWhiteSpace(formula) || !formula.Contains("\\")) continue;
            var key = start + ":" + end;
            if (!seen.Add(key)) continue;
            if (result.Any(existing => start < existing.Index + existing.Length && end > existing.Index)) continue;
            var lineStart = source.LastIndexOfAny(new[] { '\r', '\n' }, Math.Max(0, start - 1));
            lineStart = lineStart < 0 ? 0 : lineStart + 1;
            var lineEnd = source.IndexOfAny(new[] { '\r', '\n' }, end);
            if (lineEnd < 0) lineEnd = source.Length;
            var wholeLine = String.IsNullOrWhiteSpace(source.Substring(lineStart, start - lineStart)) &&
                            String.IsNullOrWhiteSpace(source.Substring(end, lineEnd - end));
            var hasOrphanClosingDelimiter = Regex.IsMatch(raw, @"\\{1,2}[\)\]]\s*$");
            if (!IsReliableBareFormula(formula, wholeLine, hasOrphanClosingDelimiter)) continue;
            result.Add(new AiFormulaCandidate
            {
                Index = start,
                Length = end - start,
                Source = raw,
                Formula = formula,
                Display = wholeLine && IsDisplayStyleFormula(formula)
            });
        }
        return result.OrderBy(item => item.Index).ThenByDescending(item => item.Length).ToList();
    }

    private static bool IsReliableBareFormula(string formula, bool wholeLine, bool hasOrphanClosingDelimiter)
    {
        if (String.IsNullOrWhiteSpace(formula)) return false;
        var converted = ConvertAiLatexExpression(formula);
        // 结构命令仍然残留，通常意味着复制到的是命令说明或残缺参数，
        // 这时保留原文比把说明文字创建成公式更合适。
        if (Regex.IsMatch(converted, @"\\(?:dfrac|tfrac|frac|sqrt|binom|begin|end|operatorname|text)\b",
            RegexOptions.IgnoreCase)) return false;
        if (hasOrphanClosingDelimiter) return true;

        var withoutCommands = Regex.Replace(formula, @"\\[A-Za-z]+", String.Empty);
        // 英文说明常见于“\frac{a}{b} means a fraction”。剔除命令后若仍有
        // 两个及以上的长英文单词，就不把整句当作数学表达式。
        if (Regex.Matches(withoutCommands, @"\b[A-Za-z]{3,}\b").Count >= 2) return false;
        var mathSignals = Regex.Matches(formula, @"[=<>+\-*/^_]|\d|\\(?:frac|sqrt|binom|sum|prod|coprod|bigcup|bigcap|int|iint|iiint|oint|lim|det|vec|overline|mathbb)\b",
            RegexOptions.IgnoreCase).Count;
        if (mathSignals == 0) return false;
        if (wholeLine) return true;
        // 行内无定界符公式至少需要一个完整的结构命令或明显的等式/不等式，
        // 避免把普通句子里的单个希腊字母命令误识别成整段公式。
        return Regex.IsMatch(formula, @"\\(?:dfrac|tfrac|cfrac|frac|sqrt|binom|vec|overrightarrow|overline|widehat|sum|prod|int|iint|iiint|oint|det|mathbb)\b|[=<>]",
            RegexOptions.IgnoreCase);
    }

    private static bool IsBareFormulaCharacter(char value, bool allowSpace)
    {
        if (value == '\r' || value == '\n' || value == '\a') return false;
        if (value == ' ' || value == '\t') return allowSpace;
        if (value == '\\' || value == '_' || value == '^') return true;
        if (value <= 127 && (Char.IsLetterOrDigit(value) || "{}()[]+-*/=<>,.|:;!%&~'".IndexOf(value) >= 0)) return true;
        return value >= '\u0370' && value <= '\u03ff' || value >= '\u2190' && value <= '\u22ff' ||
               value == '\u00b0' || value == '\u00b1' || value == '\u00d7' || value == '\u00f7' || value == '\u25b3';
    }

    private static string NormalizeBareFormulaBoundary(string source)
    {
        var value = (source ?? String.Empty).Trim();
        value = Regex.Replace(value, @"^\\{1,2}\(\s*", String.Empty);
        value = Regex.Replace(value, @"^\\{1,2}\[\s*", String.Empty);
        value = Regex.Replace(value, @"\\{1,2}\)(?=\)*\s*$)", String.Empty);
        value = Regex.Replace(value, @"\\{1,2}\](?=\]*\s*$)", String.Empty);
        value = Regex.Replace(value, @"\\+\s*$", String.Empty);
        return value.Trim();
    }

    private static string NormalizeLatexInput(string source)
    {
        var value = NormalizeBareFormulaBoundary(source);
        // 浏览器复制时偶尔会给普通运算符多加一个反斜杠，例如 \<、\-。
        value = Regex.Replace(value, @"\\(?=[<>=+\-*/])", String.Empty);
        value = Regex.Replace(value, @"([<>])\s*-\s*(?=[\\\d(])", "$1 -");
        value = RemoveUnmatchedLatexBraces(value);
        value = value.Replace("\\mleft", "\\left").Replace("\\mright", "\\right");
        value = Regex.Replace(value,
            @"\\(?:Biggl|Biggr|biggl|biggr|Biggm|biggm|Bigl|Bigr|bigl|bigr|Bigm|bigm|Bigg|bigg|Big|big)\b\s*",
            String.Empty);
        value = Regex.Replace(value,
            @"\\(?:middle|displaystyle|textstyle|scriptstyle|scriptscriptstyle|limits|nolimits)\b\s*",
            String.Empty);
        value = value.Replace("\\dfrac", "\\frac").Replace("\\tfrac", "\\frac");
        value = value.Replace("\\operatorname*", "\\operatorname");
        value = NormalizeBareFractionSyntax(value);
        return value.Trim();
    }

    private static string RemoveUnmatchedLatexBraces(string source)
    {
        var remove = new HashSet<int>();
        var opens = new Stack<int>();
        for (var index = 0; index < source.Length; index++)
        {
            if (source[index] != '{' && source[index] != '}') continue;
            var slashCount = 0;
            for (var cursor = index - 1; cursor >= 0 && source[cursor] == '\\'; cursor--) slashCount++;
            if (slashCount % 2 == 1) continue;
            if (source[index] == '{') opens.Push(index);
            else if (opens.Count > 0) opens.Pop();
            else remove.Add(index);
        }
        while (opens.Count > 0) remove.Add(opens.Pop());
        if (remove.Count == 0) return source;
        var builder = new StringBuilder(source.Length - remove.Count);
        for (var index = 0; index < source.Length; index++) if (!remove.Contains(index)) builder.Append(source[index]);
        return builder.ToString();
    }

    private static bool ReplaceWithOfficeMath(dynamic document, int start, int length, string latex, bool center)
    {
        latex = NormalizeLatexInput((latex ?? String.Empty).Replace('\a', ' ').Replace('\v', ' '));
        try
        {
            var documentStart = (int)document.Content.Start;
            var documentEnd = (int)document.Content.End;
            start = Math.Max(documentStart, Math.Min(start, documentEnd));
            length = Math.Max(0, Math.Min(length, documentEnd - start));
            if (length == 0 || String.IsNullOrWhiteSpace(latex)) return false;
        }
        catch { return false; }
        if (center && !IsWpsDocument(document))
        {
            var omml = OmmlMathBuilder.BuildEquationXml(latex);
            var fragmentPath = OmmlDocxFragment.Create(omml, true);
            if (!String.IsNullOrWhiteSpace(fragmentPath))
            {
                var originalText = String.Empty;
                var emptyDocumentEnd = -1;
                try
                {
                    var beforeCount = 0;
                    try { beforeCount = (int)document.OMaths.Count; } catch { }
                    dynamic sourceRange = document.Range(start, start + length);
                    originalText = sourceRange.Text as string ?? String.Empty;
                    sourceRange.Text = String.Empty;
                    emptyDocumentEnd = (int)document.Content.End;
                    dynamic insertion = document.Range(start, start);
                    try { insertion.InsertFile(fragmentPath); }
                    catch { insertion.InsertFile(fragmentPath, Type.Missing, false, false, false); }
                    var insertedLength = Math.Max(0, (int)document.Content.End - emptyDocumentEnd);
                    var afterCount = 0;
                    try { afterCount = (int)document.OMaths.Count; } catch { }
                    // WPS 有时延迟刷新 OMaths.Count；文档长度已经增加即可确认
                    // 片段写入成功，避免恢复原文后再走线性转换而产生双公式框。
                    if (afterCount <= beforeCount && insertedLength == 0)
                        throw new InvalidOperationException("公式片段未写入文档。");

                    dynamic insertedMath = null;
                    var nearestStart = Int32.MaxValue;
                    for (var equationIndex = 1; equationIndex <= afterCount; equationIndex++)
                    {
                        dynamic candidate = document.OMaths[equationIndex];
                        var candidateStart = (int)candidate.Range.Start;
                        if (candidateStart < start || candidateStart >= nearestStart) continue;
                        insertedMath = candidate;
                        nearestStart = candidateStart;
                    }
                    if (insertedMath != null)
                    {
                        insertedMath.Range.Font.Name = "Cambria Math";
                        dynamic paragraph = insertedMath.Range.Paragraphs[1].Range;
                        paragraph.ParagraphFormat.Alignment = 1;
                        paragraph.ParagraphFormat.SpaceBefore = 6f;
                        paragraph.ParagraphFormat.SpaceAfter = 8f;
                    }
                    return true;
                }
                catch
                {
                    try
                    {
                        if (emptyDocumentEnd >= 0)
                        {
                            var insertedLength = Math.Max(0, (int)document.Content.End - emptyDocumentEnd);
                            if (insertedLength > 0)
                                document.Range(start, Math.Min((int)document.Content.End, start + insertedLength)).Delete();
                        }
                    }
                    catch { }
                    try { document.Range(start, start).Text = originalText; } catch { }
                }
                finally
                {
                    try { File.Delete(fragmentPath); } catch { }
                }
            }
        }

        var rows = PrepareLatexRows(latex).Select(ConvertAiLatexExpression).Where(row => !String.IsNullOrWhiteSpace(row)).ToArray();
        if (rows.Length == 0) return false;
        var linear = String.Join("\r", rows);
        dynamic target = document.Range(start, start + length);
        var originalLinearSource = target.Text as string ?? String.Empty;
        try
        {
            target.Text = linear;
        }
        catch (Exception error)
        {
            // Word 表格、内容控件等结构可能临时锁定范围；跳过该范围，
            // 保留原文并继续处理其余公式。
            _ = error;
            return false;
        }
        var rowStart = start + linear.Length;
        var converted = 0;
        for (var index = rows.Length - 1; index >= 0; index--)
        {
            rowStart -= rows[index].Length;
            try
            {
                dynamic mathRange = document.Range(rowStart, rowStart + rows[index].Length);
                // OMaths.Add returns a Word.Range.  BuildUp belongs to the OMath/OMaths
                // inside that returned range; calling it on the Range silently left the
                // equation in linear form (for example "(a+b)/(c)").
                dynamic equationRange = null;
                try { equationRange = document.OMaths.Add(mathRange); }
                catch
                {
                    // WPS 主要在 Selection.OMaths 上公开 Add；参考其原生调用
                    // 顺序，选中当前小范围后再创建公式。
                }
                if (ResolveCreatedMath(document, equationRange, mathRange, rowStart, rowStart + rows[index].Length) == null)
                {
                    mathRange.Select();
                    dynamic selection = document.Application.Selection;
                    try { equationRange = selection.OMaths.Add(selection.Range); } catch { }
                }
                dynamic math = ResolveCreatedMath(document, equationRange, mathRange, rowStart, rowStart + rows[index].Length);
                if (math != null)
                {
                    var built = TryBuildMath(math);
                    if (!built)
                    {
                        try
                        {
                            mathRange.Select();
                            dynamic selection = document.Application.Selection;
                            try { selection.Range.OMaths.BuildUp(); built = true; }
                            catch { try { selection.OMaths.BuildUp(); built = true; } catch { } }
                            math = ResolveCreatedMath(document, equationRange, mathRange, rowStart, rowStart + rows[index].Length);
                            built = built && math != null;
                        }
                        catch { built = false; }
                    }
                    if (built)
                    {
                        try { math.Range.Font.Name = "Cambria Math"; } catch { }
                        if (center || rows.Length > 1) try { math.Range.ParagraphFormat.Alignment = 1; } catch { }
                        converted++;
                    }
                }
                else
                {
                    // 某些 WPS 版本只在选区集合上公开 BuildUp。
                    try
                    {
                        dynamic selection = document.Application.Selection;
                        selection.Range.OMaths.BuildUp();
                        math = ResolveCreatedMath(document, equationRange, mathRange, rowStart, rowStart + rows[index].Length);
                        if (math != null) converted++;
                    }
                    catch { }
                }
            }
            catch { }
            rowStart--;
        }
        if (converted != rows.Length)
        {
            // 多行公式按整体事务处理：只要任意一行失败，就恢复原始 TeX，
            // 避免留下半公式、线性文本或下次运行时重复创建公式框。
            try
            {
                var restoreEnd = Math.Min((int)document.Content.End, Math.Max((int)target.End, start + linear.Length));
                document.Range(start, restoreEnd).Text = originalLinearSource;
            }
            catch { }
            return false;
        }
        return true;
    }

    private static dynamic ResolveCreatedMath(dynamic document, dynamic equationRange, dynamic mathRange, int start, int end)
    {
        dynamic math = TryGetFirstMath(equationRange);
        if (math != null && MathRangesOverlap(math, start, end)) return math;
        math = TryGetFirstMath(mathRange);
        if (math != null && MathRangesOverlap(math, start, end)) return math;
        try
        {
            dynamic selection = document.Application.Selection;
            math = TryGetFirstMath(selection.Range);
            if (math != null && MathRangesOverlap(math, start, end)) return math;
        }
        catch { }
        dynamic nearest = null;
        var nearestDistance = Int32.MaxValue;
        var count = SafeCollectionCount(document.OMaths);
        for (var index = 1; index <= count; index++)
        {
            try
            {
                dynamic candidate = document.OMaths[index];
                if (!MathRangesOverlap(candidate, start, end)) continue;
                var distance = Math.Abs((int)candidate.Range.Start - start);
                if (distance >= nearestDistance) continue;
                nearest = candidate;
                nearestDistance = distance;
            }
            catch { }
        }
        return nearest;
    }

    private static dynamic TryGetFirstMath(dynamic holder)
    {
        if (holder == null) return null;
        try
        {
            dynamic collection = holder.OMaths;
            if (SafeCollectionCount(collection) > 0)
            {
                try { return collection.Item(1); }
                catch { return collection[1]; }
            }
        }
        catch { }
        return null;
    }

    private static bool MathRangesOverlap(dynamic math, int start, int end)
    {
        try { return (int)math.Range.Start < end && (int)math.Range.End > start; }
        catch { return false; }
    }

    private static bool TryBuildMath(dynamic math)
    {
        try { math.Range.OMaths.BuildUp(); return true; }
        catch
        {
            try { math.BuildUp(); return true; }
            catch
            {
                try { math.Parent.BuildUp(); return true; }
                catch { return false; }
            }
        }
    }

    private static bool IsWpsDocument(dynamic document)
    {
        try
        {
            var name = Convert.ToString(document.Application.Name) ?? String.Empty;
            return name.IndexOf("WPS", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.IndexOf("Kingsoft", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.IndexOf("金山", StringComparison.OrdinalIgnoreCase) >= 0;
        }
        catch { return false; }
    }

    private static string[] PrepareLatexRows(string latex)
    {
        var value = latex.Trim();
        var environment = Regex.Match(
            value,
            @"\\begin\s*\{(?<name>math|displaymath|equation\*?|align\*?|aligned|alignedat|alignat\*?|flalign\*?|gather\*?|gathered|multline\*?|multlined|split|cases|dcases|array|matrix|smallmatrix|pmatrix|bmatrix|Bmatrix|vmatrix|Vmatrix)\}(?:\s*\{[^{}]*\})?(?<body>.*?)\\end\s*\{\k<name>\}",
            RegexOptions.Singleline | RegexOptions.IgnoreCase);
        if (!environment.Success)
        {
            environment = Regex.Match(
                value,
                @"\\begin\s*\(\s*(?<name>align\s*\*?|aligned|gather\s*\*?|gathered|split|equation\s*\*?)\s*\)(?<body>.*?)\\end\s*\(\s*[^)]*\)",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);
        }
        if (environment.Success)
        {
            var prefix = NormalizeAiEscapes(value.Substring(0, environment.Index)).Trim();
            var suffix = NormalizeAiEscapes(value.Substring(environment.Index + environment.Length)).Trim();
            var rawName = environment.Groups["name"].Value.Replace(" ", String.Empty);
            var name = rawName.ToLowerInvariant();
            var body = environment.Groups["body"].Value.Trim();
            if (Regex.IsMatch(body, @"^\\begin\s*\{", RegexOptions.IgnoreCase)) return PrepareLatexRows(body);
            var rows = Regex.Split(body, @"\\\\(?=\s|\\|$|[0-9+\-(]|(?!(?:sin|cos|tan|cot|sec|csc|arcsin|arccos|arctan|frac|dfrac|tfrac|cfrac|sqrt|text|mathrm|mathbf|mathit|mathsf|mathtt|operatorname|left|right|begin|end|alpha|beta|gamma|delta|epsilon|theta|lambda|mu|nu|xi|pi|rho|sigma|tau|phi|varphi|omega|sum|prod|int|lim|log|ln|exp|quad|qquad|cdot|times)\b)[A-Za-z])", RegexOptions.IgnoreCase)
                .Select(row => NormalizeAiEscapes(row).Trim().TrimEnd('\\').Trim())
                .Where(row => row.Length > 0)
                .ToArray();
            if (name.Contains("matrix") || name == "array")
            {
                var matrix = "\\matrix(" + String.Join("@", rows.Select(row => row.Replace("&", "&"))) + ")";
                if (name == "pmatrix") matrix = "(" + matrix + ")";
                else if (name == "bmatrix") matrix = "[" + matrix + "]";
                else if (rawName == "Bmatrix") matrix = "{" + matrix + "}";
                else if (rawName == "Vmatrix") matrix = "‖" + matrix + "‖";
                else if (name == "vmatrix") matrix = "|" + matrix + "|";
                return new[] { prefix + matrix + suffix };
            }
            if (name == "cases" || name == "dcases") return new[] { prefix + "\\cases(" + String.Join("@", rows) + ")" + suffix };
            if (name.StartsWith("align", StringComparison.Ordinal) || name.StartsWith("flalign", StringComparison.Ordinal) || name == "aligned" || name == "split")
                return new[] { prefix + "\\eqarray(" + String.Join("@", rows.Select(row => NormalizeAlignmentPoints(row))) + ")" + suffix };
            if (name.StartsWith("gather", StringComparison.Ordinal) || name.StartsWith("multline", StringComparison.Ordinal) || name.StartsWith("equation", StringComparison.Ordinal) || name == "displaymath" || name == "math")
                return new[] { prefix + "\\eqarray(" + String.Join("@", rows.Select(row => row.Replace("&", String.Empty))) + ")" + suffix };
            var plainRows = rows.Select(row => row.Replace("&", String.Empty)).ToArray();
            if (plainRows.Length > 0)
            {
                plainRows[0] = prefix + plainRows[0];
                plainRows[plainRows.Length - 1] += suffix;
            }
            return plainRows;
        }
        return new[] { NormalizeAiEscapes(value).Replace("&=", "=").Replace("&", String.Empty) };
    }

    private static string NormalizeAlignmentPoints(string row)
    {
        var value = row.Trim();
        var first = value.IndexOf('&');
        if (first < 0) return value;
        value = value.Substring(0, first) + "&" + value.Substring(first + 1).Replace("&", String.Empty);
        return value;
    }

    private static string NormalizeAiEscapes(string value)
    {
        value = Regex.Replace(value, @"\\\\(?=[A-Za-z])", "\\");
        value = value.Replace("\\\\", " ");
        value = Regex.Replace(value, @"\\(?=[ \t])", String.Empty);
        return value;
    }

    private static string ConvertAiLatexExpression(string latex)
    {
        var value = NormalizeAiEscapes(NormalizeLatexInput(latex.Trim()));
        // 这些命令只控制 TeX 括号大小或显示样式；Office Math 会自行调整，
        // 保留它们反而会在 Word/WPS 中显示成“\big”等原始文本。
        value = Regex.Replace(value, @"\\(?:mleft|mright|left|right)\s*\.", String.Empty);
        value = Regex.Replace(value,
            @"\\(?:mleft|mright|middle|left|right|Biggl|Biggr|biggl|biggr|Biggm|biggm|Bigl|Bigr|bigl|bigr|Bigm|bigm|Bigg|bigg|Big|big|displaystyle|textstyle|scriptstyle|scriptscriptstyle|limits|nolimits)\b\s*",
            String.Empty);
        value = value.Replace("\\dfrac", "\\frac").Replace("\\tfrac", "\\frac").Replace("\\cfrac", "\\frac");
        value = NormalizeBareFractionSyntax(value);
        value = NormalizeSingleArgumentCommands(value);
        value = ReplaceTwoGroupCommand(value, "\\frac", (left, right) => "(" + ConvertAiLatexExpression(left) + ")/(" + ConvertAiLatexExpression(right) + ")");
        value = ReplaceTwoGroupCommand(value, "\\binom", (top, bottom) => "\\binom(" + ConvertAiLatexExpression(top) + "," + ConvertAiLatexExpression(bottom) + ")");
        value = ReplaceGenfracCommand(value);
        value = ReplaceIndexedRootCommands(value);
        value = ReplaceOneGroupCommand(value, "\\sqrt", inner => "√(" + ConvertAiLatexExpression(inner) + ")");
        value = ReplaceOneGroupCommand(value, "\\text", inner => inner);
        value = ReplaceOneGroupCommand(value, "\\mathrm", inner => inner);
        value = ReplaceOneGroupCommand(value, "\\mathbf", inner => inner);
        value = ReplaceOneGroupCommand(value, "\\boldsymbol", ConvertAiLatexExpression);
        value = ReplaceOneGroupCommand(value, "\\bm", ConvertAiLatexExpression);
        value = ReplaceOneGroupCommand(value, "\\pmb", ConvertAiLatexExpression);
        value = ReplaceOneGroupCommand(value, "\\mathnormal", ConvertAiLatexExpression);
        value = ReplaceOneGroupCommand(value, "\\mathit", inner => inner);
        value = ReplaceOneGroupCommand(value, "\\mathsf", inner => inner);
        value = ReplaceOneGroupCommand(value, "\\mathtt", inner => inner);
        value = ReplaceOneGroupCommand(value, "\\operatorname", inner => inner + " ");
        value = Regex.Replace(value, @"\\mathbb\s*([A-Za-z0-9])", match => ToDoubleStruckText(match.Groups[1].Value));
        value = Regex.Replace(value, @"\\overline\s*\{([A-Za-z0-9])\}", "$1\u0305");
        value = ReplaceOneGroupCommand(value, "\\mathbb", ToDoubleStruckText);
        value = ReplaceOneGroupCommand(value, "\\mathcal", inner => inner);
        value = ReplaceOneGroupCommand(value, "\\mathfrak", inner => inner);
        value = ReplaceOneGroupCommand(value, "\\overline", inner => "\\bar(" + ConvertAiLatexExpression(inner) + ")");
        value = ReplaceOneGroupCommand(value, "\\underline", inner => UnderlineLinear(ConvertAiLatexExpression(inner)));
        value = ReplaceOneGroupCommand(value, "\\vec", inner => "\\vec(" + ConvertAiLatexExpression(inner) + ")");
        value = ReplaceOneGroupCommand(value, "\\overrightarrow", inner => "\\vec(" + ConvertAiLatexExpression(inner) + ")");
        value = ReplaceOneGroupCommand(value, "\\overleftarrow", inner => "\\vec(" + ConvertAiLatexExpression(inner) + ")");
        value = ReplaceOneGroupCommand(value, "\\widehat", inner => "\\hat(" + ConvertAiLatexExpression(inner) + ")");
        value = ReplaceOneGroupCommand(value, "\\widetilde", inner => "\\tilde(" + ConvertAiLatexExpression(inner) + ")");
        value = ReplaceOneGroupCommand(value, "\\underrightarrow", inner => "\\vec(" + ConvertAiLatexExpression(inner) + ")");
        value = ReplaceOneGroupCommand(value, "\\underleftarrow", inner => "\\vec(" + ConvertAiLatexExpression(inner) + ")");
        value = ReplaceOneGroupCommand(value, "\\wideparen", inner => "\\hat(" + ConvertAiLatexExpression(inner) + ")");
        value = ReplaceOneGroupCommand(value, "\\overparen", inner => "\\hat(" + ConvertAiLatexExpression(inner) + ")");
        foreach (var accent in new[] { "bar", "hat", "tilde", "dot", "ddot", "breve", "check", "acute", "grave" })
        {
            var command = "\\" + accent;
            value = ReplaceOneGroupCommand(value, command, inner => command + "(" + ConvertAiLatexExpression(inner) + ")");
        }
        value = ReplaceOneGroupCommand(value, "\\boxed", inner => "□(" + ConvertAiLatexExpression(inner) + ")");
        value = ReplaceOneGroupCommand(value, "\\cancel", inner => "⊘(" + ConvertAiLatexExpression(inner) + ")");
        value = ReplaceOneGroupCommand(value, "\\bcancel", inner => "⊘(" + ConvertAiLatexExpression(inner) + ")");
        value = ReplaceOneGroupCommand(value, "\\xcancel", inner => "⊠(" + ConvertAiLatexExpression(inner) + ")");
        value = ReplaceOneGroupCommand(value, "\\overbrace", inner => "⏞(" + ConvertAiLatexExpression(inner) + ")");
        value = ReplaceOneGroupCommand(value, "\\underbrace", inner => "⏟(" + ConvertAiLatexExpression(inner) + ")");
        value = ReplaceOneGroupCommand(value, "\\substack", inner => ConvertAiLatexExpression(inner.Replace("\\\\", ",")));
        value = ReplaceTwoGroupCommand(value, "\\textcolor", (color, body) => ConvertAiLatexExpression(body));
        value = ReplaceTwoGroupCommand(value, "\\overset", (top, body) => "(" + ConvertAiLatexExpression(body) + ")^(" + ConvertAiLatexExpression(top) + ")");
        value = ReplaceTwoGroupCommand(value, "\\stackrel", (top, body) => "(" + ConvertAiLatexExpression(body) + ")^(" + ConvertAiLatexExpression(top) + ")");
        value = ReplaceTwoGroupCommand(value, "\\underset", (bottom, body) => "(" + ConvertAiLatexExpression(body) + ")_(" + ConvertAiLatexExpression(bottom) + ")");
        value = ReplaceOneGroupCommand(value, "\\xrightarrow", top => "(→)^(" + ConvertAiLatexExpression(top) + ")");
        value = ReplaceOneGroupCommand(value, "\\xleftarrow", top => "(←)^(" + ConvertAiLatexExpression(top) + ")");
        value = ReplaceOneGroupCommand(value, "\\pmod", inner => "( mod " + ConvertAiLatexExpression(inner) + ")");
        value = ReplaceOneGroupCommand(value, "\\pod", inner => "(" + ConvertAiLatexExpression(inner) + ")");
        value = Regex.Replace(value, @"\\(?:bmod|mod)(?![A-Za-z])", " mod ");
        value = value.Replace("\\qquad", "\u2003\u2003").Replace("\\quad", "\u2003");
        value = value.Replace("\\,", " ").Replace("\\;", " ").Replace("\\!", String.Empty);
        value = value.Replace("\\%", "%").Replace("\\#", "#").Replace("\\&", "&");
        value = Regex.Replace(value,
            @"\\(?<name>arcsin|arccos|arctan|sinh|cosh|tanh|coth|sin|cos|tan|cot|sec|csc|limsup|liminf|lim|log|ln|exp|min|max|sup|inf|det|gcd|lcm|rank|diag|span|sgn|tr|ker|dim|arg|Pr|Var|Cov)(?![A-Za-z])",
            match => " " + match.Groups["name"].Value + " ", RegexOptions.IgnoreCase);
        // 箭头命令必须先于 \le/\ge 等短命令处理；否则 \leftarrow 会被
        // \le 的前缀替换破坏成“≤ftarrow”。
        value = value.Replace("\\Longrightarrow", "\u27f9").Replace("\\Longleftarrow", "\u27f8").Replace("\\Longleftrightarrow", "\u27fa");
        value = value.Replace("\\longrightarrow", "\u27f6").Replace("\\longleftarrow", "\u27f5").Replace("\\longleftrightarrow", "\u27f7");
        value = value.Replace("\\Rightarrow", "\u21d2").Replace("\\Leftarrow", "\u21d0").Replace("\\Leftrightarrow", "\u21d4").Replace("\\rightarrow", "\u2192").Replace("\\leftarrow", "\u2190").Replace("\\leftrightarrow", "\u2194");
        value = value.Replace("\\implies", "\u21d2").Replace("\\iff", "\u21d4").Replace("\\mapsto", "\u21a6").Replace("\\to", "\u2192");
        // 先处理长命令，避免 \int 被 \in、\cdots 被 \cdot 截断。
        value = value.Replace("\\oiiint", "\u2230").Replace("\\oiint", "\u222f").Replace("\\iiint", "\u222d").Replace("\\iint", "\u222c").Replace("\\oint", "\u222e");
        value = value.Replace("\\bigcup", "\u22c3").Replace("\\bigcap", "\u22c2").Replace("\\coprod", "\u2210").Replace("\\sum", "\u2211").Replace("\\prod", "\u220f").Replace("\\int", "\u222b");
        value = value.Replace("\\cdots", "\u22ef").Replace("\\ldots", "\u2026").Replace("\\dots", "\u2026");
        value = ReplaceAdditionalLatexSymbols(value);
        value = value.Replace("\\leqslant", "\u2264").Replace("\\geqslant", "\u2265").Replace("\\subsetneq", "\u228a").Replace("\\supsetneq", "\u228b");
        value = value.Replace("\\times", "\u00d7").Replace("\\cdot", "\u00b7").Replace("\\leq", "\u2264").Replace("\\geq", "\u2265").Replace("\\le", "\u2264").Replace("\\ge", "\u2265").Replace("\\neq", "\u2260").Replace("\\ne", "\u2260");
        value = value.Replace("\\sim", "\u223c").Replace("\\circ", "\u2218").Replace("\\complement", "\u2201");
        value = value.Replace("\\approx", "\u2248").Replace("\\equiv", "\u2261").Replace("\\infty", "\u221e").Replace("\\pm", "\u00b1").Replace("\\mp", "\u2213").Replace("\\div", "\u00f7");
        value = value.Replace("\\in", "\u2208").Replace("\\notin", "\u2209").Replace("\\subseteq", "\u2286").Replace("\\supseteq", "\u2287").Replace("\\subset", "\u2282").Replace("\\supset", "\u2283");
        value = value.Replace("\\cup", "\u222a").Replace("\\cap", "\u2229").Replace("\\forall", "\u2200").Replace("\\exists", "\u2203").Replace("\\partial", "\u2202").Replace("\\nabla", "\u2207");
        value = value.Replace("\\therefore", "\u2234").Replace("\\because", "\u2235").Replace("\\perp", "\u22a5").Replace("\\parallel", "\u2225").Replace("\\angle", "\u2220");
        value = value.Replace("\\lVert", "\u2016").Replace("\\rVert", "\u2016").Replace("\\Vert", "\u2016");
        value = value.Replace("\\lvert", "|").Replace("\\rvert", "|").Replace("\\vert", "|");
        value = value.Replace("\\langle", "\u27e8").Replace("\\rangle", "\u27e9");
        value = value.Replace("\\lfloor", "\u230a").Replace("\\rfloor", "\u230b").Replace("\\lceil", "\u2308").Replace("\\rceil", "\u2309");
        value = value.Replace("\\alpha", "\u03b1").Replace("\\beta", "\u03b2").Replace("\\gamma", "\u03b3").Replace("\\delta", "\u03b4").Replace("\\epsilon", "\u03b5").Replace("\\varepsilon", "\u03b5");
        value = value.Replace("\\zeta", "\u03b6").Replace("\\eta", "\u03b7").Replace("\\theta", "\u03b8").Replace("\\iota", "\u03b9").Replace("\\kappa", "\u03ba").Replace("\\lambda", "\u03bb");
        value = value.Replace("\\mu", "\u03bc").Replace("\\nu", "\u03bd").Replace("\\xi", "\u03be").Replace("\\rho", "\u03c1").Replace("\\sigma", "\u03c3").Replace("\\tau", "\u03c4");
        value = value.Replace("\\upsilon", "\u03c5").Replace("\\omega", "\u03c9").Replace("\\pi", "\u03c0").Replace("\\varphi", "\u03c6").Replace("\\phi", "\u03c6");
        value = value.Replace("\\triangle", "\u25b3");
        value = value.Replace("\\emptyset", "\u2205").Replace("\\varnothing", "\u2205");
        value = ReplaceAdditionalLatexSymbols(value);
        value = value.Replace("\\{", "\uE100").Replace("\\}", "\uE101");
        value = value.Replace("{", "(").Replace("}", ")").Replace("\uE100", "{").Replace("\uE101", "}");
        return RemoveUnmatchedClosingParentheses(Regex.Replace(value, @"[ \t]+", " ").Trim());
    }

    private static string NormalizeSingleArgumentCommands(string source)
    {
        return Regex.Replace(source ?? String.Empty,
            @"\\(?<name>sqrt|overline|bar|underline|vec|overrightarrow|overleftarrow|underrightarrow|underleftarrow|widehat|hat|widetilde|tilde|dot|ddot|breve|check|acute|grave|wideparen|overparen|mathbb|mathcal|mathfrak|mathrm|mathbf|mathit|mathsf|mathtt|boldsymbol|bm|pmb|pmod|pod)(?![A-Za-z])(?!\s*\{)\s*(?<arg>\\[A-Za-z]+|[A-Za-z0-9])",
            match => "\\" + match.Groups["name"].Value + "{" + match.Groups["arg"].Value + "}");
    }

    private static string ReplaceGenfracCommand(string source)
    {
        const string command = "\\genfrac";
        var searchFrom = 0;
        while (searchFrom < source.Length)
        {
            var index = source.IndexOf(command, searchFrom, StringComparison.Ordinal);
            if (index < 0) break;
            var cursor = index + command.Length;
            var groups = new string[6];
            var valid = true;
            for (var group = 0; group < groups.Length; group++)
            {
                while (cursor < source.Length && Char.IsWhiteSpace(source[cursor])) cursor++;
                if (!TryReadGroup(source, cursor, out groups[group], out cursor)) { valid = false; break; }
            }
            if (!valid)
            {
                searchFrom = index + command.Length;
                continue;
            }
            var top = ConvertAiLatexExpression(groups[4]);
            var bottom = ConvertAiLatexExpression(groups[5]);
            var noBar = String.Equals(groups[2].Trim(), "0pt", StringComparison.OrdinalIgnoreCase);
            var fraction = noBar ? "\\binom(" + top + "," + bottom + ")" : "(" + top + ")/(" + bottom + ")";
            if (!String.IsNullOrWhiteSpace(groups[0]) || !String.IsNullOrWhiteSpace(groups[1]))
                fraction = ConvertAiLatexExpression(groups[0]) + fraction + ConvertAiLatexExpression(groups[1]);
            source = source.Substring(0, index) + fraction + source.Substring(cursor);
            searchFrom = index + fraction.Length;
        }
        return source;
    }

    private static string ReplaceIndexedRootCommands(string source)
    {
        var searchFrom = 0;
        const string command = "\\sqrt";
        while (searchFrom < source.Length)
        {
            var commandIndex = source.IndexOf(command, searchFrom, StringComparison.Ordinal);
            if (commandIndex < 0) break;
            var cursor = commandIndex + command.Length;
            while (cursor < source.Length && Char.IsWhiteSpace(source[cursor])) cursor++;
            if (cursor >= source.Length || source[cursor] != '[')
            {
                searchFrom = cursor;
                continue;
            }
            var close = source.IndexOf(']', cursor + 1);
            if (close < 0)
            {
                searchFrom = cursor + 1;
                continue;
            }
            var degree = source.Substring(cursor + 1, close - cursor - 1);
            cursor = close + 1;
            while (cursor < source.Length && Char.IsWhiteSpace(source[cursor])) cursor++;
            if (!TryReadGroup(source, cursor, out var body, out var end))
            {
                searchFrom = close + 1;
                continue;
            }
            var result = "√(" + ConvertAiLatexExpression(degree) + "&" + ConvertAiLatexExpression(body) + ")";
            source = source.Substring(0, commandIndex) + result + source.Substring(end);
            searchFrom = commandIndex + result.Length;
        }
        return source;
    }

    private static bool HasUnmatchedClosingParenthesis(string source)
    {
        var depth = 0;
        foreach (var character in source ?? String.Empty)
        {
            if (character == '(') depth++;
            else if (character == ')' && depth == 0) return true;
            else if (character == ')') depth--;
        }
        return false;
    }

    private static string RemoveUnmatchedClosingParentheses(string source)
    {
        if (String.IsNullOrEmpty(source)) return source ?? String.Empty;
        var depth = 0;
        var builder = new StringBuilder(source.Length);
        foreach (var character in source)
        {
            if (character == '(')
            {
                depth++;
                builder.Append(character);
            }
            else if (character == ')')
            {
                if (depth == 0) continue;
                depth--;
                builder.Append(character);
            }
            else builder.Append(character);
        }
        return builder.ToString();
    }

    private static string NormalizeBareFractionSyntax(string source)
    {
        var searchFrom = 0;
        const string command = "\\frac";
        while (searchFrom < source.Length)
        {
            var commandIndex = source.IndexOf(command, searchFrom, StringComparison.Ordinal);
            if (commandIndex < 0) break;
            var cursor = commandIndex + command.Length;
            if (cursor < source.Length && Char.IsLetter(source[cursor]))
            {
                searchFrom = cursor;
                continue;
            }
            if (!TryReadLatexArgument(source, cursor, out var numerator, out cursor) ||
                !TryReadLatexArgument(source, cursor, out var denominator, out var end))
            {
                searchFrom = commandIndex + command.Length;
                continue;
            }
            var replacement = command + "{" + numerator + "}{" + denominator + "}";
            source = source.Substring(0, commandIndex) + replacement + source.Substring(end);
            searchFrom = commandIndex + replacement.Length;
        }
        return source;
    }

    private static bool TryReadLatexArgument(string source, int start, out string value, out int next)
    {
        value = String.Empty;
        next = start;
        while (next < source.Length && Char.IsWhiteSpace(source[next])) next++;
        if (next >= source.Length) return false;
        if (source[next] == '{') return TryReadGroup(source, next, out value, out next);
        var argumentStart = next;
        if (source[next] == '\\')
        {
            next++;
            if (next < source.Length && Char.IsLetter(source[next]))
                while (next < source.Length && Char.IsLetter(source[next])) next++;
            else if (next < source.Length) next++;
        }
        else next++;
        value = source.Substring(argumentStart, next - argumentStart);
        return value.Length > 0;
    }

    private static string UnderlineLinear(string value)
    {
        if (String.IsNullOrWhiteSpace(value) || value.Trim().All(character => Char.IsWhiteSpace(character))) return "────────";
        var builder = new StringBuilder();
        foreach (var character in value) builder.Append(character).Append('\u0332');
        return builder.ToString();
    }

    private static string ToDoubleStruckText(string value)
    {
        var builder = new StringBuilder();
        foreach (var character in value ?? String.Empty)
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

    private static string ReplaceAdditionalLatexSymbols(string value)
    {
        var symbols = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["alpha"]="α", ["beta"]="β", ["gamma"]="γ", ["delta"]="δ", ["epsilon"]="ε", ["varepsilon"]="ϵ",
            ["zeta"]="ζ", ["eta"]="η", ["theta"]="θ", ["vartheta"]="ϑ", ["iota"]="ι", ["kappa"]="κ", ["lambda"]="λ",
            ["mu"]="μ", ["nu"]="ν", ["xi"]="ξ", ["pi"]="π", ["varpi"]="ϖ", ["rho"]="ρ", ["varrho"]="ϱ",
            ["sigma"]="σ", ["varsigma"]="ς", ["tau"]="τ", ["upsilon"]="υ", ["phi"]="φ", ["varphi"]="ϕ", ["chi"]="χ", ["psi"]="ψ", ["omega"]="ω",
            ["Gamma"]="Γ", ["Delta"]="Δ", ["Theta"]="Θ", ["Lambda"]="Λ", ["Xi"]="Ξ", ["Pi"]="Π", ["Sigma"]="Σ", ["Upsilon"]="Υ", ["Phi"]="Φ", ["Psi"]="Ψ", ["Omega"]="Ω",
            ["subsetneq"]="⊊", ["supsetneq"]="⊋", ["nsubseteq"]="⊈", ["nsupseteq"]="⊉",
            ["setminus"]="∖", ["smallsetminus"]="∖", ["oplus"]="⊕", ["ominus"]="⊖", ["otimes"]="⊗", ["oslash"]="⊘", ["odot"]="⊙",
            ["uplus"]="⊎", ["sqcup"]="⊔", ["sqcap"]="⊓", ["sqsubseteq"]="⊑", ["sqsupseteq"]="⊒",
            ["land"]="∧", ["wedge"]="∧", ["lor"]="∨", ["vee"]="∨", ["neg"]="¬", ["lnot"]="¬",
            ["top"]="⊤", ["bot"]="⊥", ["vdash"]="⊢", ["dashv"]="⊣", ["models"]="⊨", ["ni"]="∋", ["nexists"]="∄",
            ["propto"]="∝", ["cong"]="≅", ["simeq"]="≃", ["asymp"]="≍", ["doteq"]="≐", ["prec"]="≺", ["succ"]="≻", ["preceq"]="≼", ["succeq"]="≽",
            ["ll"]="≪", ["gg"]="≫", ["lesssim"]="≲", ["gtrsim"]="≳", ["lessapprox"]="⪅", ["gtrapprox"]="⪆", ["mid"]="∣", ["nmid"]="∤",
            ["hbar"]="ℏ", ["aleph"]="ℵ", ["Re"]="ℜ", ["Im"]="ℑ", ["wp"]="℘", ["ell"]="ℓ",
            ["measuredangle"]="∡", ["degree"]="°", ["vdots"]="⋮", ["ddots"]="⋱",
            ["uparrow"]="↑", ["downarrow"]="↓", ["updownarrow"]="↕", ["Uparrow"]="⇑", ["Downarrow"]="⇓", ["Updownarrow"]="⇕",
            ["hookrightarrow"]="↪", ["hookleftarrow"]="↩", ["rightharpoonup"]="⇀", ["leftharpoonup"]="↼", ["rightleftharpoons"]="⇌",
            ["bigvee"]="⋁", ["bigwedge"]="⋀"
        };
        return Regex.Replace(value, @"\\(?<name>[A-Za-z]+)(?![A-Za-z])", match =>
            symbols.TryGetValue(match.Groups["name"].Value, out var replacement) ? replacement : match.Value);
    }

    private static string ReplaceTwoGroupCommand(string source, string command, Func<string, string, string> replacement)
    {
        var searchFrom = 0;
        while (searchFrom < source.Length)
        {
            var commandIndex = source.IndexOf(command, searchFrom, StringComparison.Ordinal);
            if (commandIndex < 0) break;
            var cursor = commandIndex + command.Length;
            while (cursor < source.Length && Char.IsWhiteSpace(source[cursor])) cursor++;
            if (!TryReadGroup(source, cursor, out var first, out cursor)) { searchFrom = commandIndex + command.Length; continue; }
            while (cursor < source.Length && Char.IsWhiteSpace(source[cursor])) cursor++;
            if (!TryReadGroup(source, cursor, out var second, out var end)) { searchFrom = commandIndex + command.Length; continue; }
            var result = replacement(first, second);
            source = source.Substring(0, commandIndex) + result + source.Substring(end);
            searchFrom = commandIndex + result.Length;
        }
        return source;
    }

    private static string ReplaceOneGroupCommand(string source, string command, Func<string, string> replacement)
    {
        var searchFrom = 0;
        while (searchFrom < source.Length)
        {
            var commandIndex = source.IndexOf(command, searchFrom, StringComparison.Ordinal);
            if (commandIndex < 0) break;
            var cursor = commandIndex + command.Length;
            while (cursor < source.Length && Char.IsWhiteSpace(source[cursor])) cursor++;
            if (!TryReadGroup(source, cursor, out var inner, out var end)) { searchFrom = commandIndex + command.Length; continue; }
            var result = replacement(inner);
            source = source.Substring(0, commandIndex) + result + source.Substring(end);
            searchFrom = commandIndex + result.Length;
        }
        return source;
    }

    private static bool TryReadGroup(string source, int openIndex, out string value, out int nextIndex)
    {
        value = String.Empty;
        nextIndex = openIndex;
        if (openIndex >= source.Length || source[openIndex] != '{') return false;
        var depth = 0;
        for (var index = openIndex; index < source.Length; index++)
        {
            if (source[index] == '{') depth++;
            else if (source[index] == '}')
            {
                depth--;
                if (depth == 0)
                {
                    value = source.Substring(openIndex + 1, index - openIndex - 1);
                    nextIndex = index + 1;
                    return true;
                }
            }
        }
        return false;
    }

    public int CollapseRepeatedSpaces()
    {
        return CollapseRepeatedSpaces("Document");
    }

    public int CollapseRepeatedSpaces(string scope)
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic range = document.Content;
        var selectionOnly = String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase);
        if (selectionOnly)
        {
            dynamic selected = ((dynamic)_application).Selection.Range;
            if ((int)selected.End <= (int)selected.Start) throw new InvalidOperationException("请先选中需要处理的内容。");
            range = selected.Duplicate;
        }
        dynamic trackingRange = range.Duplicate;
        var passes = 0;
        while (passes < 30)
        {
            dynamic passRange = trackingRange.Duplicate;
            dynamic find = passRange.Find;
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            find.Text = "  ";
            find.Replacement.Text = " ";
            find.Forward = true;
            find.Wrap = selectionOnly ? 0 : 1;
            find.Format = false;
            find.MatchWildcards = false;
            var changed = ExecuteReplaceAll(find);
            if (!changed) changed = ReplaceLiteralRangesInRange(document, passRange, "  ", " ") > 0;
            if (!changed) break;
            passes++;
        }
        SelectScopeRange(trackingRange, scope);
        return passes;
    }

    public int NormalizeBlankLines(string scope)
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic range = document.Content;
        if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase))
        {
            dynamic selected = ((dynamic)_application).Selection.Range;
            if ((int)selected.End <= (int)selected.Start) throw new InvalidOperationException("请先选中需要处理的内容。");
            range = selected.Duplicate;
        }
        dynamic trackingRange = range.Duplicate;
        var rounds = 0;
        while (rounds < 12 && ReplaceLiteralRangesInRange(document, trackingRange.Duplicate, "\r\r\r", "\r\r") > 0) rounds++;
        SelectScopeRange(trackingRange, scope);
        return rounds;
    }

    private static int ReplaceLiteralRangesInRange(dynamic document, dynamic scope, string findText, string replaceText)
    {
        var source = scope.Text as string ?? String.Empty;
        var positions = new List<int>();
        for (var index = source.IndexOf(findText, StringComparison.Ordinal);
             index >= 0;
             index = source.IndexOf(findText, index + Math.Max(1, findText.Length), StringComparison.Ordinal)) positions.Add(index);
        var start = (int)scope.Start;
        var changed = 0;
        foreach (var position in positions.OrderByDescending(value => value))
        {
            try { document.Range(start + position, start + position + findText.Length).Text = replaceText ?? String.Empty; changed++; }
            catch { }
        }
        return changed;
    }

    private static void ReplaceAllInDocument(dynamic document, string findText, string replaceText)
    {
        if (String.IsNullOrEmpty(findText)) return;
        dynamic range = document.Content;
        dynamic find = range.Find;
        find.ClearFormatting(); find.Replacement.ClearFormatting();
        find.Text = findText; find.Replacement.Text = replaceText;
        find.Forward = true; find.Wrap = 1; find.Format = false; find.MatchWildcards = false;
        if (!ExecuteReplaceAll(find)) ReplaceLiteralRanges(document, findText, replaceText);
    }

    public void ApplyRule(TextReplacement rule)
    {
        ApplyRule(rule, "Document");
    }

    public void ApplyRule(TextReplacement rule, string scope)
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic range = document.Content;
        var selectionOnly = String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase);
        if (selectionOnly)
        {
            dynamic selected = ((dynamic)_application).Selection.Range;
            if ((int)selected.End <= (int)selected.Start) throw new InvalidOperationException("请先选中需要应用规则的内容。");
            range = selected.Duplicate;
        }
        dynamic trackingRange = range.Duplicate;
        ApplyRuleToRange(document, trackingRange.Duplicate, rule, selectionOnly);
        SelectScopeRange(trackingRange, scope);
    }

    private static void ApplyRuleToDocument(dynamic document, TextReplacement rule)
    {
        ApplyRuleToRange(document, document.Content, rule, false);
    }

    private static void ApplyRuleToRange(dynamic document, dynamic range, TextReplacement rule, bool selectionOnly)
    {
        if (rule == null || String.IsNullOrEmpty(rule.Find)) return;
        dynamic trackingRange = range.Duplicate;
        dynamic searchRange = trackingRange.Duplicate;
        dynamic find = searchRange.Find;
        find.ClearFormatting(); find.Replacement.ClearFormatting();
        find.Text = rule.Find; find.Replacement.Text = rule.Replace ?? String.Empty;
        find.Forward = true; find.Wrap = selectionOnly ? 0 : 1; find.Format = rule.HasFormatting; find.MatchWildcards = rule.UseWildcards;
        if (rule.Bold.HasValue) find.Replacement.Font.Bold = rule.Bold.Value ? 1 : 0;
        if (rule.Italic.HasValue) find.Replacement.Font.Italic = rule.Italic.Value ? 1 : 0;
        if (!String.IsNullOrWhiteSpace(rule.FontName)) find.Replacement.Font.Name = rule.FontName;
        if (rule.FontSize.HasValue) find.Replacement.Font.Size = rule.FontSize.Value;
        if (!String.IsNullOrWhiteSpace(rule.ColorHex) && Regex.IsMatch(rule.ColorHex, "^[0-9a-fA-F]{6}$"))
        {
            var red = Convert.ToInt32(rule.ColorHex.Substring(0, 2), 16); var green = Convert.ToInt32(rule.ColorHex.Substring(2, 2), 16); var blue = Convert.ToInt32(rule.ColorHex.Substring(4, 2), 16);
            find.Replacement.Font.Color = (blue << 16) | (green << 8) | red;
        }
        if (!ExecuteReplaceAll(find) && !rule.UseWildcards && !rule.HasFormatting)
        {
            if (selectionOnly) ReplaceLiteralRangesInRange(document, trackingRange.Duplicate, rule.Find, rule.Replace ?? String.Empty);
            else ReplaceLiteralRanges(document, rule.Find, rule.Replace ?? String.Empty);
        }
    }

    private static void CollapseRepeatedSpacesInDocument(dynamic document)
    {
        for (var pass = 0; pass < 30; pass++)
        {
            dynamic range = document.Content;
            dynamic find = range.Find;
            find.ClearFormatting(); find.Replacement.ClearFormatting();
            find.Text = "  "; find.Replacement.Text = " ";
            find.Forward = true; find.Wrap = 1; find.Format = false; find.MatchWildcards = false;
            var changed = ExecuteReplaceAll(find);
            if (!changed) changed = ReplaceLiteralRanges(document, "  ", " ") > 0;
            if (!changed) break;
        }
    }

    private static bool ExecuteReplaceAll(dynamic find)
    {
        try { return Convert.ToBoolean(find.Execute(Replace: 2)); }
        catch
        {
            try
            {
                var missing = Type.Missing;
                return Convert.ToBoolean(find.Execute(missing, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing, 2, missing, missing, missing, missing));
            }
            catch { return false; }
        }
    }

    public int MarkInlineLatex()
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic content = document.Content;
        var source = content.Text as string ?? String.Empty;
        var matches = FindAiFormulaCandidates(source);
        var marked = 0;
        foreach (var match in matches)
        {
            try
            {
                dynamic range = document.Range((int)content.Start + match.Index, (int)content.Start + match.Index + match.Length);
                range.HighlightColorIndex = 7; // 黄色
                marked++;
            }
            catch { }
        }
        return marked;
    }

    public void ApplyTextColor(string hex)
    {
        ApplyTextColor(hex, "Document");
    }

    public void ApplyTextColor(string hex, string scope)
    {
        var wordColor = ParseWordColor(hex);
        dynamic document = ((dynamic)_application).ActiveDocument;
        dynamic range = document.Content;
        if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase))
        {
            dynamic selected = ((dynamic)_application).Selection.Range;
            if ((int)selected.End <= (int)selected.Start) throw new InvalidOperationException("请先选中需要处理的文字。");
            range = selected;
        }
        range.Font.Color = wordColor;
    }

    public int ApplyFormulaFont(string fontName)
    {
        return ApplyFormulaFont(fontName, "Document");
    }

    public int ApplyFormulaFont(string fontName, string scope)
    {
        dynamic document = ((dynamic)_application).ActiveDocument;
        var count = 0;
        if (String.IsNullOrWhiteSpace(fontName)) fontName = "Cambria Math";
        var scopeStart = (int)document.Content.Start;
        var scopeEnd = (int)document.Content.End;
        if (String.Equals(scope, "Selection", StringComparison.OrdinalIgnoreCase))
        {
            dynamic selected = ((dynamic)_application).Selection.Range;
            scopeStart = (int)selected.Start;
            scopeEnd = (int)selected.End;
        }
        for (var index = 1; index <= SafeCollectionCount(document.OMaths); index++)
        {
            try
            {
                dynamic math = document.OMaths[index];
                if (!MathRangesOverlap(math, scopeStart, scopeEnd)) continue;
                math.Range.Font.Name = fontName;
                math.Range.Font.NameFarEast = fontName;
                count++;
            }
            catch { }
        }
        return count;
    }
}

public sealed class TextReplacement
{
    public TextReplacement(string find, string replace) { Find = find; Replace = replace; }
    public string Find { get; private set; }
    public string Replace { get; private set; }
    public bool UseWildcards { get; set; }
    public bool? Bold { get; set; }
    public bool? Italic { get; set; }
    public string FontName { get; set; }
    public float? FontSize { get; set; }
    public string ColorHex { get; set; }
    public bool HasFormatting { get { return Bold.HasValue || Italic.HasValue || !String.IsNullOrWhiteSpace(FontName) || FontSize.HasValue || !String.IsNullOrWhiteSpace(ColorHex); } }

    public static TextReplacement Parse(string line)
    {
        if (String.IsNullOrWhiteSpace(line)) return null;
        var sections = line.Split('|');
        var pair = sections[0].Split(new[] { "=>" }, StringSplitOptions.None);
        if (pair.Length != 2 || String.IsNullOrWhiteSpace(pair[0])) return null;
        var rule = new TextReplacement(pair[0].Trim(), pair[1].Trim());
        foreach (var raw in sections.Skip(1))
        {
            var option = raw.Trim();
            if (option.Equals("wildcard", StringComparison.OrdinalIgnoreCase) || option == "通配符") rule.UseWildcards = true;
            else if (option.Equals("bold", StringComparison.OrdinalIgnoreCase) || option == "加粗") rule.Bold = true;
            else if (option.Equals("italic", StringComparison.OrdinalIgnoreCase) || option == "斜体") rule.Italic = true;
            else if (option.StartsWith("font=", StringComparison.OrdinalIgnoreCase)) rule.FontName = option.Substring(5).Trim();
            else if (option.StartsWith("size=", StringComparison.OrdinalIgnoreCase)) { float size; if (Single.TryParse(option.Substring(5), out size)) rule.FontSize = size; }
            else if (option.StartsWith("color=", StringComparison.OrdinalIgnoreCase)) rule.ColorHex = option.Substring(6).Trim().TrimStart('#');
        }
        return rule;
    }
}

public sealed class FormulaPreviewItem
{
    public int Start { get; set; }
    public int Length { get; set; }
    public string Source { get; set; }
    public string Formula { get; set; }
    public string Preview { get; set; }
    public bool Display { get; set; }
}
