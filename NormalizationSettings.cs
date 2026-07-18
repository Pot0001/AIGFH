using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace AIGFH;

public sealed class NormalizationSettings
{
    private static readonly string FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AIGFH", "normalization.ini");
    private static readonly string LegacyFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OfflineOfficeAddIn", "normalization.ini");

    public bool AutoPaste { get; set; } = false;
    public bool ApplyPresetAfterNormalize { get; set; } = false;
    public bool RemoveMarkdownHeadings { get; set; } = true;
    public bool NormalizePlainTextWithoutMarkers { get; set; } = false;
    public bool ThreeLineTables { get; set; } = false;
    public bool AutoFitTables { get; set; } = true;
    public bool DeepseekStyleTables { get; set; } = true;
    public bool AutoRunFormula { get; set; } = true;
    public bool AutoRunConvert { get; set; } = true;
    public bool AutoRunText { get; set; } = true;
    public bool AutoRunTables { get; set; } = true;
    public bool AutoRunSpaces { get; set; } = false;
    public bool AutoRunFont { get; set; } = false;
    public string ProcessingScope { get; set; } = "Document";
    public float LineSpacing { get; set; } = 1.5F;
    public float ParagraphSpaceAfter { get; set; } = 6F;
    public bool FirstLineIndent { get; set; } = true;
    public int AnswerLineCount { get; set; } = 4;
    public int AnswerLineLength { get; set; } = 32;
    public float ExamMarginCm { get; set; } = 1.6F;
    public int ChoiceColumns { get; set; } = 4;
    public string FormulaMode { get; set; } = "Word";
    public string InlinePrefix { get; set; } = "\\(";
    public string InlineSuffix { get; set; } = "\\)";
    public string DisplayPrefix { get; set; } = "\\[";
    public string DisplaySuffix { get; set; } = "\\]";
    public string MultiPrefix { get; set; } = "\\begin{align*}";
    public string MultiSuffix { get; set; } = "\\end{align*}";
    public string FontName { get; set; } = "宋体";
    public string FormulaFontName { get; set; } = "Cambria Math";
    public float FontSize { get; set; } = 12F;
    public string RulesText { get; set; } = "AI生成 => AI 生成\r\n（ => (\r\n） => )";

    public static NormalizationSettings Load()
    {
        var value = new NormalizationSettings();
        var sourcePath = File.Exists(FilePath) ? FilePath : LegacyFilePath;
        if (!File.Exists(sourcePath)) return value;
        foreach (var line in File.ReadAllLines(sourcePath, Encoding.UTF8))
        {
            var separator = line.IndexOf('=');
            if (separator < 1) continue;
            var key = line.Substring(0, separator);
            var data = Decode(line.Substring(separator + 1));
            bool flag; float number;
            switch (key)
            {
                case "AutoPaste": if (Boolean.TryParse(data, out flag)) value.AutoPaste = flag; break;
                case "ApplyPreset": if (Boolean.TryParse(data, out flag)) value.ApplyPresetAfterNormalize = flag; break;
                case "RemoveHeadings": if (Boolean.TryParse(data, out flag)) value.RemoveMarkdownHeadings = flag; break;
                case "NormalizePlainText": if (Boolean.TryParse(data, out flag)) value.NormalizePlainTextWithoutMarkers = flag; break;
                case "ThreeLine": if (Boolean.TryParse(data, out flag)) value.ThreeLineTables = flag; break;
                case "AutoFit": if (Boolean.TryParse(data, out flag)) value.AutoFitTables = flag; break;
                case "DeepseekStyle": if (Boolean.TryParse(data, out flag)) value.DeepseekStyleTables = flag; break;
                case "AutoRunFormula": if (Boolean.TryParse(data, out flag)) value.AutoRunFormula = flag; break;
                case "AutoRunConvert": if (Boolean.TryParse(data, out flag)) value.AutoRunConvert = flag; break;
                case "AutoRunText": if (Boolean.TryParse(data, out flag)) value.AutoRunText = flag; break;
                case "AutoRunTables": if (Boolean.TryParse(data, out flag)) value.AutoRunTables = flag; break;
                case "AutoRunSpaces": if (Boolean.TryParse(data, out flag)) value.AutoRunSpaces = flag; break;
                case "AutoRunFont": if (Boolean.TryParse(data, out flag)) value.AutoRunFont = flag; break;
                case "ProcessingScope": value.ProcessingScope = data; break;
                case "LineSpacing": if (Single.TryParse(data, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) value.LineSpacing = number; break;
                case "ParagraphSpaceAfter": if (Single.TryParse(data, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) value.ParagraphSpaceAfter = number; break;
                case "FirstLineIndent": if (Boolean.TryParse(data, out flag)) value.FirstLineIndent = flag; break;
                case "AnswerLineCount": if (Int32.TryParse(data, out var integer)) value.AnswerLineCount = integer; break;
                case "AnswerLineLength": if (Int32.TryParse(data, out integer)) value.AnswerLineLength = integer; break;
                case "ExamMarginCm": if (Single.TryParse(data, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) value.ExamMarginCm = number; break;
                case "ChoiceColumns": if (Int32.TryParse(data, out integer)) value.ChoiceColumns = integer; break;
                case "FormulaMode": value.FormulaMode = data; break;
                case "InlinePrefix": value.InlinePrefix = data; break;
                case "InlineSuffix": value.InlineSuffix = data; break;
                case "DisplayPrefix": value.DisplayPrefix = data; break;
                case "DisplaySuffix": value.DisplaySuffix = data; break;
                case "MultiPrefix": value.MultiPrefix = data; break;
                case "MultiSuffix": value.MultiSuffix = data; break;
                case "FontName": value.FontName = data; break;
                case "FormulaFontName": value.FormulaFontName = data; break;
                case "FontSize": if (Single.TryParse(data, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) value.FontSize = number; break;
                case "Rules": value.RulesText = data; break;
            }
        }
        // 兼容早期版本遗留的空自动任务配置。全部关闭会让“一键规范”
        // 看起来没有反应，新版本恢复为最常用的文本、公式和表格组合。
        if (!value.AutoRunText && !value.AutoRunFormula && !value.AutoRunConvert &&
            !value.AutoRunTables && !value.AutoRunSpaces && !value.AutoRunFont &&
            !value.ApplyPresetAfterNormalize)
        {
            value.AutoRunText = true;
            value.AutoRunFormula = true;
            value.AutoRunConvert = true;
            value.AutoRunTables = true;
        }
        // 1.1.0 起由功能区的统一范围按钮控制。早期的 Auto 值按全文处理，
        // 避免用户未选中文字时得到难以预期的范围。
        if (!String.Equals(value.ProcessingScope, "Selection", StringComparison.OrdinalIgnoreCase) &&
            !String.Equals(value.ProcessingScope, "Document", StringComparison.OrdinalIgnoreCase))
            value.ProcessingScope = "Document";
        return value;
    }

    public void Save()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
        var lines = new List<string>
        {
            Pair("AutoPaste", AutoPaste.ToString()), Pair("ApplyPreset", ApplyPresetAfterNormalize.ToString()),
            Pair("RemoveHeadings", RemoveMarkdownHeadings.ToString()), Pair("NormalizePlainText", NormalizePlainTextWithoutMarkers.ToString()), Pair("ThreeLine", ThreeLineTables.ToString()), Pair("AutoFit", AutoFitTables.ToString()),
            Pair("DeepseekStyle", DeepseekStyleTables.ToString()), Pair("AutoRunFormula", AutoRunFormula.ToString()), Pair("AutoRunConvert", AutoRunConvert.ToString()),
            Pair("AutoRunText", AutoRunText.ToString()), Pair("AutoRunTables", AutoRunTables.ToString()), Pair("AutoRunSpaces", AutoRunSpaces.ToString()), Pair("AutoRunFont", AutoRunFont.ToString()),
            Pair("ProcessingScope", ProcessingScope), Pair("LineSpacing", LineSpacing.ToString(CultureInfo.InvariantCulture)), Pair("ParagraphSpaceAfter", ParagraphSpaceAfter.ToString(CultureInfo.InvariantCulture)),
            Pair("FirstLineIndent", FirstLineIndent.ToString()), Pair("AnswerLineCount", AnswerLineCount.ToString(CultureInfo.InvariantCulture)), Pair("AnswerLineLength", AnswerLineLength.ToString(CultureInfo.InvariantCulture)),
            Pair("ExamMarginCm", ExamMarginCm.ToString(CultureInfo.InvariantCulture)), Pair("ChoiceColumns", ChoiceColumns.ToString(CultureInfo.InvariantCulture)),
            Pair("FormulaMode", FormulaMode), Pair("InlinePrefix", InlinePrefix), Pair("InlineSuffix", InlineSuffix),
            Pair("DisplayPrefix", DisplayPrefix), Pair("DisplaySuffix", DisplaySuffix), Pair("MultiPrefix", MultiPrefix), Pair("MultiSuffix", MultiSuffix),
            Pair("FontName", FontName), Pair("FormulaFontName", FormulaFontName), Pair("FontSize", FontSize.ToString(CultureInfo.InvariantCulture)), Pair("Rules", RulesText)
        };
        File.WriteAllLines(FilePath, lines, Encoding.UTF8);
    }

    private static string Pair(string key, string value) { return key + "=" + Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? String.Empty)); }
    private static string Decode(string value) { try { return Encoding.UTF8.GetString(Convert.FromBase64String(value)); } catch { return String.Empty; } }
}
