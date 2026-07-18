using System;
using System.Drawing;
using System.Windows.Forms;

namespace AIGFH;

internal sealed class SettingsForm : Form
{
    private readonly NormalizationSettings _settings;
    private readonly TextBox _font = new TextBox();
    private readonly TextBox _formulaFont = new TextBox();
    private readonly NumericUpDown _fontSize = Number(8, 36, 12, 1);
    private readonly NumericUpDown _lineSpacing = Number(1, 3, 1.5M, 0.1M);
    private readonly NumericUpDown _spaceAfter = Number(0, 30, 6, 1);
    private readonly NumericUpDown _margin = Number(0.8M, 3.5M, 1.6M, 0.1M);
    private readonly NumericUpDown _choiceColumns = Number(1, 4, 4, 1);
    private readonly NumericUpDown _answerLines = Number(1, 20, 4, 1);
    private readonly NumericUpDown _answerLength = Number(8, 80, 32, 1);
    private readonly CheckBox _autoPaste = Check("空文档时自动粘贴剪贴板内容（默认关闭）");
    private readonly CheckBox _plainText = Check("普通文字也统一版式（默认关闭）");
    private readonly CheckBox _firstIndent = Check("正文首行缩进");
    private readonly CheckBox _autoFit = Check("表格自适应列宽");
    private readonly CheckBox _threeLine = Check("表格使用三线表样式");
    private readonly CheckBox _runText = Check("整理文本结构和版式");
    private readonly CheckBox _runFormula = Check("转换并规范公式");
    private readonly CheckBox _runTables = Check("还原并整理表格");
    private readonly CheckBox _runSpaces = Check("清理多余空格和空行");
    private readonly CheckBox _runFont = Check("统一公式字体");
    private readonly CheckBox _runLayout = Check("应用讲义版式");

    internal SettingsForm(NormalizationSettings settings)
    {
        _settings = settings ?? NormalizationSettings.Load();
        Text = "AI规范化设置";
        Font = new Font("Microsoft YaHei UI", 9F);
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(600, 570);
        Size = new Size(680, 640);
        BackColor = Color.FromArgb(246, 248, 252);

        var tabs = new TabControl { Dock = DockStyle.Fill, Padding = new Point(18, 6) };
        tabs.TabPages.Add(MakePage("常规", _autoPaste, _plainText,
            Note("普通文字选项关闭时，只整理带 Markdown、HTML 或 TeX 标记的内容，避免改动已排好的文档。")));
        tabs.TabPages.Add(MakePage("文档版式", Row("正文字体", _font), Row("正文字号", _fontSize), Row("公式字体", _formulaFont),
            Row("行距倍数", _lineSpacing), Row("段后间距（磅）", _spaceAfter), _firstIndent));
        tabs.TabPages.Add(MakePage("试卷", Row("页边距（厘米）", _margin), Row("选择题每行选项数", _choiceColumns),
            Row("答题区行数", _answerLines), Row("答题线长度", _answerLength)));
        tabs.TabPages.Add(MakePage("表格", _autoFit, _threeLine));
        tabs.TabPages.Add(MakePage("一键规范", _runText, _runFormula, _runTables, _runSpaces, _runFont, _runLayout));

        var save = new Button { Text = "保存设置", Width = 112, Height = 36, BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
        save.FlatAppearance.BorderSize = 0;
        save.Click += (_, __) => SaveAndClose();
        var reset = new Button { Text = "恢复默认", Width = 100, Height = 36, FlatStyle = FlatStyle.Flat };
        reset.Click += (_, __) => LoadValues(new NormalizationSettings());
        var close = new Button { Text = "取消", Width = 90, Height = 36, FlatStyle = FlatStyle.Flat };
        close.Click += (_, __) => Close();
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 58, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(12, 10, 12, 8) };
        buttons.Controls.Add(save); buttons.Controls.Add(close); buttons.Controls.Add(reset);
        Controls.Add(tabs); Controls.Add(buttons);
        LoadValues(_settings);
    }

    private static TabPage MakePage(string title, params Control[] controls)
    {
        var page = new TabPage(title) { BackColor = Color.White, Padding = new Padding(18) };
        var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };
        foreach (var control in controls) flow.Controls.Add(control);
        page.Controls.Add(flow);
        return page;
    }

    private static Control Row(string label, Control editor)
    {
        var panel = new Panel { Width = 570, Height = 42, Margin = new Padding(0, 2, 0, 2) };
        panel.Controls.Add(new Label { Text = label, AutoSize = false, Width = 205, Height = 30, TextAlign = ContentAlignment.MiddleLeft, Location = new Point(2, 6) });
        editor.Width = 300; editor.Height = 28; editor.Location = new Point(215, 7);
        if (editor is ComboBox combo) combo.DropDownStyle = ComboBoxStyle.DropDownList;
        panel.Controls.Add(editor);
        return panel;
    }

    private static CheckBox Check(string text) => new CheckBox { Text = text, AutoSize = true, Margin = new Padding(3, 8, 3, 8) };
    private static Label Note(string text) => new Label { Text = text, AutoSize = false, Width = 560, Height = 48, ForeColor = Color.FromArgb(71, 85, 105), Margin = new Padding(3, 8, 3, 8) };
    private static NumericUpDown Number(decimal min, decimal max, decimal value, decimal step) => new NumericUpDown { Minimum = min, Maximum = max, Value = value, Increment = step, DecimalPlaces = step < 1 ? 1 : 0 };

    private void LoadValues(NormalizationSettings value)
    {
        _font.Text = value.FontName; _fontSize.Value = Clamp(_fontSize, (decimal)value.FontSize);
        _formulaFont.Text = String.IsNullOrWhiteSpace(value.FormulaFontName) ? "Cambria Math" : value.FormulaFontName;
        _lineSpacing.Value = Clamp(_lineSpacing, (decimal)value.LineSpacing); _spaceAfter.Value = Clamp(_spaceAfter, (decimal)value.ParagraphSpaceAfter);
        _margin.Value = Clamp(_margin, (decimal)value.ExamMarginCm); _choiceColumns.Value = Clamp(_choiceColumns, value.ChoiceColumns);
        _answerLines.Value = Clamp(_answerLines, value.AnswerLineCount); _answerLength.Value = Clamp(_answerLength, value.AnswerLineLength);
        _autoPaste.Checked = value.AutoPaste; _plainText.Checked = value.NormalizePlainTextWithoutMarkers;
        _firstIndent.Checked = value.FirstLineIndent; _autoFit.Checked = value.AutoFitTables; _threeLine.Checked = value.ThreeLineTables;
        _runText.Checked = value.AutoRunText; _runFormula.Checked = value.AutoRunFormula; _runTables.Checked = value.AutoRunTables; _runSpaces.Checked = value.AutoRunSpaces; _runFont.Checked = value.AutoRunFont; _runLayout.Checked = value.ApplyPresetAfterNormalize;
    }

    private static decimal Clamp(NumericUpDown field, decimal value) => Math.Max(field.Minimum, Math.Min(field.Maximum, value));

    private void SaveAndClose()
    {
        _settings.FontName = String.IsNullOrWhiteSpace(_font.Text) ? "宋体" : _font.Text.Trim(); _settings.FontSize = (float)_fontSize.Value;
        _settings.FormulaFontName = String.IsNullOrWhiteSpace(_formulaFont.Text) ? "Cambria Math" : _formulaFont.Text.Trim();
        _settings.LineSpacing = (float)_lineSpacing.Value; _settings.ParagraphSpaceAfter = (float)_spaceAfter.Value; _settings.FirstLineIndent = _firstIndent.Checked;
        _settings.ExamMarginCm = (float)_margin.Value; _settings.ChoiceColumns = (int)_choiceColumns.Value; _settings.AnswerLineCount = (int)_answerLines.Value; _settings.AnswerLineLength = (int)_answerLength.Value;
        _settings.AutoPaste = _autoPaste.Checked; _settings.NormalizePlainTextWithoutMarkers = _plainText.Checked;
        _settings.AutoFitTables = _autoFit.Checked; _settings.ThreeLineTables = _threeLine.Checked;
        _settings.AutoRunText = _runText.Checked; _settings.AutoRunFormula = _runFormula.Checked; _settings.AutoRunConvert = _runFormula.Checked;
        _settings.AutoRunTables = _runTables.Checked; _settings.AutoRunSpaces = _runSpaces.Checked; _settings.AutoRunFont = _runFont.Checked;
        _settings.ApplyPresetAfterNormalize = _runLayout.Checked;
        _settings.Save(); DialogResult = DialogResult.OK; Close();
    }
}
