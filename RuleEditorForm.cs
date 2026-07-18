using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AIGFH;

internal sealed class RuleEditorForm : Form
{
    private readonly NormalizationSettings _settings;
    private readonly OfficeDocumentService _service;
    private readonly DataGridView _grid = new DataGridView();

    internal RuleEditorForm(NormalizationSettings settings, OfficeDocumentService service)
    {
        _settings = settings ?? NormalizationSettings.Load();
        _service = service;
        Text = "文本替换规则";
        Font = new Font("Microsoft YaHei UI", 9F);
        StartPosition = FormStartPosition.CenterScreen;
        Size = new Size(920, 570);
        MinimumSize = new Size(760, 470);

        var help = new Label
        {
            Dock = DockStyle.Top,
            Height = 48,
            Padding = new Padding(12, 8, 12, 4),
            Text = "规则按表格顺序执行；“保存并应用”使用功能区当前的全文/选区范围。"
        };
        ConfigureGrid();
        LoadRules(_settings.RulesText);

        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 56, Padding = new Padding(10, 9, 10, 8), FlowDirection = FlowDirection.RightToLeft };
        buttons.Controls.Add(MakeButton("关闭", (_, __) => Close()));
        buttons.Controls.Add(MakeButton("保存", (_, __) => SaveRules(false)));
        buttons.Controls.Add(MakeButton("保存并应用", (_, __) => SaveRules(true), true));
        buttons.Controls.Add(MakeButton("导出", (_, __) => ExportRules()));
        buttons.Controls.Add(MakeButton("导入", (_, __) => ImportRules()));
        buttons.Controls.Add(MakeButton("删除所选", (_, __) => DeleteSelected()));
        buttons.Controls.Add(MakeButton("新增规则", (_, __) => _grid.Rows.Add()));

        Controls.Add(_grid);
        Controls.Add(help);
        Controls.Add(buttons);
    }

    private void ConfigureGrid()
    {
        _grid.Dock = DockStyle.Fill;
        _grid.BackgroundColor = Color.White;
        _grid.BorderStyle = BorderStyle.None;
        _grid.AllowUserToAddRows = true;
        _grid.AllowUserToDeleteRows = true;
        _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Find", HeaderText = "查找内容", FillWeight = 150 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Replace", HeaderText = "替换为", FillWeight = 150 });
        _grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Wildcard", HeaderText = "通配符", FillWeight = 55 });
        _grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Bold", HeaderText = "粗体", FillWeight = 45 });
        _grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Italic", HeaderText = "斜体", FillWeight = 45 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Font", HeaderText = "字体", FillWeight = 75 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Size", HeaderText = "字号", FillWeight = 50 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Color", HeaderText = "颜色(RRGGBB)", FillWeight = 80 });
    }

    private static Button MakeButton(string text, EventHandler action, bool primary = false)
    {
        var button = new Button { Text = text, AutoSize = true, Height = 34, FlatStyle = FlatStyle.Flat, BackColor = primary ? Color.FromArgb(37, 99, 235) : Color.White, ForeColor = primary ? Color.White : Color.FromArgb(30, 41, 59) };
        button.Click += action;
        return button;
    }

    private void LoadRules(string text)
    {
        _grid.Rows.Clear();
        foreach (var line in (text ?? String.Empty).Replace("\r\n", "\n").Split('\n'))
        {
            var rule = TextReplacement.Parse(line);
            if (rule == null) continue;
            _grid.Rows.Add(rule.Find, rule.Replace, rule.UseWildcards, rule.Bold == true, rule.Italic == true,
                rule.FontName ?? String.Empty, rule.FontSize.HasValue ? rule.FontSize.Value.ToString("0.##") : String.Empty,
                rule.ColorHex ?? String.Empty);
        }
    }

    private string SerializeRules()
    {
        var lines = _grid.Rows.Cast<DataGridViewRow>().Where(row => !row.IsNewRow)
            .Select(row =>
            {
                var find = Convert.ToString(row.Cells["Find"].Value).Trim();
                if (find.Length == 0) return null;
                var replace = Convert.ToString(row.Cells["Replace"].Value);
                var options = new System.Collections.Generic.List<string>();
                if (Convert.ToBoolean(row.Cells["Wildcard"].Value ?? false)) options.Add("wildcard");
                if (Convert.ToBoolean(row.Cells["Bold"].Value ?? false)) options.Add("bold");
                if (Convert.ToBoolean(row.Cells["Italic"].Value ?? false)) options.Add("italic");
                var font = Convert.ToString(row.Cells["Font"].Value).Trim(); if (font.Length > 0) options.Add("font=" + font);
                var size = Convert.ToString(row.Cells["Size"].Value).Trim(); if (size.Length > 0) options.Add("size=" + size);
                var color = Convert.ToString(row.Cells["Color"].Value).Trim().TrimStart('#'); if (color.Length > 0) options.Add("color=" + color);
                return find + " => " + replace + (options.Count == 0 ? String.Empty : " | " + String.Join(" | ", options));
            }).Where(line => line != null);
        return String.Join("\r\n", lines);
    }

    private void SaveRules(bool apply)
    {
        _settings.RulesText = SerializeRules();
        _settings.Save();
        var count = 0;
        if (apply && _service != null)
        {
            foreach (var line in _settings.RulesText.Replace("\r\n", "\n").Split('\n'))
            {
                var rule = TextReplacement.Parse(line);
                if (rule == null) continue;
                _service.ApplyRule(rule, _settings.ProcessingScope);
                count++;
            }
        }
        MessageBox.Show(apply ? "已保存并应用 " + count + " 条规则。" : "规则已保存。", "AI规范化", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void ImportRules()
    {
        using (var dialog = new OpenFileDialog { Filter = "规则文本 (*.txt)|*.txt|所有文件 (*.*)|*.*", Title = "导入替换规则" })
            if (dialog.ShowDialog(this) == DialogResult.OK) LoadRules(File.ReadAllText(dialog.FileName, Encoding.UTF8));
    }

    private void ExportRules()
    {
        using (var dialog = new SaveFileDialog { Filter = "规则文本 (*.txt)|*.txt", FileName = "AI规范化-替换规则.txt", Title = "导出替换规则" })
            if (dialog.ShowDialog(this) == DialogResult.OK) File.WriteAllText(dialog.FileName, SerializeRules(), new UTF8Encoding(true));
    }

    private void DeleteSelected()
    {
        foreach (DataGridViewRow row in _grid.SelectedRows.Cast<DataGridViewRow>().Where(row => !row.IsNewRow).ToArray()) _grid.Rows.Remove(row);
    }
}
