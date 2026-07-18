using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace AIGFH;

internal sealed class FormulaReviewForm : Form
{
    private readonly OfficeDocumentService _service;
    private readonly IList<FormulaPreviewItem> _items;
    private readonly DataGridView _grid = new DataGridView();

    internal FormulaReviewForm(OfficeDocumentService service, NormalizationSettings settings)
    {
        _service = service;
        _items = service.GetFormulaPreview(settings);
        Text = "公式核对";
        Font = new Font("Microsoft YaHei UI", 9F);
        StartPosition = FormStartPosition.CenterScreen;
        Size = new Size(980, 620);
        MinimumSize = new Size(760, 480);

        var summary = new Label { Dock = DockStyle.Top, Height = 46, Padding = new Padding(12, 10, 12, 4), Text = "共发现 " + _items.Count + " 个公式。取消勾选可保留原文。" };
        ConfigureGrid();
        foreach (var item in _items) _grid.Rows.Add(true, item.Display ? "行间" : "行内", item.Formula, item.Preview);

        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 56, Padding = new Padding(10, 9, 10, 8), FlowDirection = FlowDirection.RightToLeft };
        buttons.Controls.Add(Button("取消", (_, __) => Close()));
        buttons.Controls.Add(Button("转换勾选公式", (_, __) => ConvertSelected(), true));
        buttons.Controls.Add(Button("全部取消", (_, __) => SetAll(false)));
        buttons.Controls.Add(Button("全部选择", (_, __) => SetAll(true)));
        Controls.Add(_grid); Controls.Add(summary); Controls.Add(buttons);
    }

    private void ConfigureGrid()
    {
        _grid.Dock = DockStyle.Fill;
        _grid.BackgroundColor = Color.White;
        _grid.BorderStyle = BorderStyle.None;
        _grid.AllowUserToAddRows = false;
        _grid.AllowUserToDeleteRows = false;
        _grid.RowHeadersVisible = false;
        _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        _grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Selected", HeaderText = "转换", FillWeight = 35 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Type", HeaderText = "类型", FillWeight = 40, ReadOnly = true });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Source", HeaderText = "TeX 源码", FillWeight = 210, ReadOnly = true });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Preview", HeaderText = "转换后预览", FillWeight = 180, ReadOnly = true });
    }

    private static Button Button(string text, EventHandler action, bool primary = false)
    {
        var button = new Button { Text = text, AutoSize = true, Height = 34, FlatStyle = FlatStyle.Flat, BackColor = primary ? Color.FromArgb(37, 99, 235) : Color.White, ForeColor = primary ? Color.White : Color.FromArgb(30, 41, 59) };
        button.Click += action; return button;
    }

    private void SetAll(bool value)
    {
        foreach (DataGridViewRow row in _grid.Rows) row.Cells["Selected"].Value = value;
    }

    private void ConvertSelected()
    {
        _grid.EndEdit();
        var selected = new List<FormulaPreviewItem>();
        for (var index = 0; index < _items.Count; index++)
            if (Convert.ToBoolean(_grid.Rows[index].Cells["Selected"].Value ?? false)) selected.Add(_items[index]);
        var count = _service.ConvertFormulaPreview(selected);
        MessageBox.Show("已转换 " + count + " 个经核对的公式。", "AI规范化", MessageBoxButtons.OK, MessageBoxIcon.Information);
        DialogResult = DialogResult.OK;
        Close();
    }
}
