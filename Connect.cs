using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;

namespace AIGFH;

[ComVisible(true)]
[Guid("A53C7691-E736-45A0-8A8A-36BFD2D8EC03")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IRibbonCallbacks
{
    [DispId(1)]
    void OnRibbonLoad([MarshalAs(UnmanagedType.IDispatch)] object ribbon);

    [DispId(2)]
    void OnRibbonButton([MarshalAs(UnmanagedType.IDispatch)] object control);

    [DispId(3)]
    [return: MarshalAs(UnmanagedType.IDispatch)]
    object GetRibbonImage([MarshalAs(UnmanagedType.IDispatch)] object control);

    [DispId(4)]
    void OnScopeToggle([MarshalAs(UnmanagedType.IDispatch)] object control, bool pressed);

    [DispId(5)]
    string GetScopeLabel([MarshalAs(UnmanagedType.IDispatch)] object control);

    [DispId(6)]
    bool GetScopePressed([MarshalAs(UnmanagedType.IDispatch)] object control);
}

[ComVisible(true)]
[Guid("2B868EC7-70D7-469B-A8AA-9B30F47EB33F")]
[ProgId("AIGFH.Connect")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(IRibbonCallbacks))]
public sealed class Connect : IRibbonCallbacks, IDTExtensibility2, IRibbonExtensibility
{
    private object _application;
    private object _ribbon;

    public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
    {
        _application = application;
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
    {
        _ribbon = null;
        _application = null;
    }

    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) { }

    public string GetCustomUI(string ribbonId)
    {
        return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab id=""AIGFH.Tab"" label=""AI规范化"">
        <group id=""AIGFH.Core"" label=""常用"">
          <toggleButton id=""ScopeToggle"" getLabel=""GetScopeLabel"" getPressed=""GetScopePressed"" screentip=""切换处理范围"" supertip=""“选区”只处理已选内容；“全文”处理整篇文档。用于规范、公式互转、讲义版式、表格和清理。"" getImage=""GetRibbonImage"" size=""large"" onAction=""OnScopeToggle"" />
          <button id=""RunAll"" label=""一键规范"" screentip=""一次完成常用整理"" supertip=""按当前范围和自定义设置整理文本、公式、表格和版式。"" getImage=""GetRibbonImage"" size=""large"" onAction=""OnRibbonButton"" />
          <button id=""NormalizeAllAiFormula"" label=""规范公式"" screentip=""转换当前范围内的 TeX"" getImage=""GetRibbonImage"" size=""large"" onAction=""OnRibbonButton"" />
          <button id=""NormalizeAllAiText"" label=""规范文本"" screentip=""整理当前范围内的文本"" getImage=""GetRibbonImage"" size=""large"" onAction=""OnRibbonButton"" />
          <menu id=""NormalizeExamPaperMenu"" label=""试卷排版"" screentip=""选择纸张与方向"" supertip=""始终整理整篇文档，并应用所选纸张与方向。"" getImage=""GetRibbonImage"" size=""large"" itemSize=""large"">
            <button id=""NormalizeExamA3Landscape"" label=""A3 横版"" getImage=""GetRibbonImage"" onAction=""OnRibbonButton"" />
            <button id=""NormalizeExamA3Portrait"" label=""A3 竖版"" getImage=""GetRibbonImage"" onAction=""OnRibbonButton"" />
            <button id=""NormalizeExamA4Landscape"" label=""A4 横版"" getImage=""GetRibbonImage"" onAction=""OnRibbonButton"" />
            <button id=""NormalizeExamA4Portrait"" label=""A4 竖版"" getImage=""GetRibbonImage"" onAction=""OnRibbonButton"" />
          </menu>
        </group>
        <group id=""AIGFH.Preparation"" label=""备课"">
          <button id=""ApplyLayout"" label=""讲义版式"" screentip=""统一正文与段落版式"" getImage=""GetRibbonImage"" size=""large"" onAction=""OnRibbonButton"" />
          <button id=""InsertAnswerArea"" label=""插入答题区"" getImage=""GetRibbonImage"" onAction=""OnRibbonButton"" />
          <menu id=""HideAnswers"" label=""答案解析"" screentip=""制作教师版和学生版"" getImage=""GetRibbonImage"" size=""large"" itemSize=""large"">
            <button id=""HideAnswerSections"" label=""隐藏答案解析"" onAction=""OnRibbonButton"" />
            <button id=""ShowAnswerSections"" label=""显示答案解析"" onAction=""OnRibbonButton"" />
            <button id=""CreateStudentCopy"" label=""生成学生版副本"" screentip=""保留当前教师版"" supertip=""另存一个副本并移除答案、解析和评分标准。"" onAction=""OnRibbonButton"" />
          </menu>
          <button id=""GeneratePpt"" label=""文档转 PPT"" screentip=""按页面生成图片幻灯片"" supertip=""每个文档页面生成一张图片幻灯片；完成后请保存演示文稿。"" getImage=""GetRibbonImage"" onAction=""OnRibbonButton"" />
        </group>
        <group id=""AIGFH.FormulaTools"" label=""公式"">
          <button id=""FormulaReview"" label=""公式核对"" screentip=""转换前逐项核对"" supertip=""按当前范围列出待转换公式；取消勾选可保留原文。"" getImage=""GetRibbonImage"" size=""large"" onAction=""OnRibbonButton"" />
          <menu id=""TextToMath"" label=""TeX 互转"" screentip=""TeX 与可编辑公式互转"" supertip=""按左侧“范围”设置处理选区或全文。"" getImage=""GetRibbonImage"" itemSize=""large"">
            <button id=""ConvertTexToMath"" label=""TeX 转可编辑公式"" onAction=""OnRibbonButton"" />
            <button id=""ConvertMathToTex"" label=""可编辑公式转 TeX"" onAction=""OnRibbonButton"" />
          </menu>
          <menu id=""MathTypeToWord"" label=""MathType 互转"" screentip=""在 MathType 与可编辑公式之间转换"" supertip=""按左侧“范围”设置处理选区或全文。"" getImage=""GetRibbonImage"" itemSize=""large"">
            <button id=""MathTypeToOffice"" label=""MathType 转可编辑公式"" onAction=""OnRibbonButton"" />
            <button id=""OfficeToMathType"" label=""可编辑公式转 MathType"" onAction=""OnRibbonButton"" />
          </menu>
          <menu id=""FloatMath"" label=""公式位置"" screentip=""修正公式与文字的对齐"" supertip=""按左侧“范围”设置处理选区或全文。"" getImage=""GetRibbonImage"" itemSize=""large"">
            <button id=""RaiseEquations"" label=""公式上移 3 磅"" onAction=""OnRibbonButton"" />
            <button id=""ResetEquationPosition"" label=""恢复默认位置"" onAction=""OnRibbonButton"" />
          </menu>
        </group>
        <group id=""AIGFH.Tables"" label=""表格"">
          <menu id=""RestoreTables"" label=""表格处理"" screentip=""还原或整理表格"" supertip=""按左侧“范围”设置处理选区或全文。"" getImage=""GetRibbonImage"" size=""large"" itemSize=""large"">
            <button id=""RestoreMarkdownTables"" label=""Markdown 转表格"" onAction=""OnRibbonButton"" />
            <button id=""CleanExistingTables"" label=""整理现有表格"" onAction=""OnRibbonButton"" />
          </menu>
        </group>
        <group id=""AIGFH.Cleanup"" label=""文档清理"">
          <button id=""CollapseSpaces"" label=""清理空格与空行"" getImage=""GetRibbonImage"" size=""large"" onAction=""OnRibbonButton"" />
          <button id=""BlackText"" label=""统一为黑色"" screentip=""处理当前范围"" getImage=""GetRibbonImage"" onAction=""OnRibbonButton"" />
        </group>
        <group id=""AIGFH.Settings"" label=""设置"">
          <button id=""OpenSettings"" label=""自定义"" getImage=""GetRibbonImage"" size=""large"" onAction=""OnRibbonButton"" />
          <button id=""OpenUserGuide"" label=""使用说明"" screentip=""查看操作方法和 AI 提示词"" supertip=""打开使用说明，查看常用流程、公式格式和推荐提示词。"" getImage=""GetRibbonImage"" size=""large"" onAction=""OnRibbonButton"" />
          <button id=""DocumentCheck"" label=""排版检查"" screentip=""检查当前文档"" supertip=""统计待转换公式、排版标记和公式定界符问题。"" getImage=""GetRibbonImage"" onAction=""OnRibbonButton"" />
          <menu id=""MoreSettings"" label=""更多"" screentip=""规则、运行检查与帮助"" getImage=""GetRibbonImage"" itemSize=""large"">
            <button id=""OpenRuleEditor"" label=""替换规则"" onAction=""OnRibbonButton"" />
            <button id=""CompatibilityCheck"" label=""运行检查"" screentip=""检查 Word/WPS 功能"" onAction=""OnRibbonButton"" />
            <button id=""CheckUpdates"" label=""检查更新"" onAction=""OnRibbonButton"" />
            <button id=""CopyAiPrompt"" label=""复制 AI 提示词"" screentip=""复制推荐提示词"" supertip=""复制后粘贴到 AI，并在末尾填写具体任务。"" onAction=""OnRibbonButton"" />
            <button id=""OpenProjectHome"" label=""关于与反馈"" onAction=""OnRibbonButton"" />
          </menu>
          <labelControl id=""VersionLabel"" label=""版本：1.1.3"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }

    public void OnRibbonLoad(object ribbon)
    {
        _ribbon = ribbon;
    }

    public void OnScopeToggle(object control, bool pressed)
    {
        var settings = NormalizationSettings.Load();
        settings.ProcessingScope = pressed ? "Selection" : "Document";
        settings.Save();
        try { ((dynamic)_ribbon).InvalidateControl("ScopeToggle"); } catch { }
        SetStatus(pressed ? "处理范围：选区。" : "处理范围：全文。");
    }

    public string GetScopeLabel(object control)
    {
        return GetScopePressed(control) ? "范围：选区" : "范围：全文";
    }

    public bool GetScopePressed(object control)
    {
        return String.Equals(NormalizationSettings.Load().ProcessingScope, "Selection", StringComparison.OrdinalIgnoreCase);
    }

    public object GetRibbonImage(object control)
    {
        try
        {
            var controlId = GetRibbonControlId(control);
            return RibbonPictureConverter.ToPictureDisp(RibbonIconProvider.Load(controlId));
        }
        catch
        {
            return null;
        }
    }

    public void OnRibbonButton(object control)
    {
        var service = new OfficeDocumentService(_application);
        var settings = NormalizationSettings.Load();
        try
        {
            var controlId = GetRibbonControlId(control);
            switch (controlId)
            {
                case "NormalizeAllAiFormula": SetFormulaStatus(service.NormalizeUnifiedAiFormulas(settings), ScopeName(settings)); break;
                case "NormalizeAllAiText": SetStatus(service.NormalizeUnifiedAiText(settings) > 0 ? ScopeName(settings) + "文本已规范。" : ScopeName(settings) + "没有可处理文本。"); break;
                case "NormalizeFormulaSelection": settings.ProcessingScope = "Selection"; SetFormulaStatus(service.NormalizeUnifiedAiFormulas(settings), "选中内容"); break;
                case "NormalizeFormulaDocument": settings.ProcessingScope = "Document"; SetFormulaStatus(service.NormalizeUnifiedAiFormulas(settings), "整篇文档"); break;
                case "NormalizeTextSelection": settings.ProcessingScope = "Selection"; SetStatus(service.NormalizeUnifiedAiText(settings) > 0 ? "选区文本规范完成。" : "选区没有可处理文本。"); break;
                case "NormalizeTextDocument": settings.ProcessingScope = "Document"; SetStatus(service.NormalizeUnifiedAiText(settings) > 0 ? "全文文本规范完成。" : "文档没有可处理文本。"); break;
                case "NormalizeExamA3Landscape": SetStatus("A3 横版已应用；转换 " + service.NormalizeExamPaper(settings, 6, true) + " 个公式。"); break;
                case "NormalizeExamA3Portrait": SetStatus("A3 竖版已应用；转换 " + service.NormalizeExamPaper(settings, 6, false) + " 个公式。"); break;
                case "NormalizeExamA4Landscape": SetStatus("A4 横版已应用；转换 " + service.NormalizeExamPaper(settings, 7, true) + " 个公式。"); break;
                case "NormalizeExamA4Portrait": SetStatus("A4 竖版已应用；转换 " + service.NormalizeExamPaper(settings, 7, false) + " 个公式。"); break;
                case "ConvertMathToTex":
                case "MathToMarkdown":
                {
                    var count = service.ConvertOfficeMathToMarkdown(settings.ProcessingScope);
                    SetStatus(count > 0 ? ScopeName(settings) + "已转为标准 TeX，共 " + count + " 个公式。" : ScopeName(settings) + "没有可编辑公式。");
                    break;
                }
                case "ConvertTexToMath":
                case "TextToMath":
                {
                    var count = service.ConvertTextToOfficeMath(settings.ProcessingScope);
                    SetStatus(count > 0 ? ScopeName(settings) + "已转为可编辑公式，共 " + count + " 个。" : ScopeName(settings) + "没有可转换的 TeX 公式。");
                    break;
                }
                case "FormulaReview": using (var form = new FormulaReviewForm(service, settings)) form.ShowDialog(); break;
                case "MathTypeToOffice":
                {
                    var count = service.ConvertMathTypeToOfficeMath(settings.ProcessingScope);
                    SetStatus(count > 0 ? ScopeName(settings) + "MathType 转换完成，共 " + count + " 个公式。" : ScopeName(settings) + "没有可转换的 MathType 公式。");
                    break;
                }
                case "OfficeToMathType":
                {
                    var count = service.ConvertOfficeMathToMathType(settings.ProcessingScope);
                    SetStatus(count == -2 ? ScopeName(settings) + "没有可编辑公式。" : count < 0 ? "需要安装并启用 MathType 加载项。" : count == 0 ? "已打开 MathType 转换窗口，请按提示完成。" : ScopeName(settings) + "已转为 MathType 公式，共 " + count + " 个。");
                    break;
                }
                case "RaiseEquations": SetStatus(ScopeName(settings) + "已上移 " + service.FloatEquations(settings.ProcessingScope) + " 个公式。"); break;
                case "ResetEquationPosition": SetStatus(ScopeName(settings) + "已恢复 " + service.ResetEquationPosition(settings.ProcessingScope) + " 个公式的位置。"); break;
                case "RestoreMarkdownTables": SetStatus(ScopeName(settings) + "已还原 " + service.RestoreMarkdownTables(settings.ThreeLineTables, settings.AutoFitTables, false, settings.ProcessingScope) + " 个表格。"); break;
                case "CleanExistingTables": SetStatus(ScopeName(settings) + "已整理 " + service.SmartCleanTables(settings.AutoFitTables, settings.ThreeLineTables, settings.ProcessingScope) + " 个单元格。"); break;
                case "CollapseSpaces": service.CollapseRepeatedSpaces(settings.ProcessingScope); service.NormalizeBlankLines(settings.ProcessingScope); SetStatus(ScopeName(settings) + "空格与空行已清理。"); break;
                case "ApplyLayout": service.ApplyPreset(settings.FontName, settings.FontSize, 3, settings.FirstLineIndent, settings.ParagraphSpaceAfter > 0, settings.LineSpacing, settings.ParagraphSpaceAfter, settings.ProcessingScope); service.ApplyFormulaFont(settings.FormulaFontName, settings.ProcessingScope); SetStatus(ScopeName(settings) + "讲义版式已应用。"); break;
                case "InsertAnswerArea": service.InsertAnswerArea(settings); SetStatus("已插入答题区。"); break;
                case "HideAnswerSections": SetStatus("已隐藏 " + service.SetAnswerSectionsHidden(true) + " 个答案或解析段落。"); break;
                case "ShowAnswerSections": SetStatus("已显示 " + service.SetAnswerSectionsHidden(false) + " 个答案或解析段落。"); break;
                case "CreateStudentCopy":
                {
                    var result = service.CreateStudentCopy();
                    SetStatus("学生版已生成：" + Path.GetFileName(result.Item1) + "；移除 " + result.Item2 + " 个答案或解析段落。");
                    break;
                }
                case "OpenSettings": using (var form = new SettingsForm(settings)) if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK) SetStatus("自定义设置已保存。"); break;
                case "OpenUserGuide": OpenUserGuide(); break;
                case "CopyAiPrompt": CopyAiPrompt(); SetStatus("AI 提示词已复制，请粘贴到 AI 并在末尾填写任务。"); break;
                case "OpenRuleEditor": using (var form = new RuleEditorForm(settings, service)) form.ShowDialog(); break;
                case "CompatibilityCheck": ShowTextReport("运行检查", service.GetCompatibilityReport()); break;
                case "DocumentCheck": ShowTextReport("排版检查", service.GetDocumentCheckReport()); break;
                case "CheckUpdates": UpdateChecker.CheckAndShow(); break;
                case "OpenProjectHome": ShowProjectHome(); break;
                case "RunAll":
                case "RunAllSelection":
                case "RunAllDocument":
                {
                    if (controlId == "RunAllSelection") settings.ProcessingScope = "Selection";
                    else if (controlId == "RunAllDocument") settings.ProcessingScope = "Document";
                    if (settings.ProcessingScope == "Selection" && !service.HasSelectedText()) throw new InvalidOperationException("请先选中需要规范的内容。");
                    var textChanged = 0;
                    var formulas = 0;
                    var tables = 0;
                    var skipped = new List<string>();
                    if (settings.AutoRunText) RunAutoStep("文本", () => textChanged = service.NormalizeUnifiedAiText(settings), skipped);
                    if (settings.AutoRunFormula || settings.AutoRunConvert)
                        RunAutoStep("公式", () => formulas += service.NormalizeUnifiedAiFormulas(settings), skipped);
                    if (settings.AutoRunTables) RunAutoStep("表格", () => { tables += service.RestoreMarkdownTables(settings.ThreeLineTables, settings.AutoFitTables, false, settings.ProcessingScope); tables += service.SmartCleanTables(settings.AutoFitTables, settings.ThreeLineTables, settings.ProcessingScope); }, skipped);
                    if (settings.AutoRunSpaces) RunAutoStep("空格与空行", () => { service.CollapseRepeatedSpaces(settings.ProcessingScope); service.NormalizeBlankLines(settings.ProcessingScope); }, skipped);
                    if (settings.AutoRunFont) RunAutoStep("公式字体", () => service.ApplyFormulaFont(settings.FormulaFontName, settings.ProcessingScope), skipped);
                    if (settings.ApplyPresetAfterNormalize) RunAutoStep("讲义版式", () => { service.ApplyPreset(settings.FontName, settings.FontSize, 3, settings.FirstLineIndent, settings.ParagraphSpaceAfter > 0, settings.LineSpacing, settings.ParagraphSpaceAfter, settings.ProcessingScope); service.ApplyFormulaFont(settings.FormulaFontName, settings.ProcessingScope); }, skipped);
                    var result = ScopeName(settings) + "规范完成。" + (textChanged > 0 ? "文本已整理，" : String.Empty) + "公式 " + formulas + " 个，表格 " + tables + " 项。";
                    SetStatus(skipped.Count == 0
                        ? result
                        : result + " 未完成：" + String.Join("、", skipped) + "。请运行检查。");
                    break;
                }
                case "BlackText": service.ApplyTextColor("000000", settings.ProcessingScope); SetStatus(ScopeName(settings) + "文字已统一为黑色。"); break;
                case "BlackSelection": service.ApplyTextColor("000000", "Selection"); SetStatus("选中文字已统一为黑色。"); break;
                case "BlackDocument": service.ApplyTextColor("000000", "Document"); SetStatus("全文文字已统一为黑色。"); break;
                case "GeneratePpt": SetStatus("已生成 " + service.GeneratePowerPoint() + " 页图片幻灯片，请保存演示文稿。"); break;
            }
        }
        catch (Exception error)
        {
            System.Windows.Forms.MessageBox.Show(error.Message, "AI规范化", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
        }
    }

    private static string GetRibbonControlId(object control)
    {
        try { return Convert.ToString(((dynamic)control).Id); } catch { }
        try { return Convert.ToString(((dynamic)control).ID); } catch { }
        return String.Empty;
    }

    private static void RunAutoStep(string name, Action action, ICollection<string> skipped)
    {
        try { action(); }
        catch { skipped.Add(name); }
    }

    private static string ScopeName(NormalizationSettings settings)
    {
        return settings != null && String.Equals(settings.ProcessingScope, "Selection", StringComparison.OrdinalIgnoreCase)
            ? "选区："
            : "全文：";
    }

    private void SetStatus(string message)
    {
        try { ((dynamic)_application).StatusBar = message; } catch { }
    }

    private void SetFormulaStatus(int count, string scope)
    {
        SetStatus(count > 0
            ? scope + "公式已规范，共 " + count + " 个。"
            : scope + "未发现可转换的 TeX 公式。");
    }

    private static void ShowProjectHome()
    {
        var url = UpdateChecker.RepositoryUrl;
        using (var form = new System.Windows.Forms.Form())
        {
            form.Text = "关于 AI规范化";
            form.Font = new Font("Microsoft YaHei UI", 9F);
            form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            form.MaximizeBox = false;
            form.MinimizeBox = false;
            form.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            form.ClientSize = new Size(560, 300);
            form.BackColor = Color.White;

            var layout = new System.Windows.Forms.TableLayoutPanel
            {
                Dock = System.Windows.Forms.DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 7,
                Padding = new System.Windows.Forms.Padding(32, 24, 32, 20),
                BackColor = Color.White
            };
            layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            for (var index = 0; index < 6; index++)
                layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));

            var title = new System.Windows.Forms.Label
            {
                Text = "AI规范化",
                Font = new Font("Microsoft YaHei UI", 20F, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 64, 175),
                AutoSize = true,
                Margin = new System.Windows.Forms.Padding(0, 0, 0, 8)
            };
            var subtitle = new System.Windows.Forms.Label
            {
                Text = "Word / WPS 文本、公式、表格与试卷排版工具",
                Font = new Font("Microsoft YaHei UI", 10F),
                ForeColor = Color.FromArgb(51, 65, 85),
                AutoSize = true,
                Margin = new System.Windows.Forms.Padding(1, 0, 0, 12)
            };
            var version = new System.Windows.Forms.Label
            {
                Text = "版本 " + UpdateChecker.CurrentVersion + "    免费使用",
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 64, 175),
                AutoSize = true,
                Margin = new System.Windows.Forms.Padding(1, 0, 0, 18)
            };
            var description = new System.Windows.Forms.Label
            {
                Text = "项目介绍、下载、更新与问题反馈",
                ForeColor = Color.FromArgb(71, 85, 105),
                AutoSize = true,
                Margin = new System.Windows.Forms.Padding(1, 0, 0, 4)
            };
            var link = new System.Windows.Forms.LinkLabel
            {
                Text = url,
                AutoSize = true,
                LinkColor = Color.FromArgb(37, 99, 235),
                ActiveLinkColor = Color.FromArgb(30, 64, 175),
                Margin = new System.Windows.Forms.Padding(1, 0, 0, 0)
            };
            link.LinkClicked += (_, __) => OpenUrl(url);

            var buttonBar = new System.Windows.Forms.FlowLayoutPanel
            {
                Dock = System.Windows.Forms.DockStyle.Fill,
                FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new System.Windows.Forms.Padding(0, 14, 0, 0),
                Margin = new System.Windows.Forms.Padding(0)
            };
            var close = new System.Windows.Forms.Button
            {
                Text = "关闭",
                Width = 108,
                Height = 36,
                BackColor = Color.FromArgb(37, 99, 235),
                ForeColor = Color.White,
                FlatStyle = System.Windows.Forms.FlatStyle.Flat,
                DialogResult = System.Windows.Forms.DialogResult.Cancel,
                Margin = new System.Windows.Forms.Padding(0)
            };
            close.FlatAppearance.BorderSize = 0;
            close.Click += (_, __) => form.Close();
            buttonBar.Controls.Add(close);

            layout.Controls.Add(title, 0, 0);
            layout.Controls.Add(subtitle, 0, 1);
            layout.Controls.Add(version, 0, 2);
            layout.Controls.Add(description, 0, 3);
            layout.Controls.Add(link, 0, 4);
            layout.Controls.Add(new System.Windows.Forms.Panel { Height = 1, Dock = System.Windows.Forms.DockStyle.Top }, 0, 5);
            layout.Controls.Add(buttonBar, 0, 6);
            form.Controls.Add(layout);
            form.AcceptButton = close;
            form.CancelButton = close;
            form.ShowDialog();
        }
    }

    private static void OpenUserGuide()
    {
        var directory = Path.GetDirectoryName(typeof(Connect).Assembly.Location) ?? String.Empty;
        var guidePath = Path.Combine(directory, "使用说明.txt");
        if (!File.Exists(guidePath))
        {
            var installedPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AIGFH",
                UpdateChecker.CurrentVersion,
                "使用说明.txt");
            if (File.Exists(installedPath)) guidePath = installedPath;
        }

        if (!File.Exists(guidePath))
            throw new FileNotFoundException("使用说明尚未安装，请重新运行当前版本安装包。", guidePath);
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(guidePath) { UseShellExecute = true });
    }

    private static void CopyAiPrompt()
    {
        var directory = Path.GetDirectoryName(typeof(Connect).Assembly.Location) ?? String.Empty;
        var promptPath = Path.Combine(directory, "AI提示词.txt");
        if (!File.Exists(promptPath))
            throw new FileNotFoundException("AI 提示词尚未安装，请重新运行当前版本安装包。", promptPath);
        var prompt = File.ReadAllText(promptPath, System.Text.Encoding.UTF8).Trim();
        if (prompt.Length == 0) throw new InvalidDataException("AI 提示词内容为空，请重新运行当前版本安装包。");
        System.Windows.Forms.Clipboard.SetText(prompt);
    }

    private static void ShowTextReport(string title, string report)
    {
        using (var form = new System.Windows.Forms.Form())
        {
            form.Text = title;
            form.Font = new Font("Microsoft YaHei UI", 9F);
            form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            form.Size = new Size(610, 440);
            form.MinimumSize = new Size(500, 360);
            var textBox = new System.Windows.Forms.TextBox { Multiline = true, ReadOnly = true, ScrollBars = System.Windows.Forms.ScrollBars.Both, Dock = System.Windows.Forms.DockStyle.Fill, Text = report, BackColor = Color.White };
            var copy = new System.Windows.Forms.Button { Text = "复制报告", Width = 100, Height = 34, Dock = System.Windows.Forms.DockStyle.Right };
            copy.Click += (_, __) => { try { System.Windows.Forms.Clipboard.SetText(report ?? String.Empty); } catch { } };
            var close = new System.Windows.Forms.Button { Text = "关闭", Width = 90, Height = 34, Dock = System.Windows.Forms.DockStyle.Right };
            close.Click += (_, __) => form.Close();
            var bottom = new System.Windows.Forms.Panel { Dock = System.Windows.Forms.DockStyle.Bottom, Height = 50, Padding = new System.Windows.Forms.Padding(8) };
            bottom.Controls.Add(close); bottom.Controls.Add(copy);
            form.Controls.Add(textBox); form.Controls.Add(bottom);
            form.ShowDialog();
        }
    }

    private static void OpenUrl(string url)
    {
        try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url) { UseShellExecute = true }); }
        catch { System.Windows.Forms.Clipboard.SetText(url); }
    }


    private void RemoveLegacyToolbar()
    {
        try
        {
            dynamic application = _application;
            dynamic commandBars = application.CommandBars;
            try { commandBars["AI Document Normalizer"].Delete(); } catch { }
            try { commandBars["离线文档工具"].Delete(); } catch { }
        }
        catch { }
    }
}

internal static class RibbonIconProvider
{
    internal static Image Load(string controlId)
    {
        var resourceName = "AIGFH.Resources.RibbonIcons." + controlId + ".png";
        using (var stream = typeof(RibbonIconProvider).Assembly.GetManifestResourceStream(resourceName))
        {
            if (stream == null) return null;
            using (var image = Image.FromStream(stream))
            {
                return new Bitmap(image);
            }
        }
    }
}

internal sealed class RibbonPictureConverter : System.Windows.Forms.AxHost
{
    private RibbonPictureConverter() : base(String.Empty) { }

    internal static object ToPictureDisp(Image image)
    {
        return image == null ? null : GetIPictureDispFromPicture(image);
    }
}
