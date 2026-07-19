# AI规范化

> **免费插件**：Word/WPS 文本、TeX 公式、表格与试卷排版工具。项目更新、下载与问题反馈见 [GitHub 项目主页](https://github.com/Pot0001/AIGFH)。

![Version](https://img.shields.io/badge/version-1.1.3-2563eb)
![Platform](https://img.shields.io/badge/platform-Windows-0078d4)
![Office](https://img.shields.io/badge/Office-Word%20%7C%20WPS-d83b01)

AI规范化是 Microsoft Word 与 WPS Office 的本地 COM 加载项，重点帮助教师整理 AI 备课内容、讲义、试题和试卷，也适合学生整理笔记与作业。

## 常用场景

### 教师

- **AI 备课**：点击“复制 AI 提示词”，粘贴到 AI 并填写任务；将生成结果粘贴到 Word/WPS，选择“一键规范”，再用“公式核对”快速检查。
- **制作讲义**：使用“讲义版式”统一字体、段落和标题；表格与公式可单独整理。
- **制作试卷**：选择 A3/A4 与横版/竖版，整理题号、选项、小问、表格和公式。
- **制作学生版**：在“答案解析”中生成学生版副本，原教师版保持不变；还可隐藏或显示答案解析。
- **课堂展示**：将文档逐页转为图片并生成 PPT，保持 Word 页面原貌。
- **交付前检查**：用“排版检查”统计待转换 TeX、残留标记和公式定界符问题。

### 学生

- 整理 AI 笔记、知识总结和解题过程。
- 把标准 TeX 转为可编辑公式，或把 Word/WPS 公式转回 TeX。
- 使用讲义版式、答题区和公式核对整理学习材料。

## 主要功能

- **一键规范**：按自定义设置一次完成文本、公式、表格和版式整理。
- **处理范围**：用功能区的“范围：全文/选区”统一控制一键规范、文本、公式、TeX/MathType 互转、表格、清理和讲义版式。
- **规范文本**：整理 Markdown、HTML、完整 TeX 文档骨架、标题、段落、列表和表格。普通文本无结构标识时默认保持原样。
- **规范公式**：将 `$...$`、`$$...$$`、`\(...\)`、`\[...\]` 和常用 TeX 环境转换为 Office 原生公式；重复执行会跳过已有公式。
- **TeX 双向转换**：标准 TeX 与 Word/WPS 可编辑公式互转。
- **多来源兼容**：适配 ChatGPT、DeepSeek、豆包、Kimi 及常见网页复制内容，并修正常见残缺分隔符、紧凑分式和多余定界符。
- **试卷排版**：提供 A3/A4、横版/竖版四种入口，可自定义页边距、选择题列数和答题区。
- **教师工具**：答案解析隐藏/显示、学生版副本、排版检查、文档转 PPT。
- **表格处理**：Markdown 转表格，并统一现有表格的边框、字体和列宽。
- **MathType 互转**：MathType/Equation OLE 与可编辑公式互转；转为 MathType 需要已安装 MathType 加载项。
- **自定义**：设置正文字体、公式字体、字号、行距、段距、首行缩进、试卷、答题区、表格和一键规范流程。
- **替换规则**：编辑、导入和导出文本替换规则。
- **使用说明**：功能区直接打开操作说明；“更多”菜单可一键复制推荐 AI 提示词。
- **本地处理**：文档处理在本机完成，插件本身不上传文档内容。

## 数学表达覆盖

覆盖小学、初中、高中、大学基础课程与考研常用的数学表达和排版结构：

- **数与代数**：分数、根式、绝对值、方程与不等式、函数、数列、指数、对数、三角函数、复数。
- **集合与逻辑**：集合关系、交并补、映射、量词、命题逻辑、常用离散数学符号。
- **几何与向量**：平面与立体几何、解析几何、圆锥曲线、向量、角度、上下标和几何标记。
- **微积分与分析**：极限、导数、偏导、梯度、积分、重积分、曲线积分、级数、常微分方程及优化问题常用符号。
- **线性代数**：向量、矩阵、行列式、分块结构、秩、线性空间、特征值与常用算子。
- **概率统计**：排列组合、概率、条件概率、期望、方差、分布、随机过程、估计与检验常用表达。
- **离散与数论**：命题逻辑、组合、图论、同余、整除、模运算和常用离散结构。
- **多行结构**：分段函数、方程组、推导对齐、矩阵、行列式、数组、上下限、重音与括号伸缩。

插件负责识别和排版数学表达，不判断题目、推导或答案是否正确。

## 安装

1. 下载并双击 `AI规范化-Pot0001-1.1.3.exe`。
2. 安装程序会显示本机版本与安装包版本，并给出适合当前状态的操作。
3. 完成后关闭全部 Word/WPS 窗口并重新打开。
4. 在功能区进入 **AI规范化**。

卸载方式：再次打开安装包并选择“卸载”，或在 Windows“已安装的应用”中卸载。

## 快速使用

1. 在功能区打开“使用说明”，或从“更多”菜单复制 AI 提示词；生成内容后粘贴到 Word/WPS。
2. 点击“范围：全文/选区”切换当前处理范围；使用“选区”前请先选中内容。
3. 常规内容用“一键规范”；需要分步处理时，先“规范文本”，再“规范公式”。
4. 试卷使用“试卷排版”；交付前使用“公式核对”和“排版检查”。

完整操作与给 AI 的推荐提示词见 [`使用说明.txt`](使用说明.txt)。

## 支持的 TeX

- 分隔符：`$...$`、`$$...$$`、`\(...\)`、`\[...\]`
- 分式与根式：`\frac`、`\dfrac`、`\tfrac`、`\sqrt`、`\binom`
- 多行结构：`\begin{align}`、`\begin{aligned}`、`\begin{gather}`、`\begin{cases}`、`\begin{matrix}`、`\begin{array}`
- 字体与文本：`\text`、`\mathrm`、`\mathbf`、`\mathbb`、`\boldsymbol`
- 关系与逻辑：集合、量词、关系、箭头、逻辑和同余等常用命令
- 运算与修饰：积分、求和、极限、常用函数、上下标、重音、上下括号、上下标注

同时兼容 `\dfrac1a`、`\frac12` 等常见紧凑写法，并修正网页输出中的 `\lt`、`\gt`。复杂公式转换后可用“公式核对”快速确认。

## 运行环境

- Windows 10/11
- .NET Framework 4.8
- Microsoft Word 2016/2019/2021/Microsoft 365，或支持 COM 加载项的 WPS Office Windows 桌面版
- x86 与 x64 Office/WPS
- Word 公式转 MathType 时需安装并启用 MathType Office 加载项

安装包已内嵌所需 COM 接口，用户端无需单独安装 Office PIA。

## 构建

- Visual Studio 2022 或 Build Tools 2022
- .NET Framework 4.8 Developer Pack

正式发布文件：

```text
bin/Release/AIGFH.dll
dist/AI规范化-Pot0001-1.1.3.exe
```

## 项目结构

```text
AiWebFormatter.cs            Markdown/HTML 网页样式转换
Connect.cs                   COM 加载项入口与功能区界面
OfficeDocumentService.cs     Word/WPS 文档处理主逻辑
OmmlMathBuilder.cs           TeX 到 Office Math/OMML 转换
OmmlToLatexConverter.cs      Office Math/UnicodeMath 到标准 TeX 转换
NormalizationSettings.cs     用户设置持久化
SettingsForm.cs              自定义设置界面
RuleEditorForm.cs            替换规则编辑与导入导出
FormulaReviewForm.cs         公式转换前预览核对
UpdateChecker.cs             新版本检查
AI提示词.txt                 可由功能区直接复制的推荐提示词
使用说明.txt                 操作流程与格式说明
Installer/Program.cs         当前用户安装与维护程序
Resources/RibbonIcons/       功能区图标
References/                  构建时使用的 Office COM 接口
Samples/                     AI 内容与试卷排版示例
```

## TODO

- [x] 可视化规则编辑器与规则导入/导出
- [x] 扩展 TeX 宏、复杂表格和常用文档环境
- [x] 公式转换前预览与逐项核对
- [x] Word/WPS 兼容性检查与差异处理
- [x] 教师版/学生版工作流与排版检查

## 许可

本项目采用 [Apache License 2.0](LICENSE)，并附加 [Commons Clause License Condition v1.0](COMMONS-CLAUSE)。源码可查看、修改和分发；不得销售以本软件功能为全部或主要价值的产品或服务。
