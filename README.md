# 研究汇报与 PPT 分步制作

`research-presentation` 是一个把论文、案例、行业、政策或项目材料转成研究汇报的 Codex skill。可以从资料研读开始，也可以直接整理现有 Excel 底稿、调整某页预览，或根据已有 PPT 写逐页讲稿。

## 可以从哪一步开始

| 模块 | 典型交付 |
| --- | --- |
| M01 资料与内容 | 论证与证据梳理，或对照正文审查已有 PPT |
| M02 简明大纲 | 按受众和时长组织的逐页内容表 |
| M03 图片预览 | 整套缩略图、逐页高清图或指定页的改版图 |
| M04 预览定稿 | 最终逐页图、总览和采用版本清单 |
| M05 可编辑 PPT | PowerPoint 源文件、PDF、实际渲染图和编辑说明 |
| M06 演讲稿 | 与实际页序对应的逐页讲稿、计时和删减方案 |
| M07 整理交付 | 采用版本、参考文献和使用说明 |
| M08 Excel 底稿 | 数据工作簿、Excel 母图和图表对应信息 |

新建、导出或重建可编辑 PPT 前，先确认全套最终页面预览及采用版本，再明确下达制作指令。已有 PPT 可直接做只读检查。完整规则见 [SKILL.md](SKILL.md)。

## 安装与调用

将完整的 `research-presentation` 目录克隆或解压到项目的 `.agents/skills/` 下，使入口位于 `.agents/skills/research-presentation/SKILL.md`。请保留同级的参考文档、脚本和素材目录。打开该项目后，可以在 Codex 对话中输入：

```text
用 $research-presentation 按正文和模板做一版 18 分钟的汇报大纲。
用 $research-presentation 只重做第 6 页图片预览，保留现有数据和标题。
用 $research-presentation 对照这份 PPT 写逐页讲稿，总时长控制在 12 分钟。
```

## 仓库内容

- [`SKILL.md`](SKILL.md) 定义模块、输入输出及预览确认流程。
- [`references/`](references/) 提供逐模块细则、实证论文读表、视觉准则和版式说明。
- [`assets/workpapers/`](assets/workpapers/) 提供四份可直接复制的 Excel 底稿模板，覆盖公司多模块、周期跟踪、驱动测算和共享数据与专题作图。选用方式见 [底稿模板选用](references/workpaper-layouts.md#可直接复用的excel模板)。
- [`assets/presentation-layouts/`](assets/presentation-layouts/) 和 [`assets/changxin-purple/`](assets/changxin-purple/) 保存构图、样式观察和来源摘要。发布副本收录文字和结构记录，原始案例截图及本机文件路径留在本地版本中。
- [`scripts/`](scripts/) 提供 Office 文件结构检查、候选预览图清点及采用版本拼图。[`agents/openai.yaml`](agents/openai.yaml) 提供 Codex 界面元数据。

## 可选工具

`inspect_ooxml.py` 使用 Python 标准库读取 PPTX 和 XLSX 的保存结构。`preview_pack.py` 用于登记候选图和拼接已采用的图片，使用默认字体拼图时需要 Pillow 10.1 或更新版本。实际编辑 Office 文件和检查渲染需要当前环境支持的相应工具。脚本命令和参数见 [素材与效率脚本](references/efficiency-tools.md)。

```sh
python -B scripts/inspect_ooxml.py --help
python -B scripts/preview_pack.py --help
```
