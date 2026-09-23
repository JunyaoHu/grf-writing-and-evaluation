# GRF Writing and Evaluation

面向香港 RGC General Research Fund 申请书的 Codex 技能：起草、修订、模拟评审，以及可编辑 PPTX 基金图绘制。

这是独立辅助工具，不代表 RGC 或 UGC，不保证获资助。评分框架属于自评方法，不是官方评分表。

## 安装

将本目录复制到 `$CODEX_HOME/skills/grf-writing-and-evaluation/`；未设置 CODEX_HOME 时，使用用户主目录下的 `.codex/skills/grf-writing-and-evaluation/`。确认该目录直接包含 `SKILL.md`，避免多嵌套一层。重新打开会话后检查技能列表。

纯写作和评审不需要安装 Python 依赖。请向技能提供申请书、目标年度、受评范围，以及对应年度官方指南或其本地路径。

## 从零开始

在空项目目录中准备一份研究构想，说明研究问题、已有证据、拟采用方法和目标申请年度，再提供对应年度官方指南。可以先请求提纲，随后逐节起草；没有前期结果时应明确标注，不用示例补造。下列调用和最小绘图脚本均为通用示例，不需要预先存在的申请书或领域素材；修改与评审示例则需提供相应稿件。

```text
$grf-writing-and-evaluation 根据我提供的研究构想，从零规划一份 GRF 提案。先给问题—目标—任务—验证提纲，列出仍需补充的事实，不虚构已有成果。
```

## 调用示例

```text
$grf-writing-and-evaluation 起草 Research Context，依据现有材料，不虚构前期结果。
$grf-writing-and-evaluation 只评审 Task 2 和相关附图，不修改、不编译。
$grf-writing-and-evaluation 根据评审意见修改指定源文件，不编译申请书。
$grf-writing-and-evaluation 绘制可编辑 PPTX 技术路线图，并导出第一页图 PDF。
```

绘图为按需加载的内置二级模块，无需额外安装绘图技能。局部改图默认保留其他设计，文字、箭头与图形保持可编辑。只讨论或评图不会自动修改文件。

## 绘图依赖与能力边界

可选 Python 辅助库依赖 `python-pptx` 和 `lxml`：

```sh
python -m pip install -r requirements-figures.txt
python examples/minimal_figure.py --output /your/output/example.pptx
```

示例仅生成虚构流程的 PPTX，不代表研究结果，也不导出 PDF。它不会覆盖现有输出文件。

库中的 PNG 导出函数额外需要 Windows、已安装的桌面 PowerPoint 和 `pywin32`。PDF 导出需要运行环境另行提供 PowerPoint 自动化或其他可用导出工具；本包不含完整 PDF 导出、像素恢复和压缩程序。450 ppi / JPEG 75 是验证目标，不能仅靠普通 PowerPoint 导出保证。无法完成导出或视觉检查时，技能须报告实际完成范围。

默认设置见 [绘图默认值](references/figures/defaults.md)，可由用户请求或项目 `AGENTS.md` 覆盖。修改正文默认不编译 LaTeX，也不启动后台构建。

## 目录

- `SKILL.md`：技能入口和按需路由。
- `agents/openai.yaml`：技能显示信息。
- `references/`：写作、评审、指南来源与绘图流程。
- `scripts/pptx_figure_kit.py`：可选 PPTX 绘图辅助库。
- `examples/minimal_figure.py`：无真实申请内容的最小示例。

## 指南与版本

本包不分发完整官方指南或中文译文。查阅 [官方来源与版本规则](references/guidance-sources.md)，使用对应申请年度原文；历史摘要不构成当前合规结论。模拟评审不替代所属大学审核或官方判断。

## 发布与许可状态

当前为发布准备稿，尚未选定并授予开源许可证。公开可见不等于已获得开源使用授权。原创部分、改编参考和第三方指南的来源边界见 [来源与许可记录](THIRD_PARTY_NOTICES.md)。发布前须确认相关来源并添加正式 LICENSE，不应把第三方材料统一归入仓库自选许可证。

仅发布本技能目录，不上传申请书项目、素材、操作备份、凭据或本地指南。
