
# GRF 基金图：可编辑 PPTX 绘制与修改

Create or revise editable academic figures. For an existing PPTX, default to a local edit, not a redesign.

默认值及覆盖顺序见 [绘图默认值](defaults.md)。下文示例样式服从该配置与用户要求。

## Runtime and scope

This is an instruction skill for a skill-capable host, including Codex; it is not tied to a specific GPT model. Resolve bundled paths relative to this reference file. User instructions and project AGENTS.md take precedence.

Use tools actually available in the session. Reuse a working project helper first; otherwise choose available presentation tools, PowerPoint automation, or targeted OOXML editing. The Python kit is optional: it needs python-pptx and lxml; its PNG exporter additionally needs Windows, desktop PowerPoint and pywin32. Check a real Python executable before relying on the Windows python alias or a stale py launcher entry. Do not install or reconfigure tools just to load this skill.

Keep text, arrows and diagram geometry as editable PowerPoint objects. Photographs may remain raster images; do not flatten the entire figure to satisfy a PPTX request. If export or rendering is unavailable, report the precise missing capability and distinguish edited source from verified PDF/preview.

## Incremental editing: default path

1. Use the latest working PPTX. Identify the requested delta, affected shapes and output path from session context. Do not ask again for established preferences or authorization.
2. Resume from a compact project state and the latest requested delta. Reuse the current preview, shape map and working helper; load supporting references only when relevant, and do not reread unchanged guidance already available in context. Read this reference once per context; inspect only relevant XML fields or script functions. Read [incremental-editing.md](incremental-editing.md) only when setting up or repairing a sustained editing workflow.
3. Separate confirmed edits from questions explicitly left for discussion. Complete confirmed edits; inspect the manuscript before resolving terminology or taxonomy questions. For input/output correspondence and grouped examples, read [framework-content-and-layout.md](framework-content-and-layout.md). Apply all requested changes in one coherent batch. Preserve other regions, slide width, text, assets and styling unless the request changes them. Adding examples means retaining existing examples; move rejected candidates off canvas when requested.
4. Check affected bounds, spacing, crops and relationships before export. Inspect a crop of the changed region at readable scale; use a full-slide view for cross-region changes or final composition checks. Reuse an existing preview when it still represents the relevant content. Repeat only after a new change or a detected defect.
5. After PPTX edits, automatically export the corresponding figure PDF using **slide 1 only**, unless the user changes this scope. Retain all other slides in the source PPTX. Verify that the PDF has exactly one page and reflects the saved edit; follow [incremental-editing.md](incremental-editing.md) for export safeguards. This authorization does not permit compiling/rendering TeX, running bibliography builds, starting build watchers, or updating the manuscript/application PDF. Follow the project's artifact locations. Otherwise reuse a single task directory under `.work/figure-editing/` for scripts, backups, assets and previews; keep formal PPTX/PDF deliverables in their established locations.
6. Report the concrete change and deliverable links briefly. Avoid repeating the whole design, tool implementation, or a full checklist on every micro-edit. Give concise progress updates during longer work.

## Layout rules that prevent repeat corrections

- Existing palette, font sizes and geometry are the reference. Match a named label's actual formatting rather than guessing a default size.
- White inner cards have **no dark outline** by default; use white fill and rounded corners against a light-grey parent. Match visible corner radius across differently sized cards, not the same adjustment percentage.
- Keep a margin between a task-number chip and its card edge, between the chip and body text, and at both ends of every arrow. Measure rendered text before shrinking it.
- For equal-height images, preserve aspect ratio. For equal-aspect viewports, crop or pad; never stretch. Align visible subjects, not just image boxes containing different amounts of whitespace.
- Allocate column widths by occupied content, example count and label length, not equal fractions. Before shrinking pictures or fonts to add examples, reclaim excess whitespace from neighboring groups. Compare adjacent groups at a common visible subject height.
- Use one line style for category boundaries and a lighter, different style for internal subgroups. Match the user-approved exemplar; do not add separators between every example.
- Parallel panels should use a consistent reading order, arrow direction, row baselines and caption placement. Keep headings at the same semantic level typographically consistent; try spacing or line breaks before independent font shrinking.
- Space object silhouettes evenly after trimming irrelevant margins. Align category labels below their own image group and reserve clearance from the next row.
- Keep crops inside source-panel boundaries: no unrelated neighboring objects, leftover arrows, borders or grey fragments. Maintain gutters between image groups, arrows and dashed dividers.
- Do not silently substitute or discard source material. Inspect candidates in small labelled thumbnails; open only selected or ambiguous originals at full resolution.
- Preserve existing arrow semantics. Prefer editable block arrows when bold arrows are requested; match equivalent arrows' lengths and thicknesses. Do not add arrow text by default.

## New figures and larger changes

Use [reference-style.md](reference-style.md) for optional layout patterns. Start from the supplied visual reference; do not force equal-sized grids or placeholder-first construction when real assets exist.

Default to muted pastels, light grey, white, Times New Roman, short labels and minimal decoration. Choose sizes for the final paper scale rather than fixed slide-point values. Add a slide title or explanatory subtitle only when useful or requested.

For Chinese labels, use an available CJK font consistent with the source and inspect actual rendering rather than assuming Times New Roman covers the characters. For preliminary-work figures, preserve the evidence strength: examples or early observations must not become claims that the proposed system is complete or validated. Follow project-specific exclusions for names, references and publication status.

Use the available reliable editing method: existing helper, PowerPoint automation, targeted OOXML, or [PPTX figure kit](../../scripts/pptx_figure_kit.py) for new figures with a working python-pptx environment. Do not switch tools or install dependencies solely to enforce a preferred implementation.

## 基金图的论证职责

新建或大改时，先从当前正文提取该图需要解释的科学问题、关键机制、任务依赖与验证依据；给每个面板一个清楚的论证职责。缺失机制或结果不能靠图形补造，必要时明确标为拟议方案或待补信息。

- 总览图：让读者辨认问题、研究目标、任务关系和预期知识贡献；避免把全部方法细节堆进总览。
- 技术路线图：展示关键输入、核心机制、输出以及与研究问题对应的验证；区分数据流、依赖和反馈，未在正文成立的闭环不要添加。
- 前期基础图：把已有观察与拟研究部分分开，示例图不能冒充实验结果，生成的概念素材不能作为实证证据。
- 验证图或时间表：对应实际指标、比较条件、里程碑与任务依赖；不得虚构数值或完成状态。

这些是按需采用的图类型，不要求每份申请书具备固定图数或固定布局。局部样式修改不扩展为整套研究方案重写。改图后核对正文、图注、任务编号和图中术语；未授权修改正文时，只报告需要同步的位置。只评图时使用 GRF 主技能的评审框架，不自动编辑 PPTX。

## 图 PDF 的画质与交付

修改后按本模块流程同步导出第一页图 PDF；这不授权编译整份申请书。PPTX 和原始素材保留原始质量。默认参数见 [绘图默认值](defaults.md)；按 PDF 页面自身实际放置尺寸计算，超过目标才降采样，低于目标不放大。线稿、Logo、细密图案保留无损和原像素，透明素材优先保留透明度；文字、线条与图形保持矢量。用户后续指示与项目 AGENTS.md 优先。

PowerPoint 常规 PDF 导出可能把图片降至约 200 ppi，不能仅凭成功导出认定合格。检查当前 PDF 内每张相关图片的像素、编码和放置尺寸；必要时从 PPTX 原素材恢复，再对原像素中间 PDF 做一次分类优化。不要对已压缩交付 PDF 反复编码，也不要复用另一张图的图片 xref 编号或认为某个修复脚本覆盖了全部图片。

优先复用项目中已验证的恢复和优化脚本，但先核对其素材范围、匹配规则和当前图片分类；项目脚本不属于本技能的必需依赖。交付前确认单页、编辑内容一致、图片实际参数、矢量保留情况，并检查照片及线稿局部预览。仅有 PNG 预览、未经检查的低清 PDF 或旧 PDF 时，如实说明未完成的导出验证。