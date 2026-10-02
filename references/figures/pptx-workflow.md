
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
5. After PPTX edits, automatically export the corresponding figure PDF using **slide 1 only**, unless the user changes this scope. Retain all other slides in the source PPTX. Verify that the PDF has exactly one page and reflects the saved edit; follow [incremental-editing.md](incremental-editing.md) for export safeguards. This authorization does not permit compiling/rendering TeX, running bibliography builds, starting build watchers, or updating the manuscript/application PDF. Follow the project's artifact locations. Otherwise reuse a single task directory under `.work/figure-editing/` for scripts, backups, assets and previews; keep formal PPTX/PDF deliverables in their established locations. For Windows PowerPoint automation, follow [交互会话导出 SOP](pptx-export-sop.zh.md) before declaring export blocked.
6. Report the concrete change and deliverable links briefly. Avoid repeating the whole design, tool implementation, or a full checklist on every micro-edit. Give concise progress updates during longer work.

### 近期图编辑的防错检查

- 参考图驱动的结果面板先建立“行 × 列”清单，再插入图片。清单至少包含 row id、Sketch、model label、原始文件、分数和显示顺序；插入后逐行检查数量，避免少列或错配。
- 图片统一高度时只统一显示高度，使用原始宽高比计算宽度；不要用 `add_picture(..., width=..., height=...)` 同时强制两个尺寸。不要为了整齐给原图包一层白色背景，除非原始参考图确实包含该背景。
- 远程 WIR/评测素材应先复制原始 model 图片和 JSON/JSONL 分数映射到 `.work/figure-editing/`，保留来源路径和下载时间；汇总 PDF/PNG 只用于参考，不作为最终 model 素材。
- 任务标题行的灰色 chip、chip 内文字和右侧标题要作为一个对齐组处理。先按中心点对齐，再检查渲染截图；不要只按左上角坐标移动其中一个对象。对齐后必须重新检查相邻 Task 面板是否被遮挡。
- 修改圆角或画布高度后，必须渲染整页检查：圆角是否仍覆盖内容、Task 3/4 分隔线是否位于实际边界、底部安全边距是否足够。裁剪画布只能去掉空白，不能改变内容的缩放比例。

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
- For raw result grids, preserve the source image itself. Do not add an artificial white canvas behind a transparent or irregularly cropped model output merely to equalize cells; align the visible subject and keep a consistent display height instead.
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

修改后按本模块流程同步导出第一页图 PDF；这不授权编译整份申请书。PPTX 和原始素材保留原始质量。PowerPoint 的 `native.pdf` 只能作为中间文件，必须重新匹配当前 PPTX 的原始素材并执行一次分类优化后，才能将 `final.pdf` 交付到正式 `figures` 目录。默认参数见 [绘图默认值](defaults.md)；按 PDF 页面自身实际放置尺寸计算，超过目标才降采样，低于目标不放大。线稿、Logo、细密图案保留无损和原像素，透明素材优先保留透明度；文字、线条与图形保持矢量。用户后续指示与项目 AGENTS.md 优先。

PowerPoint 常规 PDF 导出可能把图片降至约 200 ppi，不能仅凭成功导出认定合格。检查当前 PDF 内每张相关图片的像素、编码和放置尺寸；必要时从 PPTX 原素材恢复，再对原像素中间 PDF 做一次分类优化。不要对已压缩交付 PDF 反复编码，也不要复用另一张图的图片 xref 编号或认为某个修复脚本覆盖了全部图片。

优先复用项目中已验证的恢复和优化脚本，但先核对其素材范围、匹配规则和当前图片分类；项目脚本不属于本技能的必需依赖。交付前确认单页、编辑内容一致、图片实际参数、矢量保留情况，并检查照片及线稿局部预览。仅有 PNG 预览、未经检查的低清 PDF 或旧 PDF 时，如实说明未完成的导出验证。
### Task-block ground truth from user-edited figures

For multi-task overview figures, treat each Task as a complete vertical block with an explicit top and bottom boundary. Keep `(a) Task 1` through `(d) Task 4` in their existing narrow label boxes and center the paragraph inside each box. Do not widen these boxes across the header band and do not change their x-position or width to simulate centering; that makes the Task label overlap the descriptive heading on the left.

Place one horizontal dashed separator at every boundary between adjacent Task blocks, including the boundary between Task 1 and Task 2. Derive each separator y-coordinate from the actual block boundary after the block contents are positioned. Do not reuse two fixed page-level line coordinates when the number or heights of Task blocks have changed.

When tightening the gap between a Task header and its first subsection, move the header, subsection heading, and the block content as a coordinated group. Preserve the left descriptive heading and the narrow Task label as separate objects. Verify the full-slide composition after vertical reflow; do not solve spacing by moving only a label or only a separator.


### 原图恢复时的透明度与去背景防错

- 原始图片文件不一定包含最终显示的透明效果。恢复画质前，同时检查素材 alpha、图片对象透明色（如 OOXML `a:blip/a:clrChange`）、PowerPoint 去背景效果，以及干净导出 PDF 的 `/SMask`、`/Mask` 和相关透明状态；不能只读取 `ppt/media` 的 RGB 像素替换 PDF 图片。
- 按图片对象实例记录素材、裁剪、透明色和去背景效果。同一素材可被多个对象以不同效果使用，不能仅以素材路径为键合并处理；缩略图相似匹配也必须核对实例和透明状态。
- 对明确的颜色键透明，按 PPTX 声明的颜色和 alpha 重建透明度，与素材原有 alpha 合成，并保持裁剪与蒙版对齐。不要对所有图片统一删除白色或擅自增加近白阈值，以免误删白色服装、Logo 或高光。复杂去背景效果应保留或准确重建其蒙版；不能假定颜色键能覆盖所有去背景方式。
- 替换 PDF 图片时不得因清空图片字典而丢弃 `/SMask` 或 `/Mask`。恢复后的 RGB 与透明蒙版须在像素尺寸、方向、裁剪和位置上对应；透明素材按项目规则保留透明度及无损编码。不能为缩小体积直接转成无 alpha 的 JPEG 或铺白底。
- 交付前逐项比对干净导出与最终 PDF 的透明图片清单，并检查蒙版实际内容；仅确认存在 `/SMask` 不足以证明透明效果正确。在原有浅灰或彩色面板背景上放大检查服饰轮廓、白色物体、细边、重叠区及周围文字，确认无白色矩形、白边、误抠、遮字或缺失。画质恢复导致文件增大时如实报告，不以牺牲透明效果换取固定压缩率。
- 已知故障：恢复原像素时忽略 PPTX 的白色透明设置，使原已去底的服饰重新出现白色矩形并遮挡相邻元素。必须通过透明来源映射、PDF 蒙版核对和局部预览共同验收，不能仅凭图片分辨率与编码检查认定导出合格。


### 上传体积受限时的分类优化（2026-09-27 验证）

- 先从当前 PPTX 干净导出第一页，再由原始素材恢复所需像素并执行一次优化。PPTX 与原始素材不降质，不反复压缩既有交付 PDF。
- 区分内容类型与透明属性：不透明照片约 450 ppi、JPEG75；透明照片也可按实际放置尺寸降至约 450 ppi，但 RGB 与 alpha 仍用无损编码。无损编码不要求保留照片的冗余像素。线稿、Logo、细密图案无论是否透明，均保留原像素及无损编码。低于目标的照片不放大。
- 透明照片降采样时保持 RGB 与蒙版同尺寸、同裁剪、同方向，保留 PPTX 的颜色键透明与去背景效果；检查原面板背景上的边缘，不能铺白底、丢 alpha 或产生白边。文字、公式、箭头、几何图形保留矢量。
- 按用户平台限制检查实际字节数并留余量。对 50 MB 上限，可按小于 50,000,000 字节保守验收；多文件同时上传时报告合计，并在画质允许时将合计也控制在限制内，但不要把保守目标说成已确认的平台合计限制或保证上传成功。
- 新导出可能改变资源编号。重新核对素材、裁剪、透明效果、图片实例与 PDF 资源，不能直接沿用旧 xref 或保护名单；仅尺寸一致不能证明匹配正确。清理未用及重复资源，检查单页、实际嵌入像素/编码/ppi、文本与矢量、透明蒙版和整页及局部预览。
- 已验证项目案例：0 图由约 75.6 MB 降至 14.36 MB，1 图由约 25.2 MB 降至 3.60 MB，合计 17.96 MB。主要收益来自透明照片的适度降采样；这些数值不是通用压缩率或固定目标。项目下 `.work/figure-editing/input-output-export/` 与 `overview-flow/` 的 `optimize-upload.py`、`quality-audit.json` 可作为实现参考，不应盲用样本编号。
