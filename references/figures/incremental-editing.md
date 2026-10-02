# Efficient incremental PPTX editing

Read when establishing or repairing the editing workflow; reuse its state on later turns.

## Keep compact project state

For sustained editing, maintain one compact state file in the existing task directory. Keep only information needed for the next edit:

- Current PPTX/PDF, source fingerprint, slide dimensions and latest verified preview.
- Confirmed decisions, unresolved questions and the next requested delta; replace superseded decisions instead of appending conversation history.
- Affected semantic shape mappings, reusable asset/crop references and relevant style values.
- Reusable helper/config paths and any active generation/export job identifiers.

Update changed entries only. Resume from this state and the user's latest request, then query the affected region. Verify the source fingerprint and cached shape IDs; refresh stale mappings locally rather than trusting them or rescanning the entire deck. Keep project-specific state out of the reusable skill. For a one-off edit, do not build a manifest or framework merely to satisfy this pattern.

Reuse one project editor and exporter when practical. Separate stable operations (text, crop, align, distribute, style, export) from a small per-edit configuration containing target groups, dimensions, spacing and replacements. Extend the helper only for operations actually needed; avoid copying a complete editor each turn or constructing a general engine for a minor change. Patch the latest artifact and retain a recoverable snapshot before each batch.

## Keep context and output small

- Read each applicable skill/reference once while its contents remain available. Load conditional guidance only when needed; keep each rule in one authoritative location and link to it elsewhere.
- Request concise tool results: affected objects, changed fields, validation failures and output paths. Avoid dumping image bytes, whole XML or repeated logs. Batch independent reads and choose a relevant output limit.
- Review one targeted preview after a coherent edit batch. Broaden to a full view only for cross-region changes or an unresolved composition concern; do not reduce necessary visual checks to save tokens.
- For a small edit, report the concrete change and deliverable links. Mention limitations or unresolved decisions when present; omit a repeated walkthrough of unchanged settings and routine checks.

## Inspect and verify economically

- Query only affected shapes: ID, name, text, bounds, style or media reference. Avoid full XML, entire scripts, recursive asset inventories and full manuscript dumps. Search the relevant section first.
- Batch independent reads. Use a labelled contact sheet or thumbnails for asset selection rather than sending many full-resolution images into context. Preserve originals; preview generation does not replace editable assets.
- Compute group layout once: common visible subject height, aspect-preserving widths, occupied span and equal gaps. Use the widest resulting silhouette and required label widths to calculate a feasible group width. Normalize content bounds where whitespace causes apparent size differences. When columns compete for space, redistribute available width across the affected row before reducing subject height. Check corresponding visible edges and baselines across neighboring groups, not just within each group.
- Before export, check text/image and image/divider clearance, intended crop ranges, row baselines, equal gaps and containment. These catch the errors that repeatedly require another PowerPoint export.
- Export once after the batch. Poll only a running export. Render just the affected PDF region/page at sufficient resolution; inspect the full composition when geometry across regions changes. Keep visual QA, but avoid duplicate whole-slide previews with no new evidence.

## OOXML and PowerPoint safeguards

Preserve group transforms and schema ordering. Use unique shape IDs. When adding images, update media, slide relationships and content types together; verify every embed resolves before opening PowerPoint. Do not ban existing connectors merely because they are cxnSp elements.

Do not reuse numeric collection indices across edits or PowerPoint saves. Re-identify targets in the current file by shape ID/name plus text, bounds and role; assert the expected match count before changing anything. A title's text and its empty background may be separate objects. Replace the complete obsolete result group, including score labels and backgrounds, rather than leaving duplicate text beneath new images. Compare the source hash before saving to avoid overwriting concurrent edits.

For horizontal centerline alignment, choose a shared vertical center `cy` and set each object's `top = cy - height / 2`. Equal top coordinates do not align objects of unequal height. Check text vertical anchoring and paragraph margins as well as shape bounds, and inspect the exported title-row crop before claiming visual alignment. A successful save or one-page export is not evidence that a visual correction is complete.

Cache extracted images by media identity/hash or by source snapshot plus shape ID; a reused shape ID or filename alone can silently select stale artwork after replacements. Keep content measurements separate from raster edits: inspect/measure bounds, then set native crops where possible.

Native PowerPoint crop coordinates keep originals editable and avoid generating a new raster per crop. Check cropped evidence visually: scaling the bounding rectangle alone cannot correct an inconsistent viewport or source whitespace.

Use PowerPoint COM read-only, windowless open for export when available.
### PowerPoint COM session fallback

On Windows, the preferred PDF exporter is PowerPoint COM with a read-only presentation and an explicit slide-1 print range. A restricted VS Code or plugin host can fail to create the COM desktop session with `0x80070520` (`A specified logon session does not exist`). When this occurs, rerun the same export script in the desktop/user session with the required execution authorization; do not change the PPTX or switch to unrestricted full-deck export. Keep the COM cleanup and shared PowerPoint safeguards below.

Operational troubleshooting details from the GRF workflow:

- If `PowerPoint.Application` exists but `Presentations.Open` returns no usable presentation object, do not report export success. Explicitly open the input PPTX in the desktop user session, then retry COM with the same file. Check `Slides.Count` before exporting.
- PowerPoint COM enumeration values matter: `ExportAsFixedFormat` with an explicit `PrintRange` must use range type `ppPrintSlideRange = 4`; value `2` is not the explicit slide-range mode and can trigger “selected slides do not exist”. Create the range with `Presentation.PrintOptions.Ranges.Add(1,1)` and pass that range object to export.
- Avoid passing integer `0` where an optional COM object is expected. Use the actual range object and suitable missing optional values. Do not set `.Visible` with a PowerShell boolean; MsoTriState uses integer values, and visibility is unnecessary for a windowless read-only open.
- Export to a temporary PDF first, validate one page and the saved edit, then copy to the formal `figures/` path. The user's manual desktop PDF is a valid fallback when explicitly supplied: copy it to the requested `figures/` directory, not `assets/figures/`, and verify page count.

When the destination PDF or preview does not exist yet, pass the output path directly to `ExportAsFixedFormat` or `Slides.Item(1).Export`; do not call `Resolve-Path` on the new output file. Use `Resolve-Path` only for existing input files. After export, verify that the PDF exists, contains exactly one page, and reflects the saved PPTX before replacing the formal figure PDF. Close only the presentation opened by the helper and release references in a finally block. Do not call Application.Quit on a potentially shared PowerPoint instance or kill PowerPoint processes. Respect environment permissions and applicable artifact-tool requirements. An export succeeding does not prove there are no overlaps.

## Export only the first slide

Use an explicit slide-1 range in the exporter; an unrestricted SaveAs(PDF) exports the entire deck. If the available exporter cannot select pages, export to a temporary PDF and extract page 1 with an available PDF tool, then verify the final PDF contains exactly one page. Do not delete slides from the source deck to implement this preference, and do not rasterize the PDF just to select a page.

Wait for the exporter to finish successfully before opening its PDF; do not mistake an in-progress missing file for an export failure. After interruption or a retry request, check pending generation/export jobs and existing valid results before starting duplicate work.

Write export output to a temporary path and replace the figure PDF only after successful page-count verification. Keep an existing PDF intact if export fails, and label it as stale rather than describing it as updated. PNG preview export alone is not PDF synchronization. The bundled Python kit supplies PNG preview only; PDF export needs a separate available exporter.

Before replacing formal files, confirm the source has not changed since the edit snapshot (for example, compare its hash). If it has, reapply the delta to the latest file. Promote the verified PPTX and its corresponding slide-1 PDF together.

