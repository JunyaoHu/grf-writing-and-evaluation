# PPTX 第一页 PDF 导出 SOP（Windows）

## 这次故障的原因

本项目使用 PowerPoint COM 导出，而不是 TeX 编译。此前失败不是 PPTX 路径或文件内容错误，而是 Codex 命令运行在后台会话，PowerPoint 运行在用户桌面会话。两者不在同一交互会话时，`New-Object -ComObject PowerPoint.Application` 会返回 `0x80070520`（指定的登录会话不存在）。即使能看到 `POWERPNT.EXE`，也可能只有后台或隐藏进程，不能据此判断 COM 可用。

用户将 Codex 的电脑访问权限切换为“完全访问权限”后，当前会话获得了访问桌面应用的能力；同时 PowerPoint 保持在可访问的交互窗口中，COM 创建成功，导出才恢复正常。另一个容易混淆的限制是：在已有可见 PowerPoint 实例中设置 `$pp.Visible = $false` 可能返回“隐藏应用窗口不被允许”，因此不要强制隐藏已有实例。

## 执行前检查

1. 确认项目根目录和四个输入文件：`2027-draft-git/assets/ppt/*.pptx`。输出放在 `2027-draft-git/figures/`，文件名与 PPTX 同名。
2. 在 Codex 的电脑访问设置中启用“完全访问权限”。权限切换后如 COM 仍失败，重启 Codex，再保持 PowerPoint 窗口打开。
3. 先做最小 COM 探针，不要直接批量覆盖 PDF：

   ```powershell
   try { $pp = New-Object -ComObject PowerPoint.Application; 'COM_OK' }
   catch { "COM_FAIL: $($_.Exception.Message)" }
   ```

   若出现 `0x80070520`，说明仍是会话隔离；此时不要反复改文件路径或重试导出。
4. 若探针成功，检查 PowerPoint 进程有可见主窗口。已有实例通常应保持可见；不要在脚本中强制 `$pp.Visible = $false`。

## 导出步骤

1. 通过 PowerPoint COM 打开每个 PPTX。
2. 使用 `PrintOptions.Ranges.Add(1, 1)` 创建只包含第 1 页的打印范围。
3. 调用 `ExportAsFixedFormat`，设置 `RangeType=3`（指定幻灯片范围），先输出到 `.work/figure-editing/<figure>/native.pdf`，仅作为中间文件。
4. 程序自动从当前 PPTX 匹配原始素材，恢复透明度和像素，并按照片约 450 ppi、JPEG quality 75，线稿/Logo/透明素材无损的规则优化。
5. 每个文件完成后关闭演示文稿；全部完成后退出脚本创建的 COM 应用对象。每个文件完成后关闭演示文稿；全部完成后退出脚本创建的 COM 应用对象。不要修改源 PPTX 的其他页。
6. 用 `pdfinfo` 或等效工具检查每个输出 PDF 的 `Pages: 1`，再用 `pdfimages -list` 检查实际嵌入图片的像素尺寸、编码和 ppi。若主要位图仍约 95–200 ppi，说明只完成了 PowerPoint 原生导出，流程没有完成。

## 故障分流

| 现象 | 判断 | 处理 |
|---|---|---|
| `0x80070520` | Codex 与 PowerPoint 不在同一登录会话 | 开启完全访问权限并重启 Codex；保持 PowerPoint 在桌面会话中打开 |
| `Application.Visible: Invalid request` | 脚本试图隐藏已有 PowerPoint 实例 | 删除 `$pp.Visible = $false`，保留实例可见 |
| “找不到指定的路径” | 脚本根目录计算错误或输入文件不存在 | 以项目根目录解析 `assets/ppt`，逐个 `Test-Path` |
| PDF 页数大于 1 | 未传入打印范围或 `RangeType` 错误 | 使用 `Ranges.Add(1,1)` 和 `RangeType=3`，再重新导出 |
| PDF 生成但画质异常 | 检查是否使用打印质量和 600 DPI 设置 |

## 最终验收

- 四个 PDF 均存在，名称与四个 PPTX 一一对应。
- 每个 PDF 恰好 1 页，内容对应 PPTX 第 1 页。
- PPTX 其余页面仍保留，未被裁剪或覆盖。
- 交付前确认 PDF 为单页，并放大抽查图片局部清晰度。
- 只更新图 PDF，不因图导出而编译整份申请书或更新申请书 PDF。

## 项目固定导出程序

本项目提供固定路径脚本 `.work/figure-editing/export-grf-figure.ps1`。用户只需指定文件名：

```powershell
& ".work\figure-editing\export-grf-figure.ps1" "outfit-overview-new"
& ".work\figure-editing\export-grf-figure.ps1" "preliminary-results-new"
& ".work\figure-editing\export-grf-figure.ps1" "sota-failures-new-1"
& ".work\figure-editing\export-grf-figure.ps1" "tasks-overview-new"
```

脚本固定读取 `2027-draft-git/assets/ppt/`，固定写入 `2027-draft-git/figures/`，只导出第一页，并在结束时检查 PDF 是否为单页。脚本使用 PowerPoint 打印质量和 600 DPI 图片设置；导出后只需检查单页和局部清晰度。



