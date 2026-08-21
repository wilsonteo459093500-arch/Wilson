# pdf2pptx — 设计稿 PDF → 可编辑 PPTX

把 InDesign / Figma 导出的 PDF 转成版式完全一致、但文字可编辑的 PPTX。

```bash
pip install -r requirements.txt
python3 pdf_to_editable_pptx.py input.pdf output.pptx \
    [--bg-width 3840] [--jpeg-quality 90] \
    [--font-map fonts.json] [--report report.json]
```

## 原理

1. **底图**：复制一份 PDF，用 redaction **只删文字**
   （`images=NONE, graphics=NONE` → 图片与矢量图形原样保留），
   再按 `--bg-width` 高倍渲染整页，作为锁定的背景图片。
2. **文字**：从原 PDF 读每个 span 的基线坐标、字号、颜色、字距，
   在 PPT 上重建为真正的文本框。
   - 行 → 段落合并条件：同字体 / 同字号 / 同颜色 / 等行距 / 同对齐。
   - 固定行距用 PDF 实测行距；文本框顶端 = `基线 − 行距 + 下伸部×字号`
     （PowerPoint 固定行距下首行基线的位置）。
   - 字距由「字符实际步进 − 字体自然步进」反解，逐 span 写回 `a:rPr/@spc`。
   - 关闭自动换行（`wrap=none`），因此增删文字不会引起整页重排。
3. **坐标**：`slide_width / page_width` 等比缩放，版式关系与原稿一致。

## 已处理的坑

- **旧式数字**：设计工具把数字映射到私用区 `U+F6B1–F6BA`，还原成 `0–9`，否则在 PPT 里是豆腐块。
- **连字**：`ﬁ ﬂ ﬀ ﬃ ﬄ` 拆回普通字母，方便编辑与搜索。
- **竖排文字**：保留在底图上，不做成文本框（横排判断 `line["dir"] == (1,0)`）。
- **对齐**：多行段落按左 / 中 / 右实测判定；单行若相对整页居中则保持居中。
- **`endParaRPr`**：在段尾继续输入时字体样式不跳变。

## 校验

生成后建议做几何回归：LibreOffice 把 PPTX 转回 PDF，逐行比对文字的
起点 x 与基线 y。本仓库的企业手册实测：横向中位数偏差 0.03 px、
纵向 1.1 px（按 1920×1080 计）。
