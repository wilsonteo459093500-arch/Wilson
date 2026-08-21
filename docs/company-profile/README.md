# Riccione Company Profile 2026 · 可编辑 PPT 版

`Riccione_Company_Profile_2026_editable.pptx` 由 `Riccione_Company_Profile_2026.pdf`
转换而来，**23 页，版式 / 位置 / 字号 / 字距与原稿一一对应**，所有文字都是真正的
PowerPoint 文本框（可点、可改、可搜索、可复制）。

## 规格

| 项目 | 值 |
|---|---|
| 页数 | 23（与 PDF 一致） |
| 画布 | 13.333 × 7.5 in（16:9 标准宽屏，导出图片即 1920×1080） |
| 底图分辨率 | 每页 3840 × 2160（288 DPI，≈ 原 PDF 内嵌图片的原生分辨率） |
| 文本框 | 193 个 / 349 行，全部保留原基线坐标与对齐方式 |
| 文件大小 | ~27 MB |

## 结构：一层底图 + 一层文字

每页只有两类对象：

1. **`Background pN`** — 整页底图（照片、色块、线条、logo、图标）。
   为避免误拖动，已加 `noMove / noResize` 锁；要换图就在
   PowerPoint 里右键 → 更改图片，或先在「选择窗格」里解锁。
2. **文本框** — 原 PDF 里的每一段文字。按「同字体 + 同字号 + 同颜色 + 等行距 + 同对齐」
   合并成段落，所以一段话是一个框，改起来不碎。

> 为什么文字不做成矢量图形：那样虽然像素级一致，但不可编辑。
> 现在这种做法是「视觉 100% 还原 + 文字 100% 可编辑」的最佳平衡。

## 位置是怎么对齐的

- 坐标按 `13.333in / 1920pt` 等比换算，横向起点取 PDF 的**排版基点**（不是字形边框），
  纵向由**基线**反推文本框顶端（固定行距 = PDF 实测行距）。
- 每个 span 的**字距（letter-spacing）**都是从 PDF 字符步进里反解出来后写回 PPT 的，
  例如英文大标题 Sabon 的 +0.02em 字距被完整保留。
- 原稿里居中 / 右对齐的段落，PPT 里也是居中 / 右对齐 —— 改了字数依然自动居中。
- 实测（LibreOffice 渲染回 PDF 后与原稿逐行比对，页面按 1920×1080 计）：
  横向偏差中位数 **0.03 px**、纵向 **1.1 px**（≈ 0.1%），肉眼不可见。

## ⚠️ 字体：请先装字体再编辑

PPT 里写的是原稿的真实字体名。**装了字体 = 与 PDF 完全一致；没装 = PowerPoint 会替换成别的字体，
版面就会走样**（位置仍然对，但字形和宽度不一样）。

| PPT 中的字体名 | 用在哪 | 说明 |
|---|---|---|
| `Source Han Sans CN Normal` | 中英文正文（7621 字，17 页） | 思源黑体 CN Normal，Adobe / Google 免费开源（= Noto Sans CJK SC DemiLight） |
| `Source Han Sans CN Light` | 页眉小标题 | 思源黑体 CN Light |
| `Source Han Serif CN` | p3 单个字符 | 思源宋体 CN |
| `youyou-yisong` | 中文标题 / 说明（806 字，21 页） | PDF 内嵌名即为 `youyou-yisong`；若你机器上装的是中文名（如「悠悠宋」），见下方"改字体名" |
| `Sabon LT Std`（粗体） | 英文大标题（Our Story / Thank You 等） | Adobe 商业字体 |
| `PingFang SC Ultralight` | p2 页码 | macOS 系统自带 |

**改字体名**：如果某个字体在你电脑里叫别的名字，不必手动逐个改 ——
用 PowerPoint 的「开始 → 替换 → 替换字体」一次换掉；
或者改 `tools/pdf2pptx/fonts.example.json` 后重新生成（见下）。

## 重新生成

```bash
pip install -r ../../tools/pdf2pptx/requirements.txt
python3 ../../tools/pdf2pptx/pdf_to_editable_pptx.py \
    Riccione_Company_Profile_2026.pdf \
    Riccione_Company_Profile_2026_editable.pptx \
    --bg-width 3840 --jpeg-quality 90 \
    --font-map ../../tools/pdf2pptx/fonts.example.json
```

`conversion-report.json` 记录了每页的文本框数量、行数与底图体积。
