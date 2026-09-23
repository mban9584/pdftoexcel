# 代码细节与踩坑清单

配套 SKILL.md；每条都是 2026-09-23 那份 5 页中文价格表上实测出来的，不是推测。

## 1. 表格几何：为什么必须洪水填充

```python
# 表格线有三种画法，只扫一种必翻车：
#   1) 逐格细实心矩形（CorelDRAW 常见，w>5pt、h≈0.25pt）
#   2) 真正的 lines 描边          3) 一条 zigzag path 描完整张表
# page.edges 是这三者的超集（每条线 + 每个 rect/curve 的四条边）
for e in page.edges:
    if e.get('object_type') == 'filter':
        continue
    w, h = e['x1'] - e['x0'], e['bottom'] - e['top']
    if e['orientation'] == 'h' and w > 5 and h <= 3.0:
        y = (e['top'] + e['bottom']) / 2
        if 1.5 < y < page.height - 1.5:               # 排除贴页面四边的框线
            h_segs.append((y, e['x0'], e['x1']))
    elif e['orientation'] == 'v' and h > 5 and w <= 3.0:
        x = (e['x0'] + e['x1']) / 2
        if 1.5 < x < page.width - 1.5:
            v_segs.append((x, e['top'], e['bottom']))
```
把线段画到 4× 分辨率位图 → `cv2.connectedComponentsWithStats(bitwise_not(img))`
→ 被围住的空白连通域就是单元格 → 边界按 2pt 聚类 + 4pt 去重成 `xs/ys`。

- **排除贴页面四边的线**，否则最外层空白变成一个大格子，症状是「格子 N 个、重叠 N−1 个」
  （实测某文件三页各 210/148/465 格、重叠恰好各差 1；加了这个过滤后全部归零）。
- 也要排除贴边的连通域（`x<=1 or y<=1 or x+w>=W-1`）——那是页面外部空白，不是格子。
- 成品必须报 `overlaps == 0`；仍有重叠说明网格聚类阈值需要调。
- 反例记录：只扫 `page.rects` 时，某 25 页文件里有两页分别用 `lines` 和 zigzag path 画线，
  得到 **0 个格子**；`page.lines` 单独扫也会漏掉逐格矩形那种。

## 2. 双路识别的取舍

| | A 整页 det+rec | B 矢量分词 + 纯识别 |
|---|---|---|
| 中文名称标签 | ✔ | ✘ 漏 |
| 窄列 1–2 位数字 | ✘ 漏 7–27 个/页 | ✔ |
| 单个 `/` | 常识成空 | 空（但框还在，可用形状判 `/`） |
| 照片里烤进去的字 | ✔ 会读到 | ✘ 读不到 |
| 框的几何精度 | 检测框偏大 | 由字形轮廓来，精确 |

**冲突时以 B 的文本为准**，但保留"A 有而 B 没有覆盖"的框（那是照片文字）。
判"覆盖"要按**并集**判，不能拿单个框比：

```python
def _cover(items, b):            # items 的 x 区间并集，裁到 b 区间内的长度
    segs = sorted((max(i[0], b[0]), min(i[2], b[2])) for i in items)
    total, cur = 0.0, None
    for lo, hi in segs:
        if hi <= lo: continue
        if cur is not None and lo <= cur[1]: cur = (cur[0], max(cur[1], hi))
        else:
            if cur is not None: total += cur[1] - cur[0]
            cur = (lo, hi)
    return total + (cur[1] - cur[0] if cur else 0.0)
```

## 3. 字号标定（最容易被忽略、又最影响"像不像"）

```python
# 用有文字层的页标定：把 ink 高 / 声明字号
# 实测：12pt 汉字 ink 11.00pt (0.917)；12pt 数字 ink 8.17pt (0.681)；18pt 标题 ink 17.50pt
```
- 无文字层页正文一律取标定出的字号（这份文件 p2–p4 与 p1 同为 12pt，行高 15pt）。
- 标题按 `ink/0.917` 反算（p1 的 18pt 标题 ink 17.5pt，公式给 18.4pt，很准）。
- 别用 `框高/0.72`——实测把 12pt 算成 15.1/16.0/15.3pt。

## 4. 出表的三条硬规则（写错就整体偏移）

1. **索引映射**：`XG = [0] + xs + [W]` 之后，原 `i` 号列 = Excel 第 `i + 2` 列；
   单元格 `[i0,j0,i1,j1]` → `merge(start_column=i0+2, end_column=i1+1, ...)`。
2. **像素吸附**（消除累计漂移）：
   ```python
   b = [0] + [round((v if isinstance(v,float) else v) * 4 / 3) for v in list(xs) + [W]]
   b = [max(b[k], b[k-1] + 1) for k in range(1, len(b))]      # 保证单调、每列 ≥1px
   width[k] = (b[k+1] - b[k] - 5) / 7
   ```
   Excel 落地时算 `round(width*7+5)` px，所以必须**按边界**吸附；按带宽吸附仍会累计（实测 1.72pt）。
3. **锚点**：
   ```python
   pic.anchor = OneCellAnchor(
       _from=AnchorMarker(col=a0, colOff=int((x0 - XG[a0]) * 12700),
                          row=b0, rowOff=int((y0 - YG[b0]) * 12700)),
       ext=XDRPositiveSize2D(int((x1-x0)*12700), int((y1-y0)*12700))
   )
```

## 5. 会静默出错的检查点

- `ws.PAPERSIZE_A4` 在 **Worksheet** 上（值 `'9'`），不在 `ws.page_setup` 上。
- 列宽算出 ≤0 会被 Excel 当隐藏列——每列至少留 1px。
- 数字写成数值后必须给 `number_format`，否则 `110.00` 显示成 `110`。
- `openpyxl.load_workbook()` 能读进 ≠ Excel 能打开：合并区重叠是写侧不报、读侧不报、
  只有 Excel 会提示"修复"，所以必须自己两两求交。
- 本机没有 Excel/LibreOffice，`soffice --convert-to pdf` 这条路不存在。
- `$env:TEMP` 指向不可写的 `C:\Windows\TEMP`，脚本临时文件一律放自建目录。
