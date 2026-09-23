# pdftoexcel — 把价格表 / 画册类 PDF 原样转成 Excel 的 Qwen Code 技能

表格线、合并单元格、行高列距、产品照片、红蓝字体颜色都按原件复刻，每页正好打印成一张 A4。
给 [Qwen Code](https://github.com/QwenLM/qwen-code) 用；也可以完全手动跑脚本，不依赖模型。

- 仓库：<https://github.com/mban9584/pdftoexcel>
- 给模型读的流程文件：[`SKILL.md`](SKILL.md)；代码细节与踩坑：[`reference.md`](reference.md)

---

# 一、安装

## 1.1 前置条件

| 需要 | 怎么确认 | 缺了怎么办 |
|---|---|---|
| Python 3.10+（含 pip） | `python --version` | Windows 去 <https://www.python.org/downloads/> 装，**勾选 "Add python.exe to PATH"** |
| Qwen Code | `qwen --version` | `npm install -g @qwen-code/qwen-code`（需要 Node.js 18+） |
| 网络能到 github.com 与 pypi.org | 见 1.4 的代理说明 | 公司内网/代理环境要额外配 `--proxy` |

Python 库由安装脚本自动装，不用手动 pip。默认只装文字层/版式复刻所需的核心依赖；只有遇到无文字层 PDF 时，才另外安装 OCR 依赖。

## 1.2 选一个 skills 目录

技能发现路径只有两个，**目录名就是技能名**，所以必须叫 `pdftoexcel`：

| 类型 | 路径 | 作用范围 |
|---|---|---|
| 个人技能（推荐） | Windows `C:\Users\<你>\.qwen\skills\pdftoexcel`<br>macOS/Linux `~/.qwen/skills/pdftoexcel` | 你所有项目都能用 |
| 项目技能 | `<仓库>\.qwen\skills\pdftoexcel` | 只在该项目里，可随 git 分发给同事 |

## 1.3 安装（三选一）

**A. 有 git（最好，以后一条命令更新）**

```powershell
# Windows PowerShell
git clone https://github.com/mban9584/pdftoexcel.git "$env:USERPROFILE\.qwen\skills\pdftoexcel"
```
```bash
# macOS / Linux
git clone https://github.com/mban9584/pdftoexcel.git ~/.qwen/skills/pdftoexcel
```

**B. 没有 git：下 zip**

```powershell
# Windows PowerShell
# 全程只在 $d 里操作，不要用 $env:TEMP——有的机器它指向不可写的 C:\Windows\TEMP
$d = "$env:USERPROFILE\.qwen\skills"
New-Item -ItemType Directory -Force $d | Out-Null
Invoke-WebRequest https://codeload.github.com/mban9584/pdftoexcel/zip/refs/heads/main -OutFile "$d\pdftoexcel.zip"
Expand-Archive "$d\pdftoexcel.zip" -DestinationPath "$d\_x" -Force
Move-Item "$d\_x\pdftoexcel-main" "$d\pdftoexcel" -Force
Remove-Item "$d\_x","$d\pdftoexcel.zip" -Recurse -Force
```
```bash
# macOS / Linux
cd /tmp && curl -LO https://codeload.github.com/mban9584/pdftoexcel/tar.gz/refs/heads/main
mkdir -p ~/.qwen/skills && tar xzf main
mv pdftoexcel-main ~/.qwen/skills/pdftoexcel    # 注意改名
```

**C. 项目里给全组用**：在你的项目目录下执行
`git clone https://github.com/mban9584/pdftoexcel.git .qwen/skills/pdftoexcel`
然后提交 `.qwen/skills/pdftoexcel` 里的文件（或把它加成 submodule），同事 `git pull` 即得。

## 1.4 装 Python 依赖（每台机器一次）

```
cd 到技能目录
Windows:   install.bat
macOS/Linux:  sh install.sh
```

脚本结尾会打印每个模块是否可用，看到这样才算装好：

```
pdfplumber   ok
pypdfium2    ok
numpy        ok
cv2          ok
openpyxl     ok
PIL          ok
rapidocr     ok (outline-ocr route available)
```

`rapidocr` 显示 `absent` 不影响主流程——只有遇到**没有文字层**的 PDF（CorelDRAW / 某些导出
软件把文字画成矢量描边）才需要它。要单独补装：

```
python -m pip install --user -r requirements-ocr.txt
```

**代理环境**（pip 走不动时）：

```
python -m pip install --user -r requirements.txt --proxy http://127.0.0.1:10808
```

## 1.5 启用并验证

1. **重启 Qwen Code**（`/exit` 再启动）。技能目录有文件监听，但重启最可靠。
2. 输入 `/skills` —— 应看到 `pdftoexcel [Personal]`（或 `[Project]`）。
3. 或直接输入 `/pdftoexcel` 强制调用它。
4. 冒烟测试（不依赖模型也能验）：

```
cd 技能目录
python scripts\pdftoexcel.py probe --pdf "任意表格类PDF.pdf"
```
能打印出每页的 `cells=... overlap=0` 和一行 `route -> ...`，就是装成功了。

---

# 二、使用

## 2.1 让模型自动做（正常用法）

直接在 Qwen Code 里说人话，技能会按阶段引导它：

| 你说 | 它会做 |
|---|---|
| "把 `xx价格表.pdf` 转成 Excel，放桌面，要能 A4 打印" | 一页一个 sheet，版式照原件 |
| "这份价格表统一打 9 折" | 只有单价列乘 0.9，保留 2 位小数，不保留原价 |
| "把单价清空做成填空模板" | 原位保留照片与表格线，单价列连同占位横杠一起清空 |
| "这个画册转 Excel" | 表格页出表、封面/照片墙页嵌图与文字 |

它会先问你三件事（处理方式 / 输出文件名 / 照片 dpi），然后自己跑 probe → model → xlsx → check，
最后交给你一份 xlsx 并说明**哪些是反推校验、哪些还需要你眼睛看**。

## 2.2 手动跑（或排查问题时）

```
python scripts/pdftoexcel.py probe   --pdf "输入.pdf" [--pages 0,1,2] [--work 目录]
python scripts/pdftoexcel.py ocr     --pdf "输入.pdf" --pages 2,3     [--dpi 300]
python scripts/pdftoexcel.py model   --pdf "输入.pdf" [--pages ...]
python scripts/pdftoexcel.py xlsx    --pdf "输入.pdf" [--out 文件] [--dpi 300] [--font 宋体]
python scripts/pdftoexcel.py check   --pdf "输入.pdf" [--xlsx 文件]
python scripts/pdftoexcel.py fontsize --pdf "输入.pdf" --page 1 --rect x0,y0,x1,y1
```

| 参数 | 默认 | 说明 |
|---|---|---|
| `--pdf` | 必填 | 源 PDF（页码从 **0** 开始计数） |
| `--work` | `~/.qwen/tmp/<PDF名>/` | 中间产物目录；也可用环境变量 `PDFTOEXCEL_WORK` |
| `--out` | `~/Desktop/<PDF名>.xlsx` | 仅 `xlsx` 用；给目录则用同名文件 |
| `--pages` | 全部 | 逗号分隔，如 `0,2,3` |
| `--dpi` | 300 | 照片裁剪清晰度。页数多又只要打印清楚 → 200 可显著减小体积 |
| `--font` | `宋体` | 目标机器没这字体就换 `黑体`/`等线`/`Microsoft YaHei` |
| `--auto-slash` / `--slash` / `--no-slash` | 自动 | 自动、强制开启或关闭"无货 `/`"的墨迹推断（有文字层时通常不要强制开） |

`check` 的三行输出怎么读：

- `overlap=0` —— 必须为 0，否则 Excel 打开会提示"修复"
- `col edges off PDF by max ≤0.4pt` —— 列边界与原件对齐；若出现 `drifts` 说明列宽没吸附整像素
- `content ok (N cells)` / `MISMATCH` —— 逐格与源比对（按**格式化后的显示值**比，
  所以 `72`+`0.00` 会被正确认成 `72.00`）

## 2.3 产物与中间文件

```
~/.qwen/tmp/<PDF名>/
├── cells/  g*.json 几何     m*.json 内容模型     v*.png 网格叠加图（要人工看一眼）
├── ocr/    p*.json 整页识别  w*.json 矢量词框      （只有无文字层文件才有）
└── imgs/   裁出来的产品照片
```
这些 JSON 是**断点续做的关键**：下次只改出表逻辑就直接跑 `xlsx`，不用重跑几何和识别。

---

# 三、常见问题

**Q：`probe` 报 `cells=0`，但我明明看到表格线？**
线被画成 `lines` 描边或一条 zigzag 路径。本技能已改扫 `page.edges`（三者超集）覆盖这种情况；
若仍为 0，看 `--pages` 选错了吗，或该页确实是图片表格（那要 OCR 路线）。

**Q：`overlap` 数字 = `cells-1`？**
页面外框线被当成表格线了（最外层空白被识别成一个大格子）。确认你用的是最新脚本；
自己改代码时记得过滤贴边的线。

**Q：Excel 打开提示"发现部分内容有问题，是否修复"？**
就是有合并区域互相重叠，`check` 里 `overlap` 不为 0。修网格，别硬推给 Excel。

**Q：数字变成 `110` 而不是 `110.00`？**
说明该格没拿到 `number_format`。跑 `check`；`xlsx` 阶段是按原文小数位设格式的。

**Q：照片很糊 / 文件太大？**
`xlsx --dpi 200`（体积约降到一半）或 `--dpi 400` 更清晰。

**Q：`pip` 装不上 / 卡住？**
先试 `python -m pip install --user -r requirements.txt --proxy http://127.0.0.1:<端口>`；
国内网络 pypi 可换 `--index-url https://pypi.tuna.tsinghua.edu.cn/simple`。

**Q：中文文件名/路径报错？**
本技能的路径全部走 `expanduser` 与 `os.path.join`，不做字符串拼接；
若在 PowerShell 里传参，路径**加英文双引号**。

**Q：能转成 PDF 自己比对吗？**
只有装了 Excel 或 LibreOffice 才能（`soffice --convert-to pdf 输出.xlsx`）。
这台机器上有就**必须**做这一步——那是唯一能发现位置错位的真视觉验证。

**Q：想只改代码不出表？**
`scripts/stages.py` 是两个阶段的实现（`build_model` / `write_xlsx`），
`scripts/_paths.py` 集中管路径，`scripts/pipeline/` 是历史存档（只读，见其 README）。

---

# 四、更新与卸载

```
更新：  cd 技能目录 && git pull
卸载：  删除技能目录即可（~/.qwen/skills/pdftoexcel），不留任何配置或注册项
```

# 五、来源与许可

做法来自 2026-09-22/23 两份真实文件的转换（一份 5 页价格表、一份 25 页产品画册），
过程中的判读规则与坑全部写进了 `SKILL.md` 和 `reference.md`。仓库里**只有代码和方法，没有任何业务数据**。

许可证：MIT，见 [`LICENSE`](LICENSE)。


