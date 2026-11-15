# WordToPdfTool

Word → PDF 保真转换工具。通过 WPS / Microsoft Word 的 COM 自动化把 `.docx` 导出为 PDF，
再用 `.docx` 包里嵌入的**原始图片**回填 PDF 图像流，去掉 Word 导出时的二次有损压缩，
在不损失画质的前提下显著缩小文件体积。支持单文件、文件夹批量转换与多 PDF 顺序合并。

## 它解决什么问题

Word/WPS 导出 PDF 时会把文档内的图片重新编码成 JPEG。以一份 2,145 页检测报告为例：
Word 原文件 55.15 MB，导出 PDF 涨到 **465.89 MB**，且图像与 Word 原图平均存在 16.14%
的像素差异。

本工具用 Word 包内的原始图片字节回填 PDF 图像流，使交付 PDF 的图像与原始素材一致，
同时把体积恢复到与 Word 文档相当的量级。在 25 份真实投标材料（1.5 万页）上的实测：

```
Word 平铺导出    1259.72 MB
本工具交付        404.30 MB   (-67.9%)
```

同时，以 `.docx` 内嵌原图为基准，交付 PDF 的图像偏差为 **0.0000**，而纯 Word 导出的 PDF
是 **0.7579–1.7293** —— 交付版比 Word 自己导出的更接近原始素材。

## 目录结构

```
src/
  pdf_tool.py              主程序：GUI（Win32/ctypes）、三种处理模式、调度、合并
  wps_to_pdf.py            导出层：WPS/Word 文档 COM、WPS 表格 COM（KET）、WPS CLI
  optimize_pdf_images.py   优化层：原图回填、无损重编码、完整性校验、逐页渲染审计
  link_fix.py              链接层：跳转全量校验与补链（只动 /Link 注解）
  page_fill.py             页码层：表格空缺页码求解与回填（只改目标单元格）
  deliver_pipeline.py      编排层：两轮导出 + 收敛检测 + 图像 + 跳转 + 合并
  run_deliver.py           编排层 CLI 入口（含 --plan-only 只出计划）
docs/
  使用说明.md               完整使用文档（含原理、限制、实测数据、版本变更）
  代码评审报告.md           对 V1.2 的代码评审（问题清单与改进依据）
  改进建议_V1.3.md          改进项的技术论证与验证过程
tools/
  build_exe.ps1             PyInstaller 打包脚本
  selftest.ps1             回归自测（PowerShell 版）
  selftest.py              回归自测（跨平台版，不依赖执行策略）
requirements.txt
LICENSE
.gitignore
```

## 版本要点

**V1.3.4 修正了一处会静默改错页面内容的缺陷**：V1.3.3 及更早版本按「同像素尺寸 + 指纹距离 ≤ 12」挑选 Word 原图回填，当同尺寸桶里存在无关插图时，会把页面图片换成**另一张同尺寸的图**（实测一份 2,865 页文档有 5 页中招），而画质复核因为用同一套判据对照，始终报告「一致」。

现改为双指纹双闸门（粗 16×16 ≤ 1.0 且细 32×32 ≤ 3.0），任一门不过即保留导出图并单列 `no_original` 计数；同时新增**逐页渲染审计**，把导出稿与成品逐页渲染比对，发现结构性差异就不删除中间稿。全量复核：8,328 个合法匹配一个未丢，5 个错配全部拦下，改后页面与干净导出逐字节一致。

**V1.4.0 在流水线中插入「跳转全量校验与补链」**：图片回填与画质审计通过之后、合并总 PDF 之前，对每份交付 PDF 比对 Word 原稿里的全部超链接域（`HYPERLINK \l "书签"` 与原生 `w:anchor`）与 PDF 里实际的 `/Link` 注解，把 WPS 漏转的几条按 Word 原稿的目标页补回去。只增删 `/Link` 注解与相关字典键，不动内容流、图像流和版式，因此不推翻已有的画质结论。

补链的定位规则（都按证据强弱排序，对不上就报告、绝不猜页）：

| 环节 | 规则 |
| --- | --- |
| 判定「缺哪条」 | 以规范化文字为键精确配对，同文按序取用（7 条「（点此返回）」各自对上，不会互相吞）；再用包含式兜底 |
| 目标页 | 同名命名目标 → **同文字其他链接已指向的页（取众数）** → 书签所在标题 → 显示文字 → 编号前缀 → 近似匹配 |
| 热区 | 字符级坐标定位，跨行自动拼接；纵向按 `字号 × 1.5` 展开（实测 WPS 中位比 1.493），与原生注解对齐 |
| 不可定位 | 文字层找不到（纯扫描页）、显示文字过短、外链，一律列进 `unresolved` 报告，不写入 |

实测：一份 1,371 页、21 个域的文档，首次运行发现并补回第 2 页缺失的「14.1.1 产品质量能力」，其热区与相邻 WPS 原生链接偏差 0.1 pt；重跑新增 0 条（幂等）。另做破坏性回归——人为删除 4 条链接（含目录页同文条目、正文条目、两条「（点此返回）」），补链后页号与目标页全部一致、热区覆盖完整。

**V1.5.0 新增「页码回填」与整体交付编排**：投标文件的表格里有「对应页码 / 佐证材料对应页码」这类列，值必须在**导出 PDF 之后**才能确定——它取决于材料在总 PDF 里的位置。于是新增两轮导出：

```
第 1 轮（探测）  全部源文件 → PDF（页码仍空）→ 得到各分册页数
页码求解        总页码 = 前面所有分册页数之和 + 该材料在本分册内的页序号
回填           写 docx 副本，只改目标单元格，不动字体/字号/颜色/下划线
第 2 轮（定稿）  重导 → 收敛检测（页数若变则后续分册整体平移，重算，最多 3 轮）
图像回填 → 跳转补链 → 合并（并把各分册书签按页偏移平移重建）
```

不必先合出一份几百 MB 的总 PDF 再回头改，这是「两轮导出」的关键。页码定位不依赖 docx 的自动编号，而是读分册 PDF 正文里「编号 + 标题」的行，按三级证据定页：行内文字唯一命中标题 → 表格列 N 行对应 `X.1..X.N` 的结构匹配（缺号按相邻页插值）→ 行内明写的小节号直查。全部失败就留空并写进明细，绝不填一个「看起来差不多」的页。

实测 23 个分册（22 docx + 1 xlsx）共 11,335 页：求解 1,201 格、未解析 59 格，回填未改变任何分册页数（1 轮收敛），产出总 PDF 397.8 MB 并重建 11,703 条书签（此前合并结果的大纲是 **0 条**）。抽查 28/30/31/32 号四份分册的填入值与总 PDF 实际页面逐一对应。

xlsx 走另一套 COM（WPS 表格注册的是 `KET.Application`，`KWps.Application` 没有 `Workbooks`），导出 PDF 后与 docx 同等对待。

**V1.5.1 修正 JPEG 色彩空间错标**：`/DCTDecode` 流的 `/ColorSpace` 必须与 JPEG 自身分量数一致（1→`/DeviceGray`、3→`/DeviceRGB`、4→`/DeviceCMYK`）。旧代码在 JPEG 直通分支里硬编码 `/DeviceRGB`，而真实投标材料里有 895 张单通道灰度扫描件，被声明成 RGB 后：深底页整页变黑、浅底页整页变红（同一根因两种表现，取决于原图底色）。

修正分三层：① 直通分支用 `jpeg_components()` 读 SOF 段按真实分量声明；② 新增 `audit_jpeg_colorspace()` 对整份 PDF 全量核对改正——这层不可省，因为错标是 WPS 带进来的，回填因尺寸不匹配而保留导出图时根本碰不到；③ `quality_check()` 加硬断言，不符即阻止删除中间稿。

原有两道校验都抓不到它：画质复核只比字节（尺寸不匹配未进比对），逐页渲染审计两侧按同一错误声明解码、错得一样、差值归零。

## 快速开始

```powershell
# 1) 环境：Windows + Python 3.9+；本机需装有 WPS Office 或 Microsoft Office
python -m pip install -r requirements.txt

# 2) 环境自检（确认 COM 可用、能导出）
python src\wps_to_pdf.py --check

# 3) 单个文档（导出 + 优化 + 校验，默认 Word COM 通道用 --engine word）
python src\wps_to_pdf.py "A.docx" --engine word --optimize

# 4) 图形界面
python src\pdf_tool.py
```

批量转换请直接用 GUI 的「整个文件夹（批量）」模式，或在代码中调用
`pdf_tool.run_batch()`；详见 [docs/使用说明.md](docs/使用说明.md)。

### 只对已有 PDF 做跳转校验与补链

不必重跑转换，可单独调用链接层：

```powershell
# 有 Word 原稿：校验并就地补链（改的是 PDF 里的 /Link 注解）
python src\link_fix.py --docx A.docx --pdf A.pdf

# 只看报告不动文件
python src\link_fix.py --docx A.docx --pdf A.pdf --dry-run --json report.json

# 没有 Word 原稿：只审既有注解是否可用（目标页越界、热区退化）
python src\link_fix.py --pdf A.pdf

# 与主程序同一个入口
python src\pdf_tool.py --links A.pdf A.docx --dry-run --json report.json
```

## 三种导出通道

| 通道 | 适用条件 | 说明 |
| --- | --- | --- |
| `wps` | 装有 WPS Office 且 COM 注册正常 | 默认；无会员要求，与手工导出结果一致 |
| `word` | 装有 Microsoft Office | 未安装 WPS 时使用；版式由 Word 渲染 |
| `cli` | 已登录 WPS VIP 账号 | 官方 `wpscli word2pdf`，首次运行需联网下载组件，输入上限 200 MB |

## 优化策略（体积优先）

默认策略是**不损失画质的前提下尽可能小**：

- 二值图 → 1bpp；索引图按调色板规模选最小位深（2/4/8bpp）
- 超过 256 色的图 → 无损 `/DeviceRGB` + Flate
- 与 Word 原图逐字节相同的 JPEG → 原样保留
- **若某个对象的无损编码比 Word 存的原始字节更大，则保留原件**（省体积），
  日志单独记数。需要强制全部无损时加 `--lossless-all`。

详细论证与实测数据见 [docs/使用说明.md](docs/使用说明.md) 第 11 节。

## 构建可执行文件

```powershell
.\tools\build_exe.ps1
```

产物：`dist\WordToPdfTool.exe`（单文件，免 Python 环境，复制到任意装了 WPS/Office 的 Windows 机器即可运行）。
`dist` 与 `build` 目录不入库（体积过大且为可再生产物），需要发布包时按上表构建后用 `Compress-Archive` 打包即可。

产物为单文件 `WordToPdfTool.exe`。构建后建议跑一次自测：

```powershell
# PowerShell 版（若执行策略受限，用 python 版）
.\tools\selftest.ps1 -Sample "A.docx" -Engine word

# 跨平台版，不依赖 PowerShell 执行策略
python tools\selftest.py "A.docx" --engine word
```

自测会真实转换一份文档，并断言：产出 PDF 存在、页数非零、中间 PDF 已清理、
画质复核像素差异为 0。

## 运行环境与依赖

- Windows 10/11（依赖 Win32 API 与 COM）
- Python 3.9 及以上（打包版无需 Python）
- 依赖：`pywin32`、`PyMuPDF`、`Pillow`（见 `requirements.txt`）

## 已知限制

- 仅接受 `.docx`；`.doc` / `.wps` 需先另存为 `.docx`（优化阶段需读取包内原始图片）
- 需要本机安装 WPS Office 或 Microsoft Office
- 单个文件上限 2 GB（COM 通道）；WPS CLI 通道官方上限 200 MB
- 已加密的 PDF 无法合并

## 许可证

本项目自有源码采用 **MIT**，Copyright (c) 2026 **Vizzini**，见 `LICENSE`。

`LICENSE` 只放纯 MIT 正文——GitHub 的许可证识别器要求文件中不含追加段落，
否则会把仓库标成「Other」。第三方授权因此全部写在 `THIRD-PARTY-NOTICES.md`。

第三方组件另行授权，其中 **PyMuPDF 是 AGPL-3.0（或 Artifex 商业授权）**——
打包出的 exe 内嵌了它，因此该 exe 的分发受 AGPL-3.0 约束，不能只以 MIT 名义
对外交付。其余依赖：Pillow（MIT-CMU/HPND）、pywin32（PSF-2.0）、
lxml（BSD-3-Clause）。

完整说明与分发时的建议做法见 `THIRD-PARTY-NOTICES.md`。

工具本身不联网、不上传文档、无遥测；但需要目标机已安装 WPS Office 或
Microsoft Office，其授权需自行具备。
