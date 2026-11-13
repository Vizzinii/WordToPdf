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
  wps_to_pdf.py            导出层：WPS COM / WPS CLI / Microsoft Word COM
  optimize_pdf_images.py   优化层：原图回填、无损重编码、完整性校验
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

MIT，见 `LICENSE`。
