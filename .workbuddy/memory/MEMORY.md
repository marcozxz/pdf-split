# MEMORY.md

## 项目: PDF智能拆分工具集

### 核心文件
- `pdf_split_tool.py` — PDF分类拆分工具 v2.0 (得分制+上下文关联)
- `pdf_page_split_tool.py` — PDF按页拆分命名工具
- `build_exe.py` — 统一打包脚本，支持 `python build_exe.py [分类拆分|按页拆分|全部]`
- `dist/` — exe输出目录

### 关键技术决策
- OCR引擎: RapidOCR (rapidocr-onnxruntime)，不用Tesseract/EasyOCR
- PyInstaller打包参数: `--onefile --windowed --collect-data=rapidocr_onnxruntime`
- PyInstaller中需用 `sys._MEIPASS` 处理RapidOCR数据路径
- PowerShell无法直接捕获Python stdout，需用 `Start-Process -RedirectStandardOutput`

### 分类算法 v2.0 (2026-04-30优化)
- **得分制**: 主关键字+10分(长词额外加分)，组合关键字+3分/个(需2+命中)
- **组合关键字**: 用于Excel表格等OCR碎片化场景(如付款申请单)
- **上下文关联**: 孤岛修正、连续性修正、未分类页跟随修正
- 付款意见书关键字需包含"审核结果/审核情况"以识别末页

### 用户PDF信息
- 源文件: 20260424142335.pdf (40页扫描PDF)
- 7类: 付款申请单(p1)、发票(p2)、支付申请表(p3)、付款意见书(p4-6)、支付情况统计表(p7)、合同文件(p8-13)、检测报告(p14-40)

### Python路径
- Windows: `C:\Python3\python.exe` (3.10.7)
