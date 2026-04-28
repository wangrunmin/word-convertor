# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 仓库定位

中文文档批量入库流水线，两阶段：

```
任意文档 (.doc/.wps/.ofd/.rtf/.txt/.pdf)
    │  convert2docxByWps.ps1   （Windows + WPS Office）
    ▼
.docx
    │  export_docx_to_json.py  （Python）
    ▼
JSON: { judgeId, errorCorrectionType, paragraphs:[{index, content}] }
```

阶段一只能在 Windows 上跑（依赖 WPS COM）；阶段二跨平台，但通常消费阶段一的产物。
用户面向的使用说明在 `README.md`。

## 常用命令

### 阶段一 — 批量转 docx（PowerShell 5.1）

```powershell
# 交互式（提示输入路径）
powershell -ExecutionPolicy Bypass -File .\convert2docxByWps.ps1

# 非交互式
powershell -ExecutionPolicy Bypass -File .\convert2docxByWps.ps1 `
    -InputPath  "D:\input"  `   # 支持文件 / .zip / 目录
    -OutputDir  "D:\output" `
    -ThrottleLimit 4        `   # WPS Worker 数；超过 ~6 后 COM 不稳定
    -CleanTemp $true            # 自动清理 zip 解压临时目录
```

报告产出在 `$OutputDir`：`convert.log`、`convert_report_<ts>.{xlsx,csv,html}`、`failed_list_<ts>.txt`。

### 阶段二 — docx → JSON（Python 3）

```bash
pip install python-docx tqdm

python export_docx_to_json.py \
    --input  D:/output \
    --output D:/json   \
    --workers 8        \
    --recurse
```

不传参数则进入交互模式，逐项询问。

## 架构要点

### 阶段一（`convert2docxByWps.ps1`，约 1100 行）

- **用 WPS Office，不是 Microsoft Word。** ProgID 是 `KWps.Application`，回退 `Wps.Application`。本机必须装好 WPS 且 COM 已注册（必要时用管理员身份打开一次 WPS 触发注册）。
- **用 Runspace Pool，不用 `ForEach-Object -Parallel`。** 每个 Worker 自己持有一个长生命周期的 `$wps` COM 实例，从 `ConcurrentQueue` 取任务。每个文件起一个 WPS 实在太慢，且 `-Parallel` 不提供 COM 需要的线程级状态隔离。Apartment State 必须是 `STA`。
- **原子计数器走内联 C#（`AtomicCounter`）。** PowerShell 的 hashtable 自增在 Runspace Pool 下会丢更新。
- **日志写入靠命名 Mutex（`WpsConvertLog`）互斥。** 每个 Worker 调用时重新打开 mutex。
- **后缀魔数嗅探（`Get-RealFormatInner`）。** 后缀是 `.docx` 但前几字节不是 `PK\x03\x04`（OOXML）的文件，会强制走 WPS 转换 —— 用户的源文档里大量 `.doc`/RTF 被错误改名。
- **死锁检测。** `$cntDone` 60 秒不前进就强制退出主循环。**不要删** —— WPS COM 会静默卡死。
- **可选能力首次运行自动安装、失败安静降级**（脚本必须在没有这些能力的情况下也能跑）：
  - `lib/` 下的 PdfPig DLL（通过 `nuget.exe` 下载）→ 区分 PDF 文字版/扫描版
  - `PyMuPDF + python-docx`（通过 `pip`）→ 文字版 PDF → docx
  - `ImportExcel`（PSGallery）→ xlsx 报告（缺失时回退 CSV）
- **三种输入形态**：目录、`.zip`（用 GBK 即 codepage 936 解压，修复中文压缩包乱码）、单文件（复制到临时目录，让后续逻辑统一面对"目录"）。

### 阶段二（`export_docx_to_json.py`）

- **输出 schema 与下游强耦合**：`{"judgeId": 114, "errorCorrectionType": 0, "paragraphs": [...]}`。两个常量是评审服务约定值，**不在本仓库内可配置**，改动前必须和下游对齐。
- **句子切分**：`re.compile(r'(?<=[。！？])')` —— 仅按中文句号/叹号/问号切分，保留分隔符。每个段落最后一句尾部追加 `\r`（下游用作段落边界标记）。
- **跨段连续去重**（`last_content` 检查）会丢掉相邻的相同句子，主要是为了清掉扫描/OCR 文档里重复的页眉。
- **遍历 `doc.element.body.iter()` 过滤 `w:p`**，这样表格内的段落也会按文档顺序进入结果；`doc.paragraphs` 单独一个属性会漏掉表格里的段落。
- **`.docm` 被改名为 `.docx` 的兼容补丁**：python-docx 看到 `[Content_Types].xml` 里的 `macroEnabled` 会拒绝打开。脚本把 zip 复制到临时文件，把宏文档相关的 content type 字符串替换成普通 docx 的，再打开。**不要简化掉这段**，真实输入里这种文件很常见。
- **空白字符集是 `' \t\n\r\x0b\x0c\xa0\x07'`**：`\xa0`（不间断空格）和 `\x07`（BEL，WPS 转换出来的表格单元格分隔符）都是有意保留的。
- 文件名以 `~$` 开头的（Office 锁文件）扫描阶段就会过滤；`.doc` 和 `.wps` 这里**不支持**，必须先经阶段一处理 —— 脚本会提示跳过的后缀但不会报错。

## 踩坑记录（`export_docx_to_json.py` 文件头，标注 2026-04-28）

下面这些是有意识的"不做"决策，没有强理由不要反着改：

1. 不要试图模拟 WPS 的精确分句逻辑。它带状态且不稳定，上面的正则是已对齐的近似方案。
2. 凡是能直接读 `.docx` zip 完成的工作，都不要走 WPS COM。COM 慢约 100 倍，且所有调用都要排队过同一个 Office 进程。
