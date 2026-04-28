# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 仓库定位

中文文档批量入库流水线，四阶段：

```
任意文档 (.doc/.wps/.ofd/.rtf/.txt/.pdf)
    │  convert2docxByWps.ps1         （Windows + WPS Office）
    ▼
.docx
    │  export_docx_to_json.py        （Python）
    ▼
JSON: { judgeId, errorCorrectionType, paragraphs:[{index, content}] }
    │  run_screening.py               （Python）
    ▼
screening.sqlite：每篇 × 每接口的原始 HTTP 返回
    │  export_screening_report.py     （Python）
    ▼
筛查报告.xlsx：总览 / 筛查结果 / 问题明细 / 失败列表
```

阶段一只能在 Windows 上跑（依赖 WPS COM）；阶段二、三、四跨平台。
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

### 阶段三 — 筛查采集（Python 3）

```bash
pip install httpx PyYAML tqdm

# 不传参数 → 交互向导（生成配置、选目录、选接口、选前台/后台）
python run_screening.py

# 直接运行
python run_screening.py run -i json/ -o screen_out/

# 后台运行
python run_screening.py run -i json/ -o screen_out/ --detach

# 查进度 / 看日志 / 停止
python run_screening.py status -o screen_out/
python run_screening.py logs   -o screen_out/ -f
python run_screening.py stop   -o screen_out/

# 健康检查（依赖、环境变量、接口连通性）
python run_screening.py doctor

# 导出 SQLite 行为文件（供下游消费）
python run_screening.py dump --all -o exported/
```

### 阶段四 — 导出报告（Python 3）

```bash
pip install openpyxl PyYAML tqdm

# 不传参数 → 交互向导
python export_screening_report.py

# 直接运行（-i 首次运行时扫描 JSON 提取法院/案号并缓存到 SQLite）
python export_screening_report.py -d screen_out/screening.sqlite -i json/ -o 筛查报告.xlsx

# 刷新法院/案号缓存
python export_screening_report.py -i json/ --rescan-meta

# 过滤导出 / 强制 CSV
python export_screening_report.py --interface A --status success
python export_screening_report.py --format csv
```

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

### 阶段三（`run_screening.py`）

定位：**只做采集和存储**，不 merge、不解析接口返回、不出报告。下游读 SQLite 自行处理。

- **接口数由 `screening_config.yaml` 决定**，代码内无任何接口硬编码。接口名是 yaml 顶层 key；`--interface A,B` 是 yaml key 的子集过滤。
- **每个接口独立 worker pool**：`asyncio.Queue(maxsize=workers×4)` + feeder 协程 + N 个 worker 协程，所有接口的 task 共用一个 event loop，互不阻塞。慢接口不会拖累快接口（不存在木桶效应）。
- **单 writer 协程串行写 SQLite**：所有接口 worker 把结果丢进共享 `result_q`（asyncio.Queue），1 个 writer 消费并执行 `INSERT OR REPLACE`。配合 `PRAGMA journal_mode=WAL`，读不阻塞写。
- **`(doc_path, interface)` 主键保证断点续跑**：每个 task 开跑前查主键，`status='success'` 就跳过；`--force` 强制 `INSERT OR REPLACE`。任何中断（Ctrl+C / SIGTERM / OOM kill）后直接重跑接续。
- **`--detach` 后台守护**：POSIX 走 double-fork + `setsid`；Windows 走 `subprocess.Popen(DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP)` 重启自身。PID 写入 `<outdir>/.screening.pid`，`status`/`logs`/`stop` 子命令依赖它。
- **不传参数必然进入交互向导**：`len(sys.argv) == 1` 直接调 `cmd_run(_Stub())`，向导引导生成配置 → 选输入目录 → 选输出目录 → 选接口 → 前台/后台。
- **存储模式**：默认 SQLite（`screen_out/screening.sqlite`）；`--storage=files` 退化为 `<doc>.<interface>.json` 散文件，仅供 <1000 篇小批量调试。`dump` 子命令按需把 SQLite 行导出成文件交给下游。
- **`${ENV_VAR}` 展开**：yaml 加载后递归替换所有字符串中的 `${VAR}` 占位，未设置的环境变量在 `doctor` 里报告，不硬退出（`enabled: false` 的接口不检查）。

### 阶段四（`export_screening_report.py`）

定位：**只做"SQLite → Excel/CSV"**，不发请求、不修改 SQLite（除写入 `document_meta` 缓存外）、不做去重合并。

- **接口 → parser 映射由 `screening_config.yaml` 决定**：每个接口下加 `parser: keypoint`（重点纠错/风险防控，解析 `corrections[]`）或 `parser: law_quote`（法律纠错，解析 `lawQuotes[]`）。未配置 `parser` 的接口只进"筛查结果" Sheet，不进"问题明细"，不会报错。
- **两种 parser 的 18 列 schema 完全对齐**：`court/case_num` 来自 `document_meta`；法院/条文/建议表述 仅 `law_quote` 填，`keypoint` 留空。
- **法院/案号缓存（`document_meta` 表）**：首次导出时扫描 input JSON，宽松正则提取，`INSERT OR IGNORE` 写入同一 SQLite 文件；后续导出 `LEFT JOIN` 复用，`--rescan-meta` 强制重扫。`executemany` 500 条一批写入。
- **openpyxl 降级**：import 失败 → 直接 CSV；`PermissionError`（文件被 Excel 占用）→ 加 `_retry` 后缀重试；任何其他写入异常 → 降级 CSV（`utf-8-sig`，4 个 Sheet 各一个文件）。
- **xlsx 样式**：`_row_styles()` 支持 `status_col`（按值着色）和 `fill_all`（统一着色，失败列表用）；`_autowidth()` 单次遍历所有列（非每列一次全表扫描）；Alignment 对象预缓存（wrap / normal 各一个）。

## 踩坑记录（`export_docx_to_json.py` 文件头，标注 2026-04-28）

下面这些是有意识的"不做"决策，没有强理由不要反着改：

1. 不要试图模拟 WPS 的精确分句逻辑。它带状态且不稳定，上面的正则是已对齐的近似方案。
2. 凡是能直接读 `.docx` zip 完成的工作，都不要走 WPS COM。COM 慢约 100 倍，且所有调用都要排队过同一个 Office 进程。

## 已否决的演进方向（避免反复重提）

每条记录格式：决策 · 日期 · 理由 · 何时重新评估。

### 阶段一（`convert2docxByWps.ps1`）不迁移到 Python · 2026-04-28

**决策**：保留 PowerShell，不用 `pywin32` 重写。

**理由**：
1. **并发模型**：PowerShell `Runspace Pool` 是同进程多线程，每个 Runspace 持长生命周期 STA COM 实例，初始化几乎零开销。Python 多线程受 GIL + COM 线程亲和性约束；多进程每个 worker 都要单独起 WPS COM，启动开销 ×N。此场景 PowerShell 是甜蜜点。
2. **无迁移收益**：PdfPig / ImportExcel / 命名 Mutex / 内联 C# `AtomicCounter` 都有 Python 等价物，但阶段一输出 `.docx`，对下游语言中性，统一技术栈不会让阶段二/三/四变好写。
3. **踩坑不可漏**：COM 死锁 60s 超时、`PK\x03\x04` 魔数嗅探、宏文档伪装 `.docx`、ZIP GBK 解压、可选能力静默降级 —— 全是真实事故留下的补丁，重写要逐条迁移，漏一个就是线上事故，且 bug 现场（真实用户文档）复现成本高。
4. **ROI 为负**：迁移约 1-2 周开发 + 数周稳定性观察，没有"必须 Python"的硬约束（如嵌入 Python 长驻服务）就不值得。

**何时重新评估**：(a) 要把转换能力嵌入更大的 Python 服务；(b) 决定砍掉 WPS、改用纯 Python 库（但届时不再是"转 docx"，而是另一条技术路线）。
