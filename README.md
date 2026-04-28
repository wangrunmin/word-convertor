# word-convertor

中文文档批量入库流水线：把任意 Office 文档转成 `.docx`，再提取为 JSON，驱动多个筛查 HTTP 接口采集原始返回，最终导出 Excel 报告。

```
任意文档 (.doc / .wps / .ofd / .rtf / .txt / .pdf)
    │  convert2docxByWps.ps1         （Windows + WPS Office）
    ▼
.docx
    │  export_docx_to_json.py        （Python，跨平台）
    ▼
JSON：{ judgeId, errorCorrectionType, paragraphs:[{index, content}] }
    │  run_screening.py               （Python，跨平台）
    ▼
screening.sqlite：每篇 × 每个接口的原始 HTTP 返回
    │  export_screening_report.py     （Python，跨平台）
    ▼
筛查报告.xlsx：总览 / 筛查结果 / 问题明细 / 失败列表
```

四个阶段相互独立，可以单独使用。

---

## 阶段一：批量转 docx（`convert2docxByWps.ps1`）

仅 Windows 可用，依赖本机已安装的 **WPS Office**（不是 Microsoft Word）。
首次运行会自动尝试安装可选能力，失败也不影响主流程：

| 可选能力 | 用途 | 失败时 |
| --- | --- | --- |
| `PdfPig`（NuGet 自动下载到 `lib/`） | 区分 PDF 文字版/扫描版/混合 | PDF 仅标记不分类 |
| `PyMuPDF + python-docx`（pip） | 文字版 PDF → docx | 跳过 PDF 转换 |
| `ImportExcel`（PSGallery） | 生成 xlsx 报告 | 回退 CSV |

### 用法

交互式（推荐首次使用，会一步步引导）：

```powershell
powershell -ExecutionPolicy Bypass -File .\convert2docxByWps.ps1
```

命令行参数：

```powershell
powershell -ExecutionPolicy Bypass -File .\convert2docxByWps.ps1 `
    -InputPath  "D:\input"  `   # 文件 / .zip 压缩包 / 文件夹 三选一
    -OutputDir  "D:\output" `
    -ThrottleLimit 4        `   # 并行 Worker 数；超过 ~6 后 WPS COM 不稳定
    -Recurse                `   # 递归子目录（默认开启）
    -CleanTemp $true            # 自动清理 zip 解压临时目录（不传则交互询问）
```

### 输入支持

- **目录**：递归扫描 `.doc / .docx / .wps / .ofd / .rtf / .txt / .pdf`
- **`.zip` 压缩包**：自动用 GBK（codepage 936）解压，避免国内压缩包的中文文件名乱码
- **单文件**：复制到临时目录后按目录模式处理

### 输出

全部落在 `-OutputDir` 下：

- `.docx`（保留相对路径结构）
- `convert.log`：所有 Worker 共用的日志
- `convert_report_<时间戳>.xlsx`：带筛选/冻结首行/状态着色的报告（缺 ImportExcel 时回退为 `.csv`）
- `convert_report_<时间戳>.html`：自带筛选和搜索的网页版报告
- `failed_list_<时间戳>.txt`：仅失败文件路径，便于二次处理

### 常见报错

- **"Worker 启动失败 / 找不到 WPS COM 组件"**：以管理员身份手动启动一次 WPS，触发 COM 注册。
- **所有 Worker 卡死、60 秒无进度**：脚本会自动检测并强制终止；先用任务管理器结束残留的 `wps.exe / wpp.exe / wpsoffice.exe`，再重跑。
- **xlsx 生成失败**：自动回退到 CSV，无需处理。

---

## 阶段二：docx → JSON（`export_docx_to_json.py`）

跨平台，依赖：

```bash
pip install python-docx tqdm
```

> ⚠️ **不支持 `.doc` / `.wps`**，必须先用阶段一转成 `.docx`。脚本扫描时会提示跳过这些后缀但不会报错。

### 用法

交互式：

```bash
python export_docx_to_json.py
# 依次输入：输入目录 → 输出目录 → 是否递归 → 并行进程数
```

命令行参数：

```bash
python export_docx_to_json.py \
    --input   D:/output \
    --output  D:/json   \
    --workers 8         \
    --recurse
```

### 输出 JSON 格式

每个 `.docx` 输出同名 `.json`（保留相对路径）：

```json
{
    "judgeId": 114,
    "errorCorrectionType": 0,
    "paragraphs": [
        { "index": 0, "content": "第一句。" },
        { "index": 1, "content": "第二句！\r" }
    ]
}
```

字段说明：

- `judgeId` / `errorCorrectionType`：**与下游评审服务约定的固定常量，不要随意修改**。
- `paragraphs[].content`：按 `。！？` 切分后的句子；段末最后一句尾部追加 `\r` 作为段落边界标记。
- 连续重复句会被去重（OCR/扫描件中常见的重复页眉）。

失败列表会写到 `<输出目录>/_failed.txt`。

---

## 完整工作流示例

```powershell
# 1. 把混杂格式的源文档全部转成 docx
powershell -ExecutionPolicy Bypass -File .\convert2docxByWps.ps1 `
    -InputPath "D:\原始文档" -OutputDir "D:\docx" -ThrottleLimit 4

# 2. 再把 docx 提取成下游需要的 JSON
python export_docx_to_json.py -i D:\docx -o D:\json -w 8 -r
```

---

---

## 阶段三：筛查采集（`run_screening.py`）

跨平台，依赖：

```bash
pip install httpx PyYAML tqdm   # tqdm 可选，缺失自动降级
```

### 快速上手

**不传任何参数直接运行**，进入交互向导（推荐新手）：

```bash
python run_screening.py
```

向导会引导完成：接口配置生成（首次）→ 输入目录 → 输出目录 → 接口选择 → 前台/后台运行。

### 命令行用法

```bash
# 前台运行（看进度条）
python run_screening.py run -i json/ -o screen_out/

# 只跑指定接口（接口名来自 screening_config.yaml 的 key）
python run_screening.py run -i json/ -o screen_out/ --interface A,B

# 后台运行（SSH 断开也继续）
python run_screening.py run -i json/ -o screen_out/ --detach

# 复用上次的交互选择
python run_screening.py run --replay

# 强制重跑（默认是断点续跑）
python run_screening.py run -i json/ -o screen_out/ --force
```

### 后台任务管理

```bash
# 查看进度
python run_screening.py status -o screen_out/

# 实时日志
python run_screening.py logs -o screen_out/ -f

# 优雅停止（当前批次写完后退出）
python run_screening.py stop -o screen_out/
```

### 健康检查

```bash
python run_screening.py doctor
# ✓ Python 3.11
# ✓ httpx / PyYAML 已安装
# ○ tqdm 未安装（无进度条，不影响主流程）
# ✓ screening_config.yaml 存在
# ✓ 接口 A: SCREENING_A_TOKEN 已设置
# ✗ 接口 B: connect timeout（可能原因：服务未启动 / VPN 未连）
```

### 输出说明

所有产物落在 `-o` 指定的输出目录：

```
screen_out/
├── screening.sqlite     # 主存储（每篇 × 每接口的原始返回）
├── screening.log        # 日志
├── _stats.json          # 实时进度统计
└── _failed.txt          # 失败任务（doc\tinterface\tstatus_code\treason）
```

`screening.sqlite` 的 schema：

```sql
CREATE TABLE screening_result (
    doc_path     TEXT NOT NULL,   -- 输入 JSON 相对路径
    interface    TEXT NOT NULL,   -- 接口名（yaml 中的 key）
    status       TEXT NOT NULL,   -- success | failed
    raw_response TEXT,            -- 接口原始 JSON 字符串
    error        TEXT,
    http_status  INTEGER,
    attempts     INTEGER NOT NULL,
    elapsed_ms   INTEGER NOT NULL,
    fetched_at   TEXT NOT NULL,
    PRIMARY KEY (doc_path, interface)
);
```

**断点续跑**：任何中断后直接重跑，已成功的行自动跳过。

### 下游取数据示例（SQL）

```python
import sqlite3, json
conn = sqlite3.connect("screen_out/screening.sqlite")

# 取某接口全部成功结果
rows = conn.execute(
    "SELECT doc_path, raw_response FROM screening_result WHERE interface=? AND status='success'",
    ("A",)
).fetchall()
for doc_path, raw in rows:
    data = json.loads(raw)
    # ... 解析 data

# 统计各接口成功/失败数
conn.execute("""
    SELECT interface, status, COUNT(*) as cnt
    FROM screening_result GROUP BY interface, status
""").fetchall()
```

### 接口配置（`screening_config.yaml`）

接口数量、URL、鉴权、并发数全部由此文件决定，代码内无硬编码。首次运行会自动生成配置模板，也可以手动编辑：

```yaml
interfaces:
  A:
    url: "${SCREENING_A_URL}"      # 支持 ${ENV_VAR} 占位
    method: POST
    auth:
      type: bearer
      token_env: SCREENING_A_TOKEN
    workers: 8
    timeout: 30
    enabled: true
```

真实 URL / token 存入环境变量，不要写死在配置文件里。

### 导出子命令

```bash
# 导出某篇所有接口的原始返回为文件
python run_screening.py dump --doc path/to/doc.json -o exported/

# 导出某接口全部失败任务
python run_screening.py dump --interface A --status failed -o failed_A/

# 导出全部
python run_screening.py dump --all -o exported/
```

---

## 阶段四：导出报告（`export_screening_report.py`）

跨平台，依赖：

```bash
pip install openpyxl PyYAML tqdm   # tqdm 可选
```

### 快速上手

```bash
# 不传参数 → 交互向导
python export_screening_report.py

# 直接运行（-i 指定阶段二 JSON 目录，首次运行时提取法院/案号并缓存到 SQLite）
python export_screening_report.py -d screen_out/screening.sqlite -i json/ -o 筛查报告.xlsx

# 刷新法院/案号缓存（提取规则改了时用）
python export_screening_report.py -i json/ --rescan-meta

# 过滤导出
python export_screening_report.py --interface A --status success

# 强制 CSV（openpyxl 缺失时自动降级，也可手动指定）
python export_screening_report.py --format csv
```

### 输出说明

默认输出 `筛查报告_<时间戳>.xlsx`，包含四个 Sheet：

| Sheet | 内容 |
| --- | --- |
| 总览 | 文档数、各接口成功/失败/问题数统计 |
| 筛查结果 | 每篇 × 每接口一行，含状态/耗时/问题数 |
| 问题明细 | 每条问题一行，18 列统一 schema（含法院、案号） |
| 失败列表 | 仅失败行 + 原因 |

xlsx 写入失败时自动降级为 CSV（UTF-8-BOM，Excel 双击不乱码）。

### 接口 → 解析器映射

在 `screening_config.yaml` 每个接口下加一行 `parser` 字段，导出器据此分发解析逻辑：

```yaml
interfaces:
  A:
    parser: keypoint      # 重点纠错 / 风险防控（解析 corrections[]）
  B:
    parser: law_quote     # 法律纠错（解析 lawQuotes[]）
```

未配置 `parser` 的接口只进"筛查结果" Sheet，不进"问题明细"。

### 法院 / 案号

首次运行时（需要 `-i` 参数），脚本扫描全部 JSON 文件，用宽松正则从前 20 段提取法院名和案号，缓存到 `screening.sqlite` 的 `document_meta` 表。后续导出直接 JOIN，不再读 JSON 文件。

---

## 仓库结构

```
word-convertor/
├── convert2docxByWps.ps1          # 阶段一：WPS COM 批量转 docx
├── export_docx_to_json.py         # 阶段二：docx → JSON
├── run_screening.py               # 阶段三：筛查接口采集器
├── export_screening_report.py     # 阶段四：筛查结果导出报告
├── screening_config.yaml          # 阶段三/四接口配置模板（可提交）
├── requirements.txt               # Python 依赖
├── lib/                           # 首次运行后自动出现，缓存 PdfPig DLL
└── CLAUDE.md                      # 给 Claude Code 的项目说明
```
