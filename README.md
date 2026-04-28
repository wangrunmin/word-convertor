# word-convertor

中文文档批量入库流水线：把任意 Office 文档转成 `.docx`，再提取为下游评审服务消费的 JSON。

```
任意文档 (.doc / .wps / .ofd / .rtf / .txt / .pdf)
    │  convert2docxByWps.ps1   （Windows + WPS Office）
    ▼
.docx
    │  export_docx_to_json.py  （Python，跨平台）
    ▼
JSON：{ judgeId, errorCorrectionType, paragraphs:[{index, content}] }
```

两个阶段相互独立，可以单独使用。

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

## 仓库结构

```
word-convertor/
├── convert2docxByWps.ps1   # 阶段一：WPS COM 批量转 docx
├── export_docx_to_json.py  # 阶段二：docx → JSON
├── lib/                    # 首次运行后自动出现，缓存 PdfPig DLL
└── CLAUDE.md               # 给 Claude Code 的项目说明
```
