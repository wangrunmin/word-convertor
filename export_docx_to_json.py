"""
export_docx_to_json.py
批量提取 docx 文件为 JSON，多进程并行，速度优先
依赖: pip install python-docx tqdm
不支持 .doc / .wps，需预先转换为 docx

踩坑记录：
2026年4月28日
1、不要试图模拟WPS的分句逻辑，很难且不稳定。
2、非必要不要通过WPS的COM接口进行文档操作，速度很慢。
"""

import os
import re
import json
import sys
import argparse
from pathlib import Path
from multiprocessing import Pool, cpu_count

try:
    from tqdm import tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph as DocxParagraph


# ── 句子分隔符（中文）────────────────────────────────────────
SENT_RE = re.compile(r'(?<=[。！？])')


# ── 获取文档所有段落（含表格内段落，按文档顺序）────────────────
def iter_all_paragraphs(doc):
    for elem in doc.element.body.iter():
        if elem.tag == qn('w:p'):
            yield DocxParagraph(elem, doc)


# ── 单文件处理（顶层函数，multiprocessing 要求可 pickle）────────
def process_file(args):
    file_path, file_index, input_dir, output_dir = args
    file_path  = Path(file_path)
    input_dir  = Path(input_dir)
    output_dir = Path(output_dir)

    try:
        # docm（含宏文档被改名为 docx）：python-docx 会因 content type 拒绝
        # 解决方案：复制文件并修改 [Content_Types].xml 后再打开
        import zipfile, tempfile, shutil as _shutil
        _suffix = file_path.suffix.lower()
        _tmp_path = None
        _src = str(file_path)
        _need_patch = False
        try:
            with zipfile.ZipFile(_src) as _z:
                _ct = _z.read('[Content_Types].xml').decode('utf-8')
                if 'macroEnabled' in _ct:
                    _need_patch = True
        except Exception:
            pass  # 非 zip 格式（如 ~$ 临时文件），后面 Document() 会报错
        if _need_patch:
            _tmp = tempfile.NamedTemporaryFile(suffix='.docx', delete=False)
            _tmp_path = _tmp.name
            _tmp.close()
            _shutil.copy2(_src, _tmp_path)
            with zipfile.ZipFile(_tmp_path, 'r') as _zin:
                _names = _zin.namelist()
                _files = {n: _zin.read(n) for n in _names}
            _files['[Content_Types].xml'] = _files['[Content_Types].xml'].replace(
                b'vnd.ms-word.document.macroEnabled.main+xml', b'vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
            ).replace(
                b'vnd.ms-word.document.macroEnabled.template+xml', b'vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
            )
            with zipfile.ZipFile(_tmp_path, 'w', zipfile.ZIP_DEFLATED) as _zout:
                for _n, _d in _files.items():
                    _zout.writestr(_n, _d)
            _src = _tmp_path
        try:
            doc = Document(_src)
        finally:
            if _tmp_path and os.path.exists(_tmp_path):
                os.unlink(_tmp_path)
    except Exception as e:
        return False, f"{file_path.name} → 打开失败: {e}"

    paragraphs   = []
    global_index = 0
    last_content = None

    for para in iter_all_paragraphs(doc):
        raw = para.text
        if not raw.strip(' \t\n\r\x0b\x0c\xa0\x07'):
            continue

        # 按 。！？ 分句，保留分隔符；段末 \r 加到最后一句末尾
        sentences = [s for s in SENT_RE.split(raw) if s]
        if not sentences:
            sentences = [raw]

        for i, sent in enumerate(sentences):
            is_last = (i == len(sentences) - 1)
            text    = sent + '\r' if is_last else sent
            if text == last_content:
                continue
            paragraphs.append({
                "index":   global_index,
                "content": text,
            })
            global_index += 1
            last_content = text

    result = {
        "judgeId":             114,
        "errorCorrectionType": 0,
        "paragraphs":          paragraphs,
    }

    # 保留相对路径结构
    rel      = file_path.relative_to(input_dir)
    dest_dir = output_dir / rel.parent
    dest_dir.mkdir(parents=True, exist_ok=True)
    out_path = dest_dir / (file_path.stem + '.json')

    try:
        with open(out_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=4)
    except Exception as e:
        return False, f"{file_path.name} → 写入失败: {e}"

    return True, None


# ── 主流程 ────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description='批量提取 docx → JSON')
    parser.add_argument('--input',   '-i', default='', help='输入目录')
    parser.add_argument('--output',  '-o', default='', help='输出目录')
    parser.add_argument('--workers', '-w', type=int, default=0, help='并行进程数')
    parser.add_argument('--recurse', '-r', action='store_true', help='递归子目录')
    args = parser.parse_args()

    print()
    print('=' * 42)
    print('  docx 批量提取工具（Python 并行版）')
    print('=' * 42)
    print()

    # 交互输入
    input_dir = args.input.strip().strip('"')
    if not input_dir:
        input_dir = input('输入目录（支持拖拽）: ').strip().strip('"').rstrip('\\/')
    if not os.path.isdir(input_dir):
        print(f'[错误] 目录不存在: {input_dir}'); sys.exit(1)

    output_dir = args.output.strip().strip('"')
    if not output_dir:
        tmp = input('输出目录（留空则在输入目录下新建 json_output）: ').strip().strip('"').rstrip('\\/')
        output_dir = tmp if tmp else os.path.join(input_dir, 'json_output')
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    recurse = args.recurse
    if not args.recurse:
        r = input('是否递归子目录？(Y/N，默认 Y): ').strip().lower()
        recurse = (r == '' or r == 'y')

    cpu = cpu_count()
    default_workers = min(max(4, cpu // 2), 16)
    workers = args.workers
    if workers <= 0:
        w = input(f'并行进程数（CPU {cpu} 核，建议 {default_workers}，直接回车默认）: ').strip()
        workers = int(w) if w.isdigit() and int(w) > 0 else default_workers

    # 收集文件
    input_path = Path(input_dir)
    exts = {'.docx', '.docm'}  # docm 与 docx 结构相同
    if recurse:
        files = [p for p in input_path.rglob('*') if p.suffix.lower() in exts and p.is_file() and not p.name.startswith('~$')]
    else:
        files = [p for p in input_path.iterdir() if p.suffix.lower() in exts and p.is_file() and not p.name.startswith('~$')]

    # 检查是否有 .doc/.wps 并提示
    skipped_exts = set()
    if recurse:
        skipped_exts = {p.suffix.lower() for p in input_path.rglob('*')
                        if p.suffix.lower() in {'.doc', '.wps'} and p.is_file()}
    if skipped_exts:
        print(f'[提示] 跳过不支持的格式: {", ".join(skipped_exts)}（需预先用 WPS 转换为 docx）')

    if not files:
        print('[提示] 未找到 docx 文件'); sys.exit(0)

    workers = min(workers, len(files))
    print()
    print(f'文件数: {len(files)}  并行进程: {workers}')
    print()

    # 构造任务列表
    tasks = [
        (str(f), i + 1, str(input_dir), str(output_dir))
        for i, f in enumerate(files)
    ]

    success = 0
    failed  = 0
    errors  = []

    import time
    t0 = time.time()

    with Pool(processes=workers) as pool:
        if HAS_TQDM:
            it = tqdm(
                pool.imap_unordered(process_file, tasks),
                total=len(tasks),
                unit='篇',
                dynamic_ncols=True,
            )
        else:
            it = pool.imap_unordered(process_file, tasks)
            done = 0

        for ok, err in it:
            if ok:
                success += 1
            else:
                failed += 1
                errors.append(err)

            if not HAS_TQDM:
                done += 1
                elapsed = time.time() - t0
                speed   = done / elapsed if elapsed > 0 else 0
                remain  = int((len(tasks) - done) / speed) if speed > 0 else 0
                print(f'\r[{done}/{len(tasks)}]  {speed:.1f} 篇/s  剩余约 {remain}s   ',
                      end='', flush=True)

    elapsed = time.time() - t0
    print()
    print()
    print('=' * 42)
    print(f'  完成')
    print('=' * 42)
    print(f'成功: {success}  失败: {failed}  耗时: {elapsed:.1f}s  '
          f'速度: {success/elapsed:.1f} 篇/s')
    print(f'输出目录: {output_dir}')

    if errors:
        print()
        print('[失败列表]')
        for e in errors:
            print(f'  {e}')
        log_path = Path(output_dir) / '_failed.txt'
        log_path.write_text('\n'.join(errors), encoding='utf-8')
        print(f'已保存到: {log_path}')

    print()


if __name__ == '__main__':
    main()