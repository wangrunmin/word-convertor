"""
run_screening.py
筛查接口数据采集器 — 阶段三

把阶段二产出的 JSON 文件批量发送给 N 个 HTTP 筛查接口，
把原始返回稳健地存入 SQLite，支持断点续跑、后台守护、新人向导。

依赖：pip install httpx PyYAML
可选：pip install tqdm

用法（不带参数进入交互向导）：
  python run_screening.py

  python run_screening.py run     -i json/ -o screen_out/
  python run_screening.py status  -o screen_out/
  python run_screening.py logs    -o screen_out/ [-f]
  python run_screening.py stop    -o screen_out/ [--force]
  python run_screening.py doctor
  python run_screening.py dump    --all -o screen_out/ --export-dir exported/
"""

import argparse
import asyncio
import json
import logging
import os
import re
import signal
import sqlite3
import subprocess
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

try:
    import httpx
    HAS_HTTPX = True
except ImportError:
    HAS_HTTPX = False

try:
    import yaml
    HAS_YAML = True
except ImportError:
    HAS_YAML = False

try:
    from tqdm import tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False


# ── 常量 ─────────────────────────────────────────────────────────────────────
VERSION          = "1.0"
DB_FILENAME      = "screening.sqlite"
LOG_FILENAME     = "screening.log"
PID_FILENAME     = ".screening.pid"
STATS_FILENAME   = "_stats.json"
FAILED_FILENAME  = "_failed.txt"
LAST_RUN_FILE    = ".last_run.yaml"
CONFIG_FILE       = "screening_config.yaml"
CONFIG_FILE_LOCAL = "screening_config.local.yaml"
DEFAULT_WORKERS  = 4
DEFAULT_TIMEOUT  = 30
DEFAULT_RETRY    = {"max": 3, "backoff_base": 1.0}


# ── 工具函数 ──────────────────────────────────────────────────────────────────

def die(msg: str):
    print(f"\n  [错误] {msg}", file=sys.stderr)
    sys.exit(1)

ENV_RE = re.compile(r'\$\{([^}]+)\}')

def expand_env(val) -> object:
    """递归把 ${VAR} 替换成环境变量值（找不到就保留原文）"""
    if isinstance(val, str):
        return ENV_RE.sub(lambda m: os.environ.get(m.group(1), m.group(0)), val)
    if isinstance(val, dict):
        return {k: expand_env(v) for k, v in val.items()}
    if isinstance(val, list):
        return [expand_env(v) for v in val]
    return val

def find_missing_env(yaml_text: str) -> List[str]:
    """扫描 yaml 文本，返回引用了但未设置的环境变量名（去重）"""
    names = ENV_RE.findall(yaml_text)
    return [n for n in dict.fromkeys(names) if os.environ.get(n) is None]

def resolve_config_path() -> Path:
    """优先使用 .local.yaml，不存在时回退到普通配置。"""
    local = Path(CONFIG_FILE_LOCAL)
    return local if local.exists() else Path(CONFIG_FILE)

def load_config(config_path: Path) -> Dict:
    if not HAS_YAML:
        die("缺少 PyYAML，请先：pip install PyYAML")
    raw = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
    return expand_env(raw)

def atomic_write_json(path: Path, data: dict):
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


# ── 日志 ─────────────────────────────────────────────────────────────────────

_logger: Optional[logging.Logger] = None

def setup_logger(log_path: Path):
    global _logger
    _logger = logging.getLogger("screening")
    _logger.setLevel(logging.DEBUG)
    if _logger.handlers:
        return
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    ))
    _logger.addHandler(fh)

def log(msg: str, level: str = "info"):
    if _logger:
        getattr(_logger, level)(msg)


# ── 数据库 ────────────────────────────────────────────────────────────────────

def open_db(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(str(db_path), check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS screening_result (
            doc_path     TEXT NOT NULL,
            interface    TEXT NOT NULL,
            status       TEXT NOT NULL,
            raw_response TEXT,
            error        TEXT,
            http_status  INTEGER,
            attempts     INTEGER NOT NULL DEFAULT 0,
            elapsed_ms   INTEGER NOT NULL DEFAULT 0,
            fetched_at   TEXT NOT NULL,
            PRIMARY KEY (doc_path, interface)
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_status    ON screening_result(status)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_interface ON screening_result(interface)")
    conn.commit()
    return conn

def load_done_set(conn: sqlite3.Connection) -> Set[Tuple[str, str]]:
    """加载所有已成功的 (doc_path, interface)，用于启动期快速跳过"""
    rows = conn.execute(
        "SELECT doc_path, interface FROM screening_result WHERE status='success'"
    ).fetchall()
    return set(rows)

def db_write(conn: sqlite3.Connection, r: Dict):
    conn.execute("""
        INSERT OR REPLACE INTO screening_result
            (doc_path, interface, status, raw_response, error,
             http_status, attempts, elapsed_ms, fetched_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        r["doc_path"], r["interface"], r["status"],
        r.get("raw_response"), r.get("error"), r.get("http_status"),
        r["attempts"], r["elapsed_ms"], r["fetched_at"],
    ))
    conn.commit()

def query_stats_from_db(db_path: Path) -> Dict:
    if not db_path.exists():
        return {}
    conn = sqlite3.connect(str(db_path))
    conn.execute("PRAGMA journal_mode=WAL")
    rows = conn.execute("""
        SELECT interface, status, COUNT(*) as cnt, AVG(elapsed_ms) as avg_ms
        FROM screening_result GROUP BY interface, status
    """).fetchall()
    conn.close()
    result: Dict = {}
    for iface, status, cnt, avg_ms in rows:
        result.setdefault(iface, {})[status] = {
            "count": cnt, "avg_ms": round(avg_ms or 0)
        }
    return result


# ── 文件模式存储 ──────────────────────────────────────────────────────────────

def files_dest(output_dir: Path, doc_path: str, interface: str) -> Path:
    p = Path(doc_path)
    return output_dir / p.parent / (p.stem + f".{interface}.json")

def files_exists(output_dir: Path, doc_path: str, interface: str) -> bool:
    return files_dest(output_dir, doc_path, interface).exists()

def files_write(output_dir: Path, r: Dict):
    dest = files_dest(output_dir, r["doc_path"], r["interface"])
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_suffix(".tmp")
    tmp.write_text(json.dumps(r, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(dest)


# ── HTTP 采集（单任务，含重试）───────────────────────────────────────────────

async def fetch_one(
    client:    "httpx.AsyncClient",
    iname:     str,
    icfg:      Dict,
    doc_path:  str,
    doc_body:  Dict,
) -> Dict:
    url       = icfg["url"]
    method    = icfg.get("method", "POST").upper()
    timeout   = icfg.get("timeout", DEFAULT_TIMEOUT)
    retry_cfg = icfg.get("retry") or DEFAULT_RETRY
    max_retry = int(retry_cfg.get("max", 3))
    backoff_b = float(retry_cfg.get("backoff_base", 1.0))
    extra     = icfg.get("payload_extra") or {}

    # 认证请求头
    headers: Dict[str, str] = {}
    auth_cfg  = icfg.get("auth") or {}
    auth_type = auth_cfg.get("type", "none")
    if auth_type == "bearer":
        token = os.environ.get(auth_cfg.get("token_env", ""), "")
        if token:
            headers["Authorization"] = f"Bearer {token}"
    elif auth_type == "header":
        headers.update(auth_cfg.get("headers") or {})

    payload   = {**doc_body, **extra}
    attempts  = 0
    last_err  = ""
    last_code: Optional[int] = None
    t0        = time.time()

    while attempts <= max_retry:
        try:
            if method == "POST":
                resp = await client.post(url, json=payload, headers=headers, timeout=timeout)
            else:
                resp = await client.get(url, headers=headers, timeout=timeout)
            attempts += 1
            last_code = resp.status_code

            if resp.status_code < 300:
                return {
                    "doc_path": doc_path, "interface": iname, "status": "success",
                    "raw_response": resp.text, "error": None,
                    "http_status": last_code, "attempts": attempts,
                    "elapsed_ms": int((time.time() - t0) * 1000),
                    "fetched_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
                }
            elif resp.status_code == 429 or resp.status_code >= 500:
                last_err = f"HTTP {resp.status_code}: {resp.text[:200]}"
                if attempts <= max_retry:
                    jitter = 0.8 + 0.4 * (hash(doc_path + iname + str(attempts)) % 100 / 100)
                    await asyncio.sleep(backoff_b * (2 ** (attempts - 1)) * jitter)
            else:
                return {
                    "doc_path": doc_path, "interface": iname, "status": "failed",
                    "raw_response": None,
                    "error": f"HTTP {resp.status_code}: {resp.text[:500]}",
                    "http_status": last_code, "attempts": attempts,
                    "elapsed_ms": int((time.time() - t0) * 1000),
                    "fetched_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
                }
        except (httpx.ConnectError, httpx.TimeoutException, httpx.ReadError) as e:
            attempts += 1
            last_err = f"{type(e).__name__}: {e}"
            if attempts <= max_retry:
                await asyncio.sleep(backoff_b * (2 ** (attempts - 1)))
        except Exception as e:
            attempts += 1
            last_err = str(e)
            break

    return {
        "doc_path": doc_path, "interface": iname, "status": "failed",
        "raw_response": None, "error": last_err, "http_status": last_code,
        "attempts": attempts, "elapsed_ms": int((time.time() - t0) * 1000),
        "fetched_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
    }


# ── 每个接口的 Worker 池 ─────────────────────────────────────────────────────

async def run_interface(
    iname:       str,
    doc_rels:    List[str],
    icfg:        Dict,
    input_dir:   Path,
    result_q:    asyncio.Queue,
    stop_event:  asyncio.Event,
):
    workers_count = icfg.get("workers", DEFAULT_WORKERS)
    q: asyncio.Queue = asyncio.Queue(maxsize=workers_count * 4)

    async def feeder():
        for rel in doc_rels:
            if stop_event.is_set():
                break
            await q.put(rel)
        for _ in range(workers_count):
            await q.put(None)  # sentinel

    async def worker():
        async with httpx.AsyncClient() as client:
            while True:
                rel = await q.get()
                if rel is None:
                    q.task_done()
                    return
                try:
                    if stop_event.is_set():
                        q.task_done()
                        continue
                    try:
                        body = json.loads((input_dir / rel).read_bytes())
                    except Exception as e:
                        await result_q.put({
                            "doc_path": rel, "interface": iname, "status": "failed",
                            "raw_response": None, "error": f"读取失败: {e}",
                            "http_status": None, "attempts": 0, "elapsed_ms": 0,
                            "fetched_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
                        })
                        q.task_done()
                        continue
                    result = await fetch_one(client, iname, icfg, rel, body)
                    await result_q.put(result)
                finally:
                    q.task_done()

    feeder_t = asyncio.create_task(feeder())
    workers  = [asyncio.create_task(worker()) for _ in range(workers_count)]
    await asyncio.gather(feeder_t, *workers)


# ── 核心运行流程 ──────────────────────────────────────────────────────────────

async def run_fetcher(
    input_dir:    Path,
    output_dir:   Path,
    config:       Dict,
    interfaces:   List[str],
    storage_mode: str,
    force:        bool,
    stop_event:   asyncio.Event,
):
    output_dir.mkdir(parents=True, exist_ok=True)
    iface_cfgs = {k: config["interfaces"][k] for k in interfaces}

    # 扫描输入文件
    json_files = sorted([
        p for p in input_dir.rglob("*.json")
        if not p.name.startswith(".") and not p.name.startswith("~$") and p.is_file()
    ])
    if not json_files:
        print("  未找到 JSON 文件")
        return

    # 初始化存储
    db_conn = None
    done_set: Set[Tuple[str, str]] = set()
    if storage_mode == "sqlite":
        db_path = output_dir / DB_FILENAME
        db_conn = open_db(db_path)
        if not force:
            done_set = load_done_set(db_conn)
            print(f"  SQLite 已有记录: {len(done_set)} 条（将跳过）")

    # 构造各接口任务列表
    tasks_by_iface: Dict[str, List[str]] = {i: [] for i in interfaces}
    skip_count = 0
    for jf in json_files:
        rel = str(jf.relative_to(input_dir))
        for iname in interfaces:
            if not force:
                if storage_mode == "sqlite" and (rel, iname) in done_set:
                    skip_count += 1
                    continue
                if storage_mode == "files" and files_exists(output_dir, rel, iname):
                    skip_count += 1
                    continue
            tasks_by_iface[iname].append(rel)

    total_pending = sum(len(v) for v in tasks_by_iface.values())
    total_all     = len(json_files) * len(interfaces)
    print(f"  文件: {len(json_files):,}  接口: {len(interfaces)}")
    print(f"  总任务: {total_all:,}  待执行: {total_pending:,}  跳过: {skip_count:,}")
    if total_pending == 0:
        print("  全部已完成！")
        if db_conn:
            db_conn.close()
        return

    # 日志
    log_path = output_dir / LOG_FILENAME
    setup_logger(log_path)
    log(f"===== 开始采集 total={total_pending} interfaces={','.join(interfaces)} =====")
    print()

    # 共享结果队列
    result_q: asyncio.Queue = asyncio.Queue(maxsize=256)
    stats_path  = output_dir / STATS_FILENAME
    failed_path = output_dir / FAILED_FILENAME

    cnt_success = {i: 0 for i in interfaces}
    cnt_failed  = {i: 0 for i in interfaces}
    cnt_done    = 0
    t_start     = time.time()

    # tqdm 进度条
    pbar = tqdm(total=total_pending, unit="篇", dynamic_ncols=True) if HAS_TQDM else None

    # Writer 协程（唯一写入者，串行安全）
    async def writer():
        nonlocal cnt_done
        stats_counter = 0
        while True:
            item = await result_q.get()
            if item is None:
                result_q.task_done()
                break
            try:
                # 落盘
                if storage_mode == "sqlite" and db_conn:
                    db_write(db_conn, item)
                else:
                    files_write(output_dir, item)

                # 统计
                iname = item["interface"]
                if item["status"] == "success":
                    cnt_success[iname] = cnt_success.get(iname, 0) + 1
                    log(f"OK  {iname}: {item['doc_path']} ({item['elapsed_ms']}ms x{item['attempts']})")
                else:
                    cnt_failed[iname] = cnt_failed.get(iname, 0) + 1
                    with open(failed_path, "a", encoding="utf-8") as f:
                        f.write(f"{item['doc_path']}\t{iname}\t{item.get('http_status','')}\t{item.get('error','')}\n")
                    log(f"ERR {iname}: {item['doc_path']} | {str(item.get('error',''))[:100]}", "warning")

                cnt_done += 1
                stats_counter += 1
                if stats_counter >= 100:
                    stats_counter = 0
                    _flush_stats(stats_path, interfaces, cnt_success, cnt_failed, t_start, total_pending)

                # 进度显示
                if pbar:
                    s = "  ".join(f"{i}:✓{cnt_success.get(i,0)}/✗{cnt_failed.get(i,0)}" for i in interfaces)
                    pbar.set_postfix_str(s)
                    pbar.update(1)
                elif cnt_done % 10 == 0 or cnt_done == total_pending:
                    elapsed = time.time() - t_start
                    spd = cnt_done / elapsed if elapsed > 0.1 else 0
                    eta = int((total_pending - cnt_done) / spd) if spd > 0 else 0
                    s   = "  ".join(f"{i}:✓{cnt_success.get(i,0)}/✗{cnt_failed.get(i,0)}" for i in interfaces)
                    print(f"\r  [{cnt_done}/{total_pending}] {spd:.1f}篇/s 剩余~{eta}s  {s}   ",
                          end="", flush=True)
            finally:
                result_q.task_done()

    # 启动 writer 和各接口的 producer
    writer_task = asyncio.create_task(writer())
    producer_tasks = [
        asyncio.create_task(
            run_interface(i, tasks_by_iface[i], iface_cfgs[i], input_dir, result_q, stop_event)
        )
        for i in interfaces
    ]

    try:
        await asyncio.gather(*producer_tasks)
    except (asyncio.CancelledError, KeyboardInterrupt):
        pass

    await result_q.put(None)  # 通知 writer 结束
    await writer_task

    if pbar:
        pbar.close()
    else:
        print()

    if db_conn:
        db_conn.close()

    _flush_stats(stats_path, interfaces, cnt_success, cnt_failed, t_start, total_pending)
    log("===== 采集完成 =====")

    # 结果摘要
    elapsed = time.time() - t_start
    print()
    print("═" * 46)
    print("  采集完成")
    print("═" * 46)
    for i in interfaces:
        s = cnt_success.get(i, 0)
        f = cnt_failed.get(i, 0)
        print(f"  {i}: 成功 {s}  失败 {f}")
    spd = cnt_done / elapsed if elapsed > 0.1 else 0
    print(f"  总耗时: {elapsed:.1f}s  速度: {spd:.1f}篇/s")
    if storage_mode == "sqlite":
        print(f"  数据库: {output_dir / DB_FILENAME}")
    if (output_dir / FAILED_FILENAME).exists():
        print(f"  失败列表: {output_dir / FAILED_FILENAME}")
    print()


def _flush_stats(stats_path, interfaces, cnt_success, cnt_failed, t_start, total):
    elapsed = time.time() - t_start
    done    = sum(cnt_success.get(i, 0) + cnt_failed.get(i, 0) for i in interfaces)
    atomic_write_json(stats_path, {
        "total": total, "done": done,
        "elapsed_s": round(elapsed, 1),
        "speed": round(done / elapsed, 2) if elapsed > 0.1 else 0,
        "interfaces": {
            i: {"success": cnt_success.get(i, 0), "failed": cnt_failed.get(i, 0)}
            for i in interfaces
        },
        "updated_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
    })


# ── 交互向导 ──────────────────────────────────────────────────────────────────

def print_welcome():
    print()
    print("╔══════════════════════════════════════════════════════════╗")
    print("║         筛查接口数据采集器  v" + VERSION + "                        ║")
    print("║  把阶段二 JSON 批量发给 N 个 HTTP 接口，存入 SQLite       ║")
    print("╚══════════════════════════════════════════════════════════╝")
    print()

def read_path(prompt: str, must_exist: bool = False, allow_create: bool = False) -> str:
    while True:
        print(f"\n  {prompt}")
        print("  提示：可直接把文件夹拖入窗口")
        raw = input("  路径: ").strip().strip('"').strip("'")
        if not raw:
            continue
        if must_exist and not os.path.exists(raw):
            print(f"  [错误] 路径不存在: {raw}")
            continue
        if allow_create and not os.path.exists(raw):
            ans = input("  目录不存在，是否创建？[Y/n] ").strip().lower()
            if ans in ("", "y"):
                Path(raw).mkdir(parents=True, exist_ok=True)
                return raw
            continue
        return raw

def wizard_gen_config(config_path: Path):
    """首次启动：引导用户生成 screening_config.yaml"""
    print()
    print("══════════════════════════════════════════")
    print("  未找到配置文件，进入接口配置向导")
    print("══════════════════════════════════════════")
    print()
    print("  为每个筛查接口填写连接信息。")
    print("  URL 建议使用 ${ENV_VAR} 占位，真实值在环境变量中设置，不进代码库。")
    print()

    ifaces: Dict = {}
    while True:
        name = input("  接口名（留空结束）: ").strip()
        if not name:
            if not ifaces:
                print("  至少需要配置一个接口。")
                continue
            break

        url = input(f"  [{name}] URL（建议如 ${{SCREENING_{name}_URL}}）: ").strip()
        if not url:
            url = f"${{{f'SCREENING_{name}_URL'}}}"

        auth_type = input(f"  [{name}] 认证方式 [bearer/header/none，默认 bearer]: ").strip().lower() or "bearer"
        auth_cfg: Dict = {}
        if auth_type == "bearer":
            env = input(f"  [{name}] Token 环境变量名（默认 SCREENING_{name}_TOKEN）: ").strip()
            auth_cfg = {"type": "bearer", "token_env": env or f"SCREENING_{name}_TOKEN"}
        elif auth_type == "header":
            auth_cfg = {"type": "header", "headers": {}}
            while True:
                hk = input(f"  [{name}] 请求头名（留空结束）: ").strip()
                if not hk:
                    break
                hv = input(f"  [{name}] {hk} 的值: ").strip()
                auth_cfg["headers"][hk] = hv
        else:
            auth_cfg = {"type": "none"}

        wk = input(f"  [{name}] 并发数（默认 {DEFAULT_WORKERS}）: ").strip()
        to = input(f"  [{name}] 超时秒数（默认 {DEFAULT_TIMEOUT}）: ").strip()

        ifaces[name] = {
            "url": url, "method": "POST", "auth": auth_cfg,
            "workers": int(wk) if wk.isdigit() else DEFAULT_WORKERS,
            "timeout": int(to) if to.isdigit() else DEFAULT_TIMEOUT,
            "retry": {"max": 3, "backoff_base": 1.0},
            "enabled": True, "payload_extra": {},
        }
        print(f"  ✓ 已添加接口 {name}\n")

    cfg_yaml = yaml.dump({"interfaces": ifaces}, allow_unicode=True,
                          sort_keys=False, default_flow_style=False)
    header = (
        f"# 筛查接口配置  (生成于 {time.strftime('%Y-%m-%d %H:%M:%S')})\n"
        "# 敏感信息（URL/Token）建议使用环境变量：\n"
        "#   export SCREENING_A_URL=https://...\n"
        "#   export SCREENING_A_TOKEN=your_token\n"
        "# 增加接口：在 interfaces 下追加一节，无需改代码。\n\n"
    )
    config_path.write_text(header + cfg_yaml, encoding="utf-8")
    print(f"  ✓ 配置已写入: {config_path}")
    print()
    print("  下一步：")
    print("  1. 设置环境变量（或直接在 yaml 里填真实 URL）")
    print("  2. 再次运行: python run_screening.py")
    print()

def run_wizard(config: Dict, last_run: Optional[Dict]) -> Optional[Dict]:
    """交互式收集运行参数，返回 dict 或 None（取消）"""
    avail = [k for k, v in config.get("interfaces", {}).items() if v.get("enabled", True)]
    if not avail:
        print("  [错误] 没有可用接口（所有接口均已禁用）")
        return None

    # 1/4 输入目录
    print("\n 1/4  输入目录（阶段二的 JSON 输出）")
    if last_run:
        print(f"      上次: {last_run.get('input_dir', '')}")
    input_dir = read_path("输入目录（拖拽 / 手动输入）:", must_exist=True)
    n_json = sum(1 for _ in Path(input_dir).rglob("*.json"))
    print(f"      ✓ 找到 {n_json:,} 篇 JSON")

    # 2/4 输出目录
    print("\n 2/4  输出目录")
    default_out = last_run.get("output_dir") if last_run else str(Path(input_dir).parent / "screen_out")
    raw_out = input(f"      路径（直接回车 → {default_out}）: ").strip().strip('"').strip("'")
    out_dir = raw_out or default_out
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    print(f"      ✓ 输出目录: {out_dir}")

    # 3/4 选接口
    print(f"\n 3/4  选择接口（共 {len(avail)} 个可用）")
    for name in avail:
        c = config["interfaces"][name]
        print(f"      • {name}  workers={c.get('workers', DEFAULT_WORKERS)}  timeout={c.get('timeout', DEFAULT_TIMEOUT)}s")
    last_ifaces = last_run.get("interfaces", []) if last_run else []
    sel_raw = input("      要运行哪些接口？（留空=全部，或输入名称如 A,B）: ").strip()
    if sel_raw:
        sel = [s.strip() for s in sel_raw.split(",") if s.strip() in avail]
        if not sel:
            print("      未匹配，使用全部接口")
            sel = avail
    else:
        sel = avail
    print(f"      ✓ 已选: {', '.join(sel)}")

    # 4/4 预估 + 运行模式
    total = n_json * len(sel)
    print(f"\n 4/4  预估")
    print(f"      待调用: {n_json:,} × {len(sel)} = {total:,} 次（已完成的会自动跳过）")
    print()
    print("      [Y] 前台运行（实时查看进度）")
    print("      [B] 后台运行（SSH 断开也继续）")
    print("      [N] 取消")
    while True:
        choice = input("      选择 [Y/B/N]: ").strip().upper() or "Y"
        if choice in ("Y", "B", "N"):
            break
    if choice == "N":
        print("  已取消")
        return None

    return {
        "input_dir":  input_dir,
        "output_dir": out_dir,
        "interfaces": sel,
        "storage":    "sqlite",
        "detach":     (choice == "B"),
    }


# ── 守护进程 ──────────────────────────────────────────────────────────────────

def daemonize_posix(log_path: Path, pid_path: Path):
    """POSIX double-fork，父进程打印提示后退出，孙进程继续"""
    pid = os.fork()
    if pid > 0:
        sys.exit(0)
    os.setsid()
    pid = os.fork()
    if pid > 0:
        sys.exit(0)
    # 孙进程：重定向 IO
    sys.stdout.flush()
    sys.stderr.flush()
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with open("/dev/null", "r") as fn:
        os.dup2(fn.fileno(), sys.stdin.fileno())
    with open(log_path, "a") as fl:
        os.dup2(fl.fileno(), sys.stdout.fileno())
        os.dup2(fl.fileno(), sys.stderr.fileno())
    pid_path.write_text(str(os.getpid()), encoding="utf-8")

def daemonize_windows(log_path: Path, pid_path: Path) -> int:
    """Windows：重启自己（去掉 --detach），以分离进程运行"""
    args = [sys.executable] + [a for a in sys.argv if a != "--detach"]
    log_path.parent.mkdir(parents=True, exist_ok=True)
    DETACHED = 0x00000008
    NEW_GRP  = 0x00000200
    with open(log_path, "a") as lf:
        p = subprocess.Popen(
            args, stdout=lf, stderr=lf, stdin=subprocess.DEVNULL,
            creationflags=DETACHED | NEW_GRP, close_fds=True,
        )
    pid_path.write_text(str(p.pid), encoding="utf-8")
    return p.pid


# ── 子命令实现 ────────────────────────────────────────────────────────────────

def cmd_run(args):
    if not HAS_HTTPX:
        die("缺少 httpx，请先：pip install httpx")
    if not HAS_YAML:
        die("缺少 PyYAML，请先：pip install PyYAML")

    config_path = resolve_config_path()
    if not config_path.exists():
        print_welcome()
        wizard_gen_config(Path(CONFIG_FILE))
        return

    config = load_config(config_path)
    raw_yaml = config_path.read_text(encoding="utf-8")
    missing  = find_missing_env(raw_yaml)
    if missing:
        print()
        print("  [警告] 以下环境变量未设置（对应接口可能无法调用）：")
        for v in missing:
            print(f"    export {v}=<值>")

    # 读取命令行参数
    input_dir_s  = getattr(args, "input",     None)
    output_dir_s = getattr(args, "output",    None)
    detach       = getattr(args, "detach",    False)
    force        = getattr(args, "force",     False)
    storage      = getattr(args, "storage",   "sqlite")
    iface_sel    = getattr(args, "interface", None)
    replay       = getattr(args, "replay",    False)
    workers_ov   = getattr(args, "workers",   None)

    # --replay
    last_run: Optional[Dict] = None
    lr_file = Path(output_dir_s or ".") / LAST_RUN_FILE
    if not lr_file.parent.exists():
        lr_file = Path(".") / LAST_RUN_FILE
    if replay and lr_file.exists() and HAS_YAML:
        last_run = yaml.safe_load(lr_file.read_text(encoding="utf-8")) or {}
        input_dir_s  = input_dir_s  or last_run.get("input_dir")
        output_dir_s = output_dir_s or last_run.get("output_dir")
        if not iface_sel:
            iface_sel = ",".join(last_run.get("interfaces", []))
        storage = last_run.get("storage", storage)

    # 是否需要向导
    need_wizard = not input_dir_s or not output_dir_s
    if need_wizard:
        print_welcome()
        params = run_wizard(config, last_run)
        if not params:
            return
        input_dir_s  = params["input_dir"]
        output_dir_s = params["output_dir"]
        detach       = params["detach"]
        storage      = params.get("storage", "sqlite")
        iface_sel    = ",".join(params["interfaces"])

    input_dir  = Path(input_dir_s).resolve()
    output_dir = Path(output_dir_s).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    # 解析接口列表
    all_enabled = [k for k, v in config.get("interfaces", {}).items() if v.get("enabled", True)]
    if iface_sel:
        interfaces = [i.strip() for i in iface_sel.split(",") if i.strip() in all_enabled]
        if not interfaces:
            die(f"没有匹配的可用接口。可用: {', '.join(all_enabled)}")
    else:
        interfaces = all_enabled
    if not interfaces:
        die("没有启用的接口，请检查配置文件。")

    # worker 数覆写
    if workers_ov and workers_ov > 0:
        for k in interfaces:
            config["interfaces"][k]["workers"] = workers_ov

    # 保存 .last_run.yaml
    if HAS_YAML:
        (output_dir / LAST_RUN_FILE).write_text(
            yaml.dump({
                "input_dir": str(input_dir), "output_dir": str(output_dir),
                "interfaces": interfaces, "storage": storage,
            }, allow_unicode=True),
            encoding="utf-8",
        )

    # 后台化
    if detach:
        log_path = output_dir / LOG_FILENAME
        pid_path = output_dir / PID_FILENAME
        if sys.platform == "win32":
            pid = daemonize_windows(log_path, pid_path)
            print(f"\n  ✓ 已后台运行  PID={pid}  日志: {log_path}")
        else:
            print(f"\n  后台运行中...  日志: {log_path}")
            print(f"  查看进度: python run_screening.py status -o {output_dir}")
            print(f"  实时日志: python run_screening.py logs -o {output_dir} -f")
            print(f"  停止任务: python run_screening.py stop -o {output_dir}")
            print()
            sys.stdout.flush()
            daemonize_posix(log_path, pid_path)
            # 以下在孙进程中运行（父进程已退出）

    # 注册优雅退出信号
    stop_event = asyncio.Event()

    def _on_stop(sig, frame):
        print("\n  收到停止信号，完成当前批次后退出...")
        stop_event.set()

    signal.signal(signal.SIGTERM, _on_stop)
    signal.signal(signal.SIGINT,  _on_stop)

    # 运行
    asyncio.run(run_fetcher(
        input_dir=input_dir,
        output_dir=output_dir,
        config=config,
        interfaces=interfaces,
        storage_mode=storage,
        force=force,
        stop_event=stop_event,
    ))

    # 清理 PID 文件
    (output_dir / PID_FILENAME).unlink(missing_ok=True)

    # 交互模式完成后询问打开目录
    if need_wizard and not detach and sys.stdout.isatty():
        ans = input("  是否打开输出目录？[Y/n] ").strip().lower()
        if ans in ("", "y"):
            try:
                if sys.platform == "win32":
                    os.startfile(str(output_dir))
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", str(output_dir)])
                else:
                    subprocess.Popen(["xdg-open", str(output_dir)])
            except Exception:
                pass


def cmd_status(args):
    out = Path(args.output).resolve()
    pid_path   = out / PID_FILENAME
    stats_path = out / STATS_FILENAME
    db_path    = out / DB_FILENAME

    print()
    print("  筛查采集器 — 当前状态")
    print("  " + "─" * 38)

    # 进程状态
    pid, running = None, False
    if pid_path.exists():
        try:
            pid = int(pid_path.read_text().strip())
            os.kill(pid, 0)
            running = True
        except (ValueError, ProcessLookupError, PermissionError):
            pass

    if running:
        print(f"  状态: 运行中  PID={pid}")
    elif pid:
        print(f"  状态: 已停止 (上次 PID={pid})")
    else:
        print("  状态: 无后台进程")

    # 进度（优先读 _stats.json，速度快）
    if stats_path.exists():
        try:
            d = json.loads(stats_path.read_text(encoding="utf-8"))
            total, done = d.get("total", 0), d.get("done", 0)
            pct = done / total * 100 if total else 0
            spd = d.get("speed", 0)
            print(f"  进度: {done:,} / {total:,}  ({pct:.1f}%)  速度 {spd:.1f}篇/s  耗时 {d.get('elapsed_s',0):.0f}s")
            print()
            print("  接口明细:")
            for iname, v in d.get("interfaces", {}).items():
                print(f"    {iname}:  ✓{v.get('success',0)}  ✗{v.get('failed',0)}")
        except Exception:
            pass
    elif db_path.exists():
        stats = query_stats_from_db(db_path)
        print()
        print("  接口明细（来自 SQLite）:")
        for iname, d in stats.items():
            s = d.get("success", {}).get("count", 0)
            f = d.get("failed",  {}).get("count", 0)
            a = d.get("success", {}).get("avg_ms", 0)
            print(f"    {iname}:  ✓{s}  ✗{f}  均{a}ms")
    else:
        print("  暂无数据")

    if running:
        print()
        print(f"  实时日志: python run_screening.py logs -o {args.output} -f")
        print(f"  停止任务: python run_screening.py stop -o {args.output}")
    print()


def cmd_logs(args):
    log_path = Path(args.output).resolve() / LOG_FILENAME
    if not log_path.exists():
        print(f"  日志文件不存在: {log_path}")
        return
    follow = getattr(args, "follow", False)
    if follow:
        print(f"  实时日志 {log_path}  (Ctrl+C 退出)\n")
        try:
            with open(log_path, encoding="utf-8", errors="replace") as f:
                f.seek(0, 2)
                while True:
                    line = f.readline()
                    if line:
                        print(line, end="")
                    else:
                        time.sleep(0.3)
        except KeyboardInterrupt:
            print()
    else:
        lines = log_path.read_text(encoding="utf-8", errors="replace").splitlines()
        for line in lines[-50:]:
            print(line)


def cmd_stop(args):
    out      = Path(args.output).resolve()
    pid_file = out / PID_FILENAME
    force    = getattr(args, "force", False)
    if not pid_file.exists():
        print("  未找到 PID 文件，没有正在运行的后台任务。")
        return
    try:
        pid = int(pid_file.read_text().strip())
    except ValueError:
        print("  PID 文件内容无效。")
        return
    sig = signal.SIGKILL if (force and hasattr(signal, "SIGKILL")) else signal.SIGTERM
    try:
        os.kill(pid, sig)
        if force:
            print(f"  已强制结束 PID={pid}")
        else:
            print(f"  已发送停止信号 PID={pid}（进程将在写完当前批次后退出，断点已保存）")
    except ProcessLookupError:
        print(f"  进程 {pid} 已不存在。")
        pid_file.unlink(missing_ok=True)
    except PermissionError:
        print(f"  权限不足，无法停止进程 {pid}。")


def cmd_doctor(_args):
    print()
    print("  筛查采集器 — 系统自检")
    print("  " + "─" * 38)
    ok = True

    # Python
    pv = sys.version_info
    if pv >= (3, 8):
        print(f"  ✓ Python {pv.major}.{pv.minor}.{pv.micro}")
    else:
        print(f"  ✗ Python {pv.major}.{pv.minor}（需要 >= 3.8）")
        ok = False

    for dep, has, opt, install in [
        ("httpx",  HAS_HTTPX, False, "httpx"),
        ("PyYAML", HAS_YAML,  False, "PyYAML"),
        ("tqdm",   HAS_TQDM,  True,  "tqdm"),
    ]:
        if has:
            print(f"  ✓ {dep}")
        elif opt:
            print(f"  ○ {dep} 未安装（可选，缺失不影响主流程）")
        else:
            print(f"  ✗ {dep} 未安装  →  pip install {install}")
            ok = False

    # 配置文件
    cp = resolve_config_path()
    if cp.exists():
        print(f"  ✓ {cp.name} 存在")
        try:
            raw_yaml = cp.read_text(encoding="utf-8")
            missing  = find_missing_env(raw_yaml)
            cfg      = load_config(cp)
            ifaces   = cfg.get("interfaces", {})
            if ifaces:
                print(f"  ✓ 配置了 {len(ifaces)} 个接口: {', '.join(ifaces.keys())}")
            else:
                print("  ✗ interfaces 为空")
                ok = False
            for v in missing:
                print(f"  ✗ 环境变量 ${{{v}}} 未设置")
                ok = False
            # 连通性测试
            if ifaces and HAS_HTTPX:
                print()
                print("  接口连通性（HEAD 测试，超时 5s）:")
                import httpx as _httpx
                for name, icfg in ifaces.items():
                    url = icfg.get("url", "")
                    if "${" in url:
                        print(f"    ○ {name}: URL 含未展开占位符，跳过")
                        continue
                    try:
                        r = _httpx.head(url, timeout=5)
                        print(f"    ✓ {name}: HTTP {r.status_code}")
                    except _httpx.ConnectError:
                        print(f"    ✗ {name}: 无法连接  可能原因: 服务器未启动 / 防火墙 / VPN")
                        ok = False
                    except _httpx.TimeoutException:
                        print(f"    ✗ {name}: 连接超时")
                        ok = False
                    except Exception as e:
                        print(f"    ✗ {name}: {e}")
                        ok = False
        except Exception as e:
            print(f"  ✗ 配置文件解析失败: {e}")
            ok = False
    else:
        print(f"  ✗ {CONFIG_FILE} / {CONFIG_FILE_LOCAL} 均不存在  →  运行 python run_screening.py 生成")
        ok = False

    print()
    print("  一切就绪！" if ok else "  发现问题，请根据上方提示处理后再运行。")
    print()


def cmd_dump(args):
    out        = Path(args.output).resolve()
    db_path    = out / DB_FILENAME
    export_dir = Path(args.export_dir).resolve() if args.export_dir else out / "exported"

    if not db_path.exists():
        print(f"  数据库不存在: {db_path}")
        return

    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row

    conds, params = [], []
    if args.doc:
        conds.append("doc_path = ?");       params.append(args.doc)
    if getattr(args, "iface_filter", None):
        conds.append("interface = ?");      params.append(args.iface_filter)
    if getattr(args, "status_filter", None):
        conds.append("status = ?");         params.append(args.status_filter)

    where = ("WHERE " + " AND ".join(conds)) if conds else ""
    rows  = conn.execute(f"SELECT * FROM screening_result {where}", params).fetchall()
    conn.close()

    if not rows:
        print("  没有匹配的记录。")
        return

    export_dir.mkdir(parents=True, exist_ok=True)
    count = 0
    for row in rows:
        doc_p = row["doc_path"]
        iface = row["interface"]
        dest  = export_dir / Path(doc_p).parent / (Path(doc_p).stem + f".{iface}.json")
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_text(json.dumps(dict(row), ensure_ascii=False, indent=2), encoding="utf-8")
        count += 1

    print(f"  导出 {count} 条到: {export_dir}")


# ── 主入口 ────────────────────────────────────────────────────────────────────

def main():
    # 不带任何参数 → 进入 run 向导
    if len(sys.argv) == 1:
        class _Stub:
            input = output = interface = None
            detach = force = replay = False
            storage = "sqlite"
            workers = None
        cmd_run(_Stub())
        return

    p = argparse.ArgumentParser(
        prog="run_screening.py",
        description="筛查接口数据采集器 — 阶段三",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = p.add_subparsers(dest="cmd")

    # run
    pr = sub.add_parser("run", help="采集（核心命令）")
    pr.add_argument("-i", "--input",     help="阶段二 JSON 输入目录")
    pr.add_argument("-o", "--output",    help="输出目录（默认 ./screen_out）")
    pr.add_argument("--interface",       help="只跑指定接口（逗号分隔）")
    pr.add_argument("--workers", type=int, help="覆写所有接口的并发数")
    pr.add_argument("--storage", choices=["sqlite", "files"], default="sqlite")
    pr.add_argument("--force",   action="store_true", help="强制重跑已有结果")
    pr.add_argument("--detach",  action="store_true", help="后台运行（SSH 断开也继续）")
    pr.add_argument("--replay",  action="store_true", help="复用上次的交互配置")

    # status
    ps = sub.add_parser("status", help="查询当前进度")
    ps.add_argument("-o", "--output", default="./screen_out")

    # logs
    pl = sub.add_parser("logs", help="查看日志")
    pl.add_argument("-o", "--output", default="./screen_out")
    pl.add_argument("-f", "--follow", action="store_true", help="实时跟踪")

    # stop
    pst = sub.add_parser("stop", help="停止后台任务")
    pst.add_argument("-o", "--output", default="./screen_out")
    pst.add_argument("--force", action="store_true", help="强制结束（SIGKILL）")

    # doctor
    sub.add_parser("doctor", help="系统自检")

    # dump
    pd = sub.add_parser("dump", help="从 SQLite 导出原始结果为文件")
    pd.add_argument("-o", "--output",       default="./screen_out", help="采集输出目录")
    pd.add_argument("--export-dir",         dest="export_dir",      help="导出目标目录")
    pd.add_argument("--doc",                                         help="只导出某篇（相对路径）")
    pd.add_argument("--all",  action="store_true",                   help="导出全部")
    pd.add_argument("--interface",          dest="iface_filter",     help="只导出某接口")
    pd.add_argument("--status",             dest="status_filter",    help="只导出某状态")

    args = p.parse_args()
    {
        "run":    cmd_run,
        "status": cmd_status,
        "logs":   cmd_logs,
        "stop":   cmd_stop,
        "doctor": cmd_doctor,
        "dump":   cmd_dump,
    }.get(args.cmd, lambda _: p.print_help())(args)


if __name__ == "__main__":
    main()
