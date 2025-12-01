"""
    后台清理守护进程
    启动方式: 与 run_web.bat 同时启动，独立后台运行。
    功能: 每隔 INTERVAL_SECONDS 秒扫描 excel_input 与 word_output，删除超过 RETENTION_HOURS 未访问的临时文件。
    注意: 仅删除基于时间戳命名的文件。
"""

import os
import re
import time
from typing import Tuple
from pathlib import Path
from datetime import datetime, timedelta
from logger import get_logger

logger = get_logger("cleanup_loop")

BASE_DIR = Path(__file__).parent.resolve()
INPUT_DIR = BASE_DIR / 'excel_input'
OUTPUT_DIR = BASE_DIR / 'word_output'
RETENTION_HOURS = 1           # 保留小时
INTERVAL_SECONDS = 1800          # 清理间隔: 30分钟
TIMESTAMP_PATTERN = re.compile(r".+_(\d{13})\..+")  # 仅匹配末尾含13位毫秒时间戳的文件名

INPUT_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


def is_timestamp_file(path: Path) -> bool:
    return TIMESTAMP_PATTERN.match(path.name) is not None


def cleanup_dir(directory: Path, cutoff: datetime) -> Tuple[int, int]:
    deleted_count = 0
    freed_bytes = 0
    for file_path in directory.iterdir():
        if not file_path.is_file():
            continue
        if not is_timestamp_file(file_path):
            continue  # 跳过非临时命名文件
        try:
            # 使用修改时间(mtime)而非访问时间(atime)，避免因扫描/查看导致无法清理
            last_modified = datetime.fromtimestamp(file_path.stat().st_mtime)
            if last_modified < cutoff:
                size = file_path.stat().st_size
                file_path.unlink()
                deleted_count += 1
                freed_bytes += size
        except Exception:
            # 忽略单文件异常
            pass
    return deleted_count, freed_bytes


def format_size(bytes_value: int) -> str:
    if bytes_value < 1024:
        return f"{bytes_value} B"
    if bytes_value < 1024 * 1024:
        return f"{bytes_value/1024:.2f} KB"
    return f"{bytes_value/1024/1024:.2f} MB"


def run_loop():
    logger.info("[清理守护] 启动成功，间隔: %ds，保留: %dh" % (INTERVAL_SECONDS, RETENTION_HOURS))
    while True:
        start_ts = time.time()
        now = datetime.now()
        cutoff = now - timedelta(hours=RETENTION_HOURS)
        in_del, in_bytes = cleanup_dir(INPUT_DIR, cutoff)
        out_del, out_bytes = cleanup_dir(OUTPUT_DIR, cutoff)
        total_del = in_del + out_del
        total_bytes = in_bytes + out_bytes
        if total_del > 0:
            logger.info(f"[清理守护] {now:%Y-%m-%d %H:%M:%S} 删除 {total_del} 个文件, 释放 {format_size(total_bytes)}")
        else:
            logger.info(f"[清理守护] {now:%Y-%m-%d %H:%M:%S} 无需清理")
        # 睡眠剩余时间
        elapsed = time.time() - start_ts
        sleep_sec = max(5, INTERVAL_SECONDS - elapsed)
        time.sleep(sleep_sec)


if __name__ == '__main__':
    try:
        run_loop()
    except KeyboardInterrupt:
        logger.info("[清理守护] 已退出")
