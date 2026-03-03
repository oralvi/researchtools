"""Shared processing logic for DCPD export tools.

This module contains functions previously defined in the console script and is
imported by both the console and GUI front ends. Keeping the logic here avoids
code duplication and simplifies maintenance.
"""
from collections import defaultdict
from datetime import datetime
from pathlib import Path
import csv
import argparse
import chardet


def detect_encodings(path: Path):
    sample = path.read_bytes()[:100000]
    guess = chardet.detect(sample)
    enc = guess.get("encoding")
    conf = guess.get("confidence", 0.0) or 0.0
    # report the detected encoding to the caller
    return enc, conf


def unique_path(base: Path, ask_user: bool = False) -> Path:
    if not base.exists():
        return base
    if ask_user:
        response = input(f"\n[warn] 文件已存在：{base.name}\n是否替换？(y/n，默认为n)：").strip().lower()
        if response == "y":
            return base
    stem, suffix = base.stem, base.suffix
    i = 1
    while True:
        p = base.with_name(f"{stem}-{i}{suffix}")
        if not p.exists():
            if ask_user:
                print(f"[ok] 文件将保存为：{p.name}")
            return p
        i += 1


def parse_source(path: Path):
    data_by_second = defaultdict(list)
    first_ts = None
    encodings = []
    enc, conf = detect_encodings(path)
    if enc and conf >= 0.7:
        encodings = [enc, "gbk", "gb2312", "utf-8", "utf-8-sig"]
    else:
        encodings = ["gbk", "gb2312", "gb18030", "utf-8", "utf-8-sig", "latin-1"]
    for enc in encodings:
        try:
            with path.open("r", encoding=enc) as f:
                for line in f:
                    line = line.strip()
                    if not line or "二次平均" in line:
                        continue
                    parts = line.split(",")
                    if len(parts) < 5:
                        continue
                    try:
                        c1 = float(parts[0])
                        c2 = float(parts[1])
                        c3 = float(parts[2])
                        c4 = float(parts[3])
                        t = datetime.strptime(parts[4], "%y%m%d%H%M%S")
                    except ValueError:
                        continue
                    if first_ts is None:
                        first_ts = t
                    sec = int((t - first_ts).total_seconds())
                    data_by_second[sec].append((c1, c2, c3, c4))
            return data_by_second, enc
        except UnicodeDecodeError:
            continue
    raise RuntimeError(f"无法解码文件: {path}")


def write_output(source: Path, unit: str, output: Path | None, ask_overwrite: bool = False):
    data_by_second, enc = parse_source(source)
    if not data_by_second:
        raise RuntimeError("未提取到有效数据")
    if output is None:
        ts = datetime.now().strftime("%y%m%d%H%M")
        output = source.with_name(f"{source.stem}_parsed_{unit}_{ts}.csv")
    output = unique_path(output, ask_user=ask_overwrite)
    sorted_seconds = sorted(data_by_second.keys())
    with output.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([unit, "V", "Vr", "a", "a/w"])
        for second in sorted_seconds:
            points = data_by_second[second]
            avg = [sum(p[i] for p in points) / len(points) for i in range(4)]
            t_val = second / 3600.0 if unit == "hr" else float(second)
            writer.writerow([f"{t_val:.6f}", f"{avg[0]:.6f}", f"{avg[1]:.6f}", f"{avg[2]:.6f}", f"{avg[3]:.6f}"])
    return output, enc, sorted_seconds


def main():
    parser = argparse.ArgumentParser(description="DCPD 数据导出工具")
    parser.add_argument("source", help="输入源文件路径")
    parser.add_argument("--unit", choices=["sec", "hr"], default="sec", help="时间单位")
    parser.add_argument("-o", "--output", help="输出 CSV 路径")
    parser.add_argument("--overwrite", action="store_true", help="若输出文件存在则覆盖")
    args = parser.parse_args()

    source = Path(args.source)
    if not source.exists():
        raise FileNotFoundError(f"输入文件不存在: {source}")

    output = Path(args.output) if args.output else None
    out_path, enc, seconds = write_output(
        source=source,
        unit=args.unit,
        output=output,
        ask_overwrite=not args.overwrite,
    )
    print(f"[ok] 编码: {enc}")
    print(f"[ok] 记录秒数: {len(seconds)}")
    print(f"[ok] 输出文件: {out_path}")


if __name__ == "__main__":
    main()
