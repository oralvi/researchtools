#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
from pathlib import Path

from DCPDexport.src.core import write_output






def run_interactive():
    """Interactive loop borrowed from the old 260109new script."""
    while True:
        try:
            file_path = input("请输入数据文件路径（例如 D:\\data\\test.txt）：").strip()
            if not file_path:
                print("[warn] 路径不能为空，请重新输入。\n")
                continue

            src = Path(file_path).expanduser().resolve()
            if not src.exists() or not src.is_file():
                print(f"[warn] 文件不存在：{src}\n")
                continue

            print("\n请选择第一列的时间单位：")
            print("1. sec  - 秒")
            print("2. hr   - 小时")
            choice = input("请输入选择（1 或 2，默认为 1）：").strip() or "1"
            unit = "hr" if choice == "2" else "sec"

            try:
                write_output(src, unit, None, ask_overwrite=True)
            except Exception as e:
                print(f"[warn] 处理失败：{e}\n")

            again = input("[info] 是否继续处理其他文件？(y/n，默认为n)：").strip().lower()
            if again != 'y':
                print("[info] 程序已退出。")
                break
            print()
        except KeyboardInterrupt:
            print("\n\n[info] 程序已退出。")
            break
        except Exception as e:
            print(f"[warn] 未知错误：{e}\n")
            continue


def main():
    parser = argparse.ArgumentParser(description="DCPD 数据导出（Console）")
    parser.add_argument("input", nargs="?", help="原始数据文件路径")
    parser.add_argument("--unit", choices=["sec", "hr"], default="sec", help="输出时间单位")
    parser.add_argument("--output", help="输出 CSV 路径（可选）")
    parser.add_argument("--interactive", action="store_true", help="交互式模式")
    parser.add_argument("--overwrite", action="store_true", help="输出文件已存在时询问覆盖")
    args = parser.parse_args()

    if args.interactive or not args.input:
        run_interactive()
        return

    source = Path(args.input).expanduser().resolve()
    if not source.exists() or not source.is_file():
        raise SystemExit(f"[warn] 输入文件不存在: {source}")
    output = Path(args.output).expanduser().resolve() if args.output else None
    # use shared write_output function
    output_file, enc, sorted_seconds = write_output(source, args.unit, output, ask_overwrite=args.overwrite)
    print(f"[ok] 输入文件: {source}")
    print(f"[ok] 输入编码: {enc}")
    print(f"[ok] 输出文件: {output_file}")
    print(f"[ok] 时间段数量: {len(sorted_seconds)}")
    print(f"[ok] 时间范围: {sorted_seconds[0]} ~ {sorted_seconds[-1]} sec")


if __name__ == "__main__":
    main()
