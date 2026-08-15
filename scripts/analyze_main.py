# -*- coding: utf-8 -*-
"""
扫描主文件，识别缺失/EARLY_CUTOFF/LOW_COUNT 日期，写 analysis_result.json

用法：
  python analyze_main.py --workspace C:\\path\\to\\workspace
  python analyze_main.py --workspace .
  python analyze_main.py --city 杭州市 --workspace C:\\path\\to\\workspace   # 指定地市（默认丽水市）
"""
import openpyxl
import json
import sys
import argparse
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path


def get_args():
    parser = argparse.ArgumentParser(
        description='分析主文件数据完整性',
        epilog='Example: python analyze_main.py --workspace C:\\path\\to\\workspace'
    )
    parser.add_argument('--workspace', type=str, default='.',
                        help='工作目录（主文件所在），默认当前目录 .')
    parser.add_argument('--city', type=str, default='丽水市',
                        help='目标地市（默认丽水市），主文件名按地市命名')
    return parser.parse_args()


def main():
    args = get_args()
    workspace = Path(args.workspace).resolve()
    city = args.city
    city_base = city[:-1] if city.endswith('市') else city
    main_file = workspace / f'{city_base}政府采购公告.xlsx'

    if not workspace.exists():
        sys.exit(f'错误：工作目录不存在：{workspace}')
    if not main_file.exists():
        print(f'主文件不存在：{main_file}')
        print('将采集最近 3 天数据（跳过本步分析，直接进入采集）')
        # 主文件缺失时同样落盘 analysis_result.json（最近 3 天），与 fetch_zfcg 自动模式
        # 内部默认口径一致，保证下游（AI 拆批预估等）始终有统一的分析产物可读
        today = datetime.now()
        output = {
            "incomplete_dates": [(today - timedelta(days=i)).strftime('%Y-%m-%d') for i in range(2, -1, -1)],
            "total_records": 0,
            "today": today.strftime('%Y-%m-%d'),
            "current_hour": today.hour
        }
        output_path = workspace / 'analysis_result.json'
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output, f, ensure_ascii=False)
        print(f"Analysis saved to {output_path}")
        return

    wb = openpyxl.load_workbook(main_file)
    ws = wb.active

    print(f"Sheet: {ws.title}")

    date_stats = defaultdict(lambda: {'count': 0, 'latest_hour': 0, 'latest_time': ''})
    total = 0

    for i, row in enumerate(ws.iter_rows(min_row=1, values_only=True)):
        if i < 3:
            row_str = [str(c)[:30] if c else '' for c in row[:6]]
            print(f"  Row {i+1}: {row_str}")
            continue
        if not row or not row[0]:
            continue
        total += 1
        dt_val = row[0]
        if isinstance(dt_val, datetime):
            date_str = dt_val.strftime('%Y-%m-%d')
            hour = dt_val.hour
            time_str = dt_val.strftime('%H:%M')
        else:
            try:
                dt_str = str(dt_val)
                if len(dt_str) >= 10:
                    date_str = dt_str[:10]
                    hour = int(dt_str[11:13]) if len(dt_str) > 12 else 0
                    time_str = dt_str[11:16] if len(dt_str) > 12 else '00:00'
                else:
                    continue
            except:
                continue
        date_stats[date_str]['count'] += 1
        if hour > date_stats[date_str]['latest_hour']:
            date_stats[date_str]['latest_hour'] = hour
            date_stats[date_str]['latest_time'] = time_str

    wb.close()

    print(f"\nTotal records: {total}")

    MIN_RECORDS = 30
    MIN_HOUR = 17

    incomplete_dates = []
    today = datetime.now()
    current_hour = today.hour

    print(f"\n{'Date':<12} {'Count':>6} {'Latest':>8} {'Status':>12}")
    print("-" * 42)

    for i in range(10):
        check_date = (today - timedelta(days=i)).strftime('%Y-%m-%d')
        stats = date_stats.get(check_date)
        count = stats['count'] if stats else 0
        latest = stats['latest_time'] if stats and stats['latest_time'] else '--:--'
        latest_hour = stats['latest_hour'] if stats else 0

        if count == 0:
            status = "MISSING"
            incomplete_dates.append(check_date)
        elif stats['latest_hour'] < MIN_HOUR:
            status = "EARLY_CUTOFF"
            incomplete_dates.append(check_date)
        elif count < MIN_RECORDS:
            status = "LOW_COUNT"
            incomplete_dates.append(check_date)
        elif check_date == today.strftime('%Y-%m-%d') and current_hour < 18:
            status = "TODAY_PENDING"
        else:
            status = "OK"

        print(f"{check_date:<12} {count:>6} {latest:>8} {status:>12}")

    output_path = workspace / 'analysis_result.json'
    if incomplete_dates:
        print(f"\nNeeds collection: {', '.join(sorted(incomplete_dates))}")
        output = {
            "incomplete_dates": sorted(incomplete_dates),
            "total_records": total,
            "today": today.strftime('%Y-%m-%d'),
            "current_hour": current_hour
        }
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output, f, ensure_ascii=False)
        print(f"Analysis saved to {output_path}")
    else:
        print("\nAll recent data is complete!")
        yesterday = (today - timedelta(days=1)).strftime('%Y-%m-%d')
        fetch = [yesterday]
        if current_hour >= 18:
            fetch.append(today.strftime('%Y-%m-%d'))
        print(f"Default collection: {', '.join(fetch)}")
        output = {
            "incomplete_dates": fetch,
            "total_records": total,
            "today": today.strftime('%Y-%m-%d'),
            "current_hour": current_hour
        }
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output, f, ensure_ascii=False)


if __name__ == '__main__':
    main()
