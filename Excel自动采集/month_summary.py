# -*- coding: utf-8 -*-
"""月度采集出图入口（图形界面版专用，v2.0）

调度链：图形界面版 server.js → collect.js --month → 本脚本

用法:
    python month_summary.py --mode summary --xls <导出文件> --month 2026-09
    python month_summary.py --mode brand   --xls <导出文件> --month 2026-09 --brand 自由点

可选参数:
    --title <标题>   覆盖自动标题（默认 "2026年9月总销量" / "2026年9月自由点总销量"）
    --outdir <目录>  覆盖输出目录（默认 <汇总脚本.OUTPUT_DIR>/<YYYY-MM>）
    --out <文件名>   覆盖输出文件名

输出:
    <输出目录>/2026-09_总销量.png        （mode=summary）
    <输出目录>/2026-09_自由点明细.png     （mode=brand）

约定（与 collect.js 对接）:
    成功 → stdout 打印一行 "PNG 月度: <绝对路径> (<字节> 字节)"，退出码 0
    失败 → stderr 打印原因，退出码 2
"""
import argparse
import importlib.util
import os
import re
import sys
import traceback

try:  # Windows 下管道默认 GBK，collect.js 按 UTF-8 解码
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def load_core():
    """按文件路径加载 汇总脚本.py（工作区根），避免受 sys.path 顺序影响"""
    spec = importlib.util.spec_from_file_location("huizong_core", os.path.join(ROOT, "汇总脚本.py"))
    mod = importlib.util.module_from_spec(spec)
    sys.modules["huizong_core"] = mod
    spec.loader.exec_module(mod)
    return mod


def month_title(month: str, suffix: str) -> str:
    """'2026-09' + '总销量' → '2026年9月总销量'"""
    y, m = month.split('-')
    return f"{int(y)}年{int(m)}月{suffix}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--mode', required=True, choices=['summary', 'brand'])
    ap.add_argument('--xls', required=True)
    ap.add_argument('--month', required=True, help='YYYY-MM')
    ap.add_argument('--brand', default='')
    ap.add_argument('--title', default='')
    ap.add_argument('--outdir', default='')
    ap.add_argument('--out', default='')
    args = ap.parse_args()

    if not re.match(r'^\d{4}-\d{2}$', args.month):
        print(f"月份格式非法: {args.month}（要求 YYYY-MM）", file=sys.stderr)
        return 2
    if not os.path.exists(args.xls):
        print(f"文件不存在: {args.xls}", file=sys.stderr)
        return 2

    core = load_core()
    rows, meta = core.read_month_rows(args.xls, None)
    print(f"读取 {len(rows)} 行 | 日期行: {meta['date_line']} | 门店行: {meta['store_line']} | 品牌行: {meta['brand_line']}")
    if not rows:
        print("导出文件里没有任何销售记录", file=sys.stderr)
        return 2

    outdir = args.outdir or os.path.join(core.OUTPUT_DIR, args.month)
    os.makedirs(outdir, exist_ok=True)

    if args.mode == 'summary':
        data = core.sum_brand_from_rows(rows)
        if not data:
            print("没有可汇总的品牌数据", file=sys.stderr)
            return 2
        title = args.title or month_title(args.month, '总销量')
        out = os.path.join(outdir, args.out or f"{args.month}_总销量.png")
        core.draw_table_month(data, title, out)
        print(f"品牌汇总: {len(data)} 个品牌, 总金额 {sum(data.values()):.2f}")
    else:
        brand = args.brand.strip()
        if not brand:
            print("brand 模式必须提供 --brand", file=sys.stderr)
            return 2
        items = core.sum_items_from_rows(rows, brand)
        if not items:
            print(f"没有找到品牌「{brand}」的数据", file=sys.stderr)
            return 2
        title = args.title or month_title(args.month, f'{brand}总销量')
        out = os.path.join(outdir, args.out or f"{args.month}_{brand}明细.png")
        core.draw_table_detail_month(items, title, out)
        total_qty = sum(x[4] for x in items)
        total_amount = sum(x[6] for x in items)
        print(f"品名明细: {len(items)} 行, 数量合计 {total_qty:g}, 金额合计 {total_amount:.2f}")

    size = os.path.getsize(out)
    print(f"PNG 月度: {out} ({size} 字节)")
    return 0


if __name__ == '__main__':
    try:
        _code = main()
    except Exception as _e:
        traceback.print_exc()
        print(f"ERROR: {type(_e).__name__}: {_e}", file=sys.stderr)
        _code = 2
    sys.exit(_code)
