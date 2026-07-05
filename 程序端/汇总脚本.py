"""
Render 部署用 — 汇总脚本适配层
================================
从原版 汇总脚本.py 导入核心函数，覆盖服务器不兼容的配置：
- OUTPUT_DIR: 用环境变量或 Linux 默认路径
- FONT_PATH: 按系统自动选字体
- tkinter: 服务器无 Tk 库，延迟导入
- os.startfile: 服务器无此 API，跳过
"""
import sys
import os

# 将项目根目录加入 sys.path，以便 import 原版汇总脚本
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 尝试导入 tkinter（服务器上可能没有 Tk 库）
try:
    import tkinter as tk
    _HAS_TK = True
except ImportError:
    _HAS_TK = False

# 导入原版核心函数（这些函数不依赖 tkinter）
from 汇总脚本 import (
    read_xls,
    read_xlsx,
    read_xls_items,
    read_xlsx_items,
    draw_table,
    draw_table_detail,
    extract_date_from_filename,
    _extract_date_from_xls,
    _extract_date_from_xlsx,
)

# ==================== 服务器配置覆盖 ====================

# OUTPUT_DIR: 优先环境变量，fallback 用 Linux 默认路径
OUTPUT_DIR = os.environ.get(
    'OUTPUT_DIR',
    os.path.join(os.path.expanduser('~'), 'output', 'sales')
)

# FONT_PATH: 按系统自动选字体
if os.name == 'nt':
    # Windows
    _FONT_PATH_DEFAULT = r'C:\Windows\Fonts\msyh.ttc'
else:
    # Linux (Render)
    _FONT_CANDIDATES = [
        '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
        '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
        '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    ]
    _FONT_PATH_DEFAULT = None
    for fp in _FONT_CANDIDATES:
        if os.path.exists(fp):
            _FONT_PATH_DEFAULT = fp
            break
    if _FONT_PATH_DEFAULT is None:
        _FONT_PATH_DEFAULT = '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc'

FONT_PATH = os.environ.get('FONT_PATH', _FONT_PATH_DEFAULT)


def show_error_server(msg, log_path=None):
    """服务器版错误处理：写日志 + print，不弹 tkinter 窗口"""
    print(f"[错误] {msg}", file=sys.stderr)
    if log_path:
        try:
            with open(log_path, 'a', encoding='utf-8') as f:
                from datetime import datetime
                f.write(f'[{datetime.now()}] {msg}\n')
                if sys.exc_info()[0] is not None:
                    import traceback
                    f.write(f'{traceback.format_exc()}\n')
        except Exception:
            pass


def process_xls(xls_path, log_path=None):
    """
    服务器端处理入口：xls/xlsx → 品牌汇总 PNG + 品名明细 PNG
    返回: (brand_png_path, item_png_path, date_str, total)
    """
    from datetime import datetime
    import traceback

    is_xlsx = xls_path.lower().endswith('.xlsx')

    # 读取数据
    data, date_str = read_xlsx(xls_path, log_path=log_path) if is_xlsx else read_xls(xls_path, log_path=log_path)

    # 创建输出目录
    date_dir = os.path.join(OUTPUT_DIR, date_str)
    os.makedirs(date_dir, exist_ok=True)

    # 品牌汇总 PNG
    brand_png = os.path.join(date_dir, f"{date_str}.png")

    # 临时覆盖 FONT_PATH（draw_table 内部用模块级 FONT_PATH）
    import 汇总脚本 as _orig
    _orig_font_backup = _orig.FONT_PATH
    _orig.FONT_PATH = FONT_PATH
    try:
        draw_table(data, date_str, brand_png)
    finally:
        _orig.FONT_PATH = _orig_font_backup

    # 品名明细 PNG
    item_png = None
    try:
        items, _ = read_xlsx_items(xls_path, log_path=log_path) if is_xlsx else read_xls_items(xls_path, log_path=log_path)
        item_png = os.path.join(date_dir, f"{date_str}_品名.png")
        _orig.FONT_PATH = FONT_PATH
        try:
            draw_table_detail(items, date_str, item_png)
        finally:
            _orig.FONT_PATH = _orig_font_backup
    except Exception as detail_err:
        if log_path:
            try:
                with open(log_path, 'a', encoding='utf-8') as f:
                    f.write(f'[{datetime.now()}] 品名明细生成失败: {detail_err}\n')
                    f.write(f'{traceback.format_exc()}\n')
            except Exception:
                pass

    total = round(sum(data.values()), 2)
    return brand_png, item_png, date_str, total
