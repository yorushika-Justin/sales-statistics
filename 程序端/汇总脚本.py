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
import threading
from datetime import datetime, timedelta, timezone

# R-5（复核）：容器时区是 UTC，日志时间戳统一按北京时间，避免同一个日志文件里
# 混着相差 8 小时的两种时间。
CST = timezone(timedelta(hours=8))

# 将项目根目录加入 sys.path，以便 import 原版汇总脚本
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 服务器上没有 Tk 库，mock 掉 tkinter 避免原版导入时报错
if os.name != 'nt':
    import types
    _tk_mock = types.ModuleType('tkinter')
    _tk_mock.Tk = type('Tk', (), {'withdraw': lambda s: None, 'destroy': lambda s: None})
    _tk_mock.messagebox = types.ModuleType('tkinter.messagebox')
    _tk_mock.messagebox.showerror = lambda *a, **kw: None
    _tk_mock.messagebox.askyesno = lambda *a, **kw: False
    sys.modules['tkinter'] = _tk_mock
    sys.modules['tkinter.messagebox'] = _tk_mock.messagebox

# 导入原版核心函数
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

# FONT_PATH: 用仓库里的微软雅黑字体（程序端/msyh.ttc）
_FONT_DIR = os.path.dirname(os.path.abspath(__file__))
_FONT_IN_REPO = os.path.join(_FONT_DIR, 'msyh.ttc')

if os.name == 'nt':
    # Windows: 优先系统字体
    _FONT_PATH_DEFAULT = r'C:\Windows\Fonts\msyh.ttc'
else:
    # Linux (Render): 用仓库里的字体
    _FONT_PATH_DEFAULT = _FONT_IN_REPO

FONT_PATH = os.environ.get('FONT_PATH', _FONT_PATH_DEFAULT)

# P-8：字体路径运行期不会变 —— 模块加载时一次性写入原版模块即可。
# 原实现在 process_xls() 里「临时覆盖 + finally 恢复」模块级 FONT_PATH，
# 在 Flask threaded=True 下线程不安全（并发请求会读到对方改过的值）。
import 汇总脚本 as _orig_module  # noqa: E402  （与原版为同一模块对象）

_orig_module.FONT_PATH = FONT_PATH

# P-15d：处理过程全局串行化。
# 同一日期的品牌图 / 品名图文件名固定，两个请求同时处理同一天会并发写同一个
# PNG（Pillow 边画边写，存在交错写坏文件的真实竞态）。单用户场景下串行没有
# 体感代价，却能换来"并发绝不出错"，因此在此加锁。
_PROCESS_LOCK = threading.Lock()


def show_error_server(msg, log_path=None):
    """服务器版错误处理：写日志 + print，不弹 tkinter 窗口"""
    print(f"[错误] {msg}", file=sys.stderr)
    if log_path:
        try:
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(f'[{datetime.now(CST)}] {msg}\n')
                if sys.exc_info()[0] is not None:
                    import traceback
                    f.write(f'{traceback.format_exc()}\n')
        except Exception:
            pass


def process_xls(xls_path, log_path=None):
    """
    服务器端处理入口：xls/xlsx → 品牌汇总 PNG + 品名明细 PNG

    P-15d：全局串行化，避免并发写同名 PNG 互相踩踏。
    返回: (brand_png_path, item_png_path, date_str, total)
    """
    with _PROCESS_LOCK:
        return _process_locked(xls_path, log_path)


def _process_locked(xls_path, log_path=None):
    import traceback

    is_xlsx = xls_path.lower().endswith('.xlsx')

    # 读取数据
    data, date_str = read_xlsx(xls_path, log_path=log_path) if is_xlsx else read_xls(xls_path, log_path=log_path)

    # 创建输出目录
    date_dir = os.path.join(OUTPUT_DIR, date_str)
    os.makedirs(date_dir, exist_ok=True)

    # 品牌汇总 PNG（字体已在模块加载时写入原版模块，此处不再改写全局变量）
    brand_png = os.path.join(date_dir, f"{date_str}.png")
    try:
        draw_table(data, date_str, brand_png)
    except Exception:
        # R-7（复核）：绘制中途失败会留下不完整 PNG，清理掉，
        # 否则 /history 会把它当成正常结果列出来。
        _remove_quietly(brand_png)
        raise

    # 品名明细 PNG
    item_png = None
    item_png_path = os.path.join(date_dir, f"{date_str}_品名.png")
    try:
        items, _ = read_xlsx_items(xls_path, log_path=log_path) if is_xlsx else read_xls_items(xls_path, log_path=log_path)
        draw_table_detail(items, date_str, item_png_path)
        item_png = item_png_path
    except Exception as detail_err:
        # P-14：失败时必须保持 item_png 为 None。
        # 原实现先赋值路径再画图，异常被吞掉后仍把该路径返回给调用方，
        # 上游据此返回 200 + 一个不存在的图片地址（前端显示破图）。
        # R-7：同时清理可能残留的不完整 PNG。
        _remove_quietly(item_png_path)
        item_png = None
        if log_path:
            try:
                with open(log_path, 'a', encoding='utf-8') as f:
                    f.write(f'[{datetime.now(CST)}] 品名明细生成失败: {detail_err}\n')
                    f.write(f'{traceback.format_exc()}\n')
            except Exception:
                pass

    total = round(sum(data.values()), 2)
    return brand_png, item_png, date_str, total


def _remove_quietly(path):
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except Exception:
        pass
