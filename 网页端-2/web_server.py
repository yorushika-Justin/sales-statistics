"""
销售汇总网页端 v2.0
===================

用途:
    本地 Web 服务，浏览器拖入 xls/xlsx 文件 → 生成 PNG 表格图片 → 系统默认图片查看器打开

用法:
    python web_server.py
    python web_server.py --silent  # 静默模式（开机自启用）

启动后自动打开浏览器访问 http://127.0.0.1:5000
在网页中拖入 xls/xlsx 文件即可处理

输入:
    浏览器拖入 .xls / .xlsx 文件

输出:
    <OUTPUT_DIR>/<日期>/<日期>.png        (品牌汇总，自动打开)
    <OUTPUT_DIR>/<日期>/<日期>_品名.png   (品名明细，不自动打开)

配置:
    输出路径从 config.json 读取，修改 config.json 即可切换路径

依赖:
    - Flask
    - xlrd 1.2.0 (读 .xls)
    - openpyxl (读 .xlsx)
    - Pillow
    - tkinter (Python 内置，用于 os.startfile)

技术架构:
    - Flask 本地服务，端口 5000
    - 前端纯 HTML/JS，内嵌在 Python 文件中
    - 后端复用汇总脚本.py 核心函数
"""
import sys
import os
import json
import uuid
import socket
import threading
import tempfile
import traceback
from datetime import datetime

from flask import Flask, request, jsonify

# 读取配置文件
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')
try:
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        config = json.load(f)
    OUTPUT_DIR = config.get('output_dir', r'D:\16678\相册\销售额')
    FONT_PATH = config.get('font_path', r'C:\Windows\Fonts\msyh.ttc')
    PORT = config.get('port', 5000)
except Exception as e:
    print(f'[警告] 读取配置文件失败: {e}，使用默认值')
    OUTPUT_DIR = r'D:\16678\相册\销售额'
    FONT_PATH = r'C:\Windows\Fonts\msyh.ttc'
    PORT = 5000

# 将上级目录加入 sys.path，以便 import 汇总脚本
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from 汇总脚本 import read_xls, read_xlsx, draw_table, read_xls_items, read_xlsx_items, draw_table_detail

# 上传文件大小限制（10MB，足够 xls 文件）
MAX_UPLOAD_SIZE = 10 * 1024 * 1024

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = MAX_UPLOAD_SIZE

HTML_PAGE = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>销售汇总工具</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: "Microsoft YaHei", "微软雅黑", sans-serif;
            background: #f5f7fa;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
        }
        .container {
            background: white;
            border-radius: 12px;
            box-shadow: 0 2px 12px rgba(0, 0, 0, 0.08);
            padding: 40px;
            width: 500px;
            text-align: center;
        }
        h1 {
            color: #333;
            font-size: 24px;
            margin-bottom: 8px;
        }
        .subtitle {
            color: #888;
            font-size: 13px;
            margin-bottom: 30px;
        }
        .drop-zone {
            border: 2px dashed #c0c0c0;
            border-radius: 8px;
            padding: 50px 20px;
            cursor: pointer;
            transition: all 0.3s;
            background: #fafafa;
        }
        .drop-zone:hover, .drop-zone.drag-over {
            border-color: #4a90d9;
            background: #f0f7ff;
        }
        .drop-zone-text {
            color: #888;
            font-size: 16px;
        }
        .drop-zone-hint {
            color: #bbb;
            font-size: 12px;
            margin-top: 10px;
        }
        .status {
            margin-top: 20px;
            padding: 12px;
            border-radius: 6px;
            font-size: 14px;
            display: none;
        }
        .status.success {
            background: #f0f9ff;
            color: #1a7f37;
            border: 1px solid #d4edda;
            display: block;
        }
        .status.error {
            background: #fff5f5;
            color: #dc3545;
            border: 1px solid #f5c6cb;
            display: block;
        }
        .status.loading {
            background: #fff3cd;
            color: #856404;
            border: 1px solid #ffc107;
            display: block;
        }
        input[type="file"] { display: none; }
    </style>
</head>
<body>
    <div class="container">
        <h1>销售汇总工具</h1>
        <div class="subtitle">拖入 .xls / .xlsx 文件，自动生成品牌汇总 + 品名明细图片</div>
        <div class="drop-zone" id="dropZone">
            <div class="drop-zone-text">将 .xls / .xlsx 文件拖入此处</div>
            <div class="drop-zone-hint">或点击选择文件</div>
        </div>
        <div class="status" id="status"></div>
        <input type="file" id="fileInput" accept=".xls,.xlsx">
        <div style="text-align: center; margin-top: 20px; color: #999; font-size: 12px;">
            <button id="stopBtn" style="background:none;border:none;color:#999;cursor:pointer;text-decoration:underline;">
                停止服务
            </button>
            <span style="margin-left: 8px;">|</span>
            <span>如需结束服务，可点上方按钮或在任务管理器结束 pythonw.exe</span>
        </div>
    </div>

    <script>
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const status = document.getElementById('status');

        // 拖拽事件
        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('drag-over');
        });

        dropZone.addEventListener('dragleave', () => {
            dropZone.classList.remove('drag-over');
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('drag-over');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                uploadFile(files[0]);
            }
        });

        // 点击选择
        dropZone.addEventListener('click', () => {
            fileInput.click();
        });

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                uploadFile(e.target.files[0]);
            }
        });

        function showStatus(message, type) {
            status.textContent = message;
            status.className = 'status ' + type;
        }

        function uploadFile(file) {
            if (!file.name.toLowerCase().endsWith('.xls') && !file.name.toLowerCase().endsWith('.xlsx')) {
                showStatus('仅支持 .xls / .xlsx 文件', 'error');
                return;
            }

            showStatus('处理中……', 'loading');

            const formData = new FormData();
            formData.append('file', file);

            fetch('/upload', {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    showStatus(data.message, 'success');
                } else {
                    showStatus(data.message, 'error');
                }
            })
            .catch(error => {
                showStatus('请求失败: ' + error.message, 'error');
            });
        }

        // 停止服务按钮
        document.getElementById('stopBtn').addEventListener('click', () => {
            if (confirm('确认停止服务？停止后需重新双击启动网页端.vbs')) {
                fetch('/shutdown', {method: 'POST'}).then(() => {
                    document.body.innerHTML = '<h1 style="text-align:center;margin-top:200px;color:#888;">服务已停止</h1>';
                });
            }
        });
    </script>
</body>
</html>
"""


@app.route('/')
def index():
    return HTML_PAGE


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify(success=False, message='未收到文件')

    file = request.files['file']
    if not file.filename:
        return jsonify(success=False, message='未选择文件')

    if not (file.filename.lower().endswith('.xls') or file.filename.lower().endswith('.xlsx')):
        return jsonify(success=False, message='仅支持 .xls / .xlsx 文件')

    # 检查输出目录
    if not os.path.isdir(OUTPUT_DIR):
        return jsonify(success=False, message=f'输出目录不存在: {OUTPUT_DIR}，请先创建该目录')

    # 保存到临时文件（使用唯一文件名避免多标签页冲突）
    is_xlsx = file.filename.lower().endswith('.xlsx')
    ext = '.xlsx' if is_xlsx else '.xls'
    temp_filename = f'sales_upload_{uuid.uuid4().hex[:8]}{ext}'
    temp_path = os.path.join(tempfile.gettempdir(), temp_filename)
    file.save(temp_path)

    try:
        # 调用现有核心函数
        log_path = os.path.join(tempfile.gettempdir(), 'sales_web_error.log')
        data, date_str = read_xlsx(temp_path, log_path=log_path) if is_xlsx else read_xls(temp_path, log_path=log_path)
        # 按日期创建子文件夹
        date_dir = os.path.join(OUTPUT_DIR, date_str)
        os.makedirs(date_dir, exist_ok=True)
        brand_png = os.path.join(date_dir, f"{date_str}.png")
        draw_table(data, date_str, brand_png)

        # 品名明细 PNG（失败不阻塞品牌汇总，但写日志 + 前端提示）
        item_png = None
        try:
            items, _ = read_xlsx_items(temp_path, log_path=log_path) if is_xlsx else read_xls_items(temp_path, log_path=log_path)
            item_png = os.path.join(date_dir, f"{date_str}_品名.png")
            draw_table_detail(items, date_str, item_png)
        except Exception as detail_err:
            try:
                with open(log_path, 'a', encoding='utf-8') as f:
                    f.write(f'[{datetime.now()}] 品名明细生成失败: {detail_err}\n')
                    f.write(f'{traceback.format_exc()}\n')
            except Exception:
                pass

        # 用系统默认图片查看器打开品牌汇总（异步，避免阻塞请求）
        threading.Thread(target=os.startfile, args=(brand_png,), daemon=True).start()

        # 计算总金额（round 防止浮点尾数误差）
        total = round(sum(data.values()), 2)

        msg = f'已生成: {date_str}/{date_str}.png（总金额: {total:.2f} 元）'
        if item_png:
            msg += f' + 品名明细'
        else:
            msg += f'（品名明细生成失败，详见 {log_path}）'

        return jsonify(
            success=True,
            message=msg,
            brand_png=f'{date_str}/{date_str}.png',
            item_png=f'{date_str}/{date_str}_品名.png' if item_png else None,
            total=total
        )
    except Exception as e:
        # 持久化 traceback 到日志文件，便于排查
        log_path = os.path.join(tempfile.gettempdir(), 'sales_web_error.log')
        try:
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(f'[{datetime.now()}] {traceback.format_exc()}\n')
        except Exception:
            pass
        return jsonify(success=False, message=f'{type(e).__name__}: {str(e)}')
    finally:
        # 清理临时文件
        try:
            os.remove(temp_path)
        except Exception:
            pass


@app.route('/shutdown', methods=['POST'])
def shutdown():
    """停止 Flask 服务（用于网页底部"停止服务"按钮）。

    注意：
        - 不做二次确认（前端 JS 已做 confirm 弹窗）
        - 用 os._exit(0) 立即终止进程，不走 Flask atexit 清理
        - 不需要鉴权：服务只监听 127.0.0.1
    """
    os._exit(0)


def open_browser():
    """等待服务启动后打开浏览器。

    用 socket 轮询探测端口就绪，替代 time.sleep(1) 硬等待。
    二次启动（端口已占）由 _check_already_running() 处理，本函数只在首次启动时调用。
    """
    # 轮询探测端口是否已 listening（最多等 5 秒）
    for _ in range(50):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(0.1)
                s.connect(('127.0.0.1', PORT))
                break  # 端口已就绪
        except (ConnectionRefusedError, socket.timeout, OSError):
            pass
    os.startfile(f'http://127.0.0.1:{PORT}')


def _show_popup(msg, kind='error'):
    """统一的 tkinter 弹窗工具函数。

    Args:
        msg: 弹窗消息。
        kind: 'error' = 错误弹窗，'yesno' = 是/否询问弹窗。

    Returns:
        bool: kind='yesno' 时返回用户选择（True=是，False=否）；kind='error' 返回 None。
    """
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    try:
        if kind == 'yesno':
            return messagebox.askyesno('销售汇总网页端', msg)
        else:
            messagebox.showerror('销售汇总网页端 - 错误', msg)
            return None
    finally:
        root.destroy()


def _check_already_running():
    """检测端口是否已被占用。

    Returns:
        bool: True = 已被占用（服务已在运行），False = 未占用。

    行为：
        - 已占用 → 弹 tkinter 弹窗询问"是否打开浏览器？"
                   用户选 是 → os.startfile(...) + sys.exit(0)
                   用户选 否 → sys.exit(0)
        - 未占用 → 返回 False，让调用者继续启动
    """
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(0.3)
    try:
        sock.connect(('127.0.0.1', PORT))
        # 端口已占用
        answer = _show_popup(
            f'服务已在运行。\n\n是否打开浏览器访问 http://127.0.0.1:{PORT} ？',
            kind='yesno'
        )
        if answer:
            os.startfile(f'http://127.0.0.1:{PORT}')
        sys.exit(0)
    except (ConnectionRefusedError, socket.timeout, OSError):
        return False
    finally:
        sock.close()


def _check_output_dir():
    """检测输出目录是否存在。

    Returns:
        bool: True = 存在，False = 不存在（已弹窗报错 + sys.exit(1)）。
    """
    if os.path.isdir(OUTPUT_DIR):
        return True

    # 目录不存在：弹窗报错
    _show_popup(f'输出目录不存在：\n{OUTPUT_DIR}\n\n请先创建该目录后再启动服务。')
    sys.exit(1)


def _show_error_popup(msg):
    """静默场景下的错误弹窗（替代 print）。

    Args:
        msg: 错误信息字符串。
    """
    try:
        _show_popup(msg)
    except Exception as e:
        # 兜底：弹窗失败时写日志
        log_path = os.path.join(tempfile.gettempdir(), 'sales_web_error.log')
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write(f'[弹窗失败] {msg}\n[原因] {e}\n')


if __name__ == '__main__':
    # 检查是否为静默模式（开机自启时使用 --silent 参数）
    silent_mode = '--silent' in sys.argv

    # 启动检查（静默场景下必须用弹窗，不能用 print）
    if not silent_mode:
        _check_already_running()  # 端口已被占 → 弹窗询问 + sys.exit(0)
    else:
        # 静默模式：端口已占则静默退出，不弹窗
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(0.3)
        try:
            sock.connect(('127.0.0.1', PORT))
            sock.close()
            sys.exit(0)  # 服务已在运行，静默退出
        except (ConnectionRefusedError, socket.timeout, OSError):
            pass
        finally:
            sock.close()

    _check_output_dir()       # 输出目录不存在 → 弹窗报错 + sys.exit(1)

    # 在新线程中打开浏览器（非静默模式才打开）
    if not silent_mode:
        threading.Thread(target=open_browser, daemon=True).start()

    # 启动 Flask 服务（threaded=True 支持多标签页并发）
    try:
        app.run(host='127.0.0.1', port=PORT, debug=False, threaded=True)
    except OSError as e:
        if 'address already in use' in str(e) or '通常每个套接字地址' in str(e):
            if not silent_mode:
                _show_error_popup(
                    f'端口 {PORT} 已被占用。\n\n'
                    '解决方案：\n'
                    '1. 关闭之前运行的 web_server.py\n'
                    f'2. 或修改 config.json 中的 port 常量（当前为 {PORT}）'
                )
        else:
            if not silent_mode:
                _show_error_popup(f'Flask 启动失败：\n{type(e).__name__}: {e}')
            raise
