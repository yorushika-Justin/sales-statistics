"""
Render 部署用 — 销售汇总网页端
================================
从原版 web_server.py 改造，适配 Linux 服务器：
- PORT: 从环境变量读取（Render 自动分配）
- host: 0.0.0.0（允许外部访问）
- os.startfile: 跳过（服务器无此 API）
- 输出目录: 从环境变量或默认路径
- 去掉 tkinter 弹窗（服务器无 GUI）
"""
import sys
import os
import uuid
import tempfile
import traceback
from datetime import datetime

from flask import Flask, request, jsonify, send_file

# 将项目根目录加入 sys.path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 从本地适配层导入
from 程序端.汇总脚本 import OUTPUT_DIR, FONT_PATH, process_xls

# ==================== 配置 ====================
PORT = int(os.environ.get('PORT', 5000))
MAX_UPLOAD_SIZE = 10 * 1024 * 1024  # 10MB

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = MAX_UPLOAD_SIZE

# ==================== 前端页面 ====================
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
            font-family: "Microsoft YaHei", "微软雅黑", "PingFang SC", "Helvetica Neue", sans-serif;
            background: #f5f7fa;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            padding: 20px;
        }
        .container {
            background: white;
            border-radius: 12px;
            box-shadow: 0 2px 12px rgba(0, 0, 0, 0.08);
            padding: 40px;
            width: 100%;
            max-width: 500px;
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
        .preview {
            margin-top: 15px;
        }
        .preview img {
            max-width: 100%;
            border: 1px solid #eee;
            border-radius: 6px;
            margin-bottom: 10px;
        }
        .btn-download {
            display: inline-block;
            padding: 8px 20px;
            background: #4a90d9;
            color: white;
            border-radius: 6px;
            text-decoration: none;
            font-size: 14px;
            margin: 5px 5px 10px 0;
        }
        .btn-download:active {
            background: #357abd;
        }
        input[type="file"] { display: none; }
        .footer {
            text-align: center;
            margin-top: 20px;
            color: #999;
            font-size: 12px;
        }
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
        <div class="footer">
            <span>上传后自动生成 PNG 图片并显示</span>
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
                    // 显示图片预览 + 下载按钮
                    if (data.brand_png_url) {
                        const preview = document.createElement('div');
                        preview.className = 'preview';
                        // 品牌汇总
                        const img1 = document.createElement('img');
                        img1.src = data.brand_png_url;
                        img1.alt = '品牌汇总';
                        preview.appendChild(img1);
                        const btn1 = document.createElement('a');
                        btn1.href = data.brand_png_url + '?dl=1';
                        btn1.className = 'btn-download';
                        btn1.textContent = '保存品牌汇总';
                        preview.appendChild(btn1);
                        // 品名明细
                        if (data.item_png_url) {
                            const img2 = document.createElement('img');
                            img2.src = data.item_png_url;
                            img2.alt = '品名明细';
                            preview.appendChild(img2);
                            const btn2 = document.createElement('a');
                            btn2.href = data.item_png_url + '?dl=1';
                            btn2.className = 'btn-download';
                            btn2.textContent = '保存品名明细';
                            preview.appendChild(btn2);
                        }
                        status.appendChild(preview);
                    }
                } else {
                    showStatus(data.message, 'error');
                }
            })
            .catch(error => {
                showStatus('请求失败: ' + error.message, 'error');
            });
        }
    </script>
</body>
</html>
"""


# ==================== 路由 ====================
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

    # 保存到临时文件
    is_xlsx = file.filename.lower().endswith('.xlsx')
    ext = '.xlsx' if is_xlsx else '.xls'
    temp_filename = f'sales_upload_{uuid.uuid4().hex[:8]}{ext}'
    temp_path = os.path.join(tempfile.gettempdir(), temp_filename)
    file.save(temp_path)

    try:
        log_path = os.path.join(tempfile.gettempdir(), 'sales_web_error.log')
        brand_png, item_png, date_str, total = process_xls(temp_path, log_path=log_path)

        msg = f'已生成: {date_str}/{date_str}.png（总金额: {total:.2f} 元）'
        if item_png:
            msg += f' + 品名明细'
        else:
            msg += f'（品名明细生成失败）'

        # 返回下载链接
        brand_url = f'/download/{date_str}/{date_str}.png'
        item_url = f'/download/{date_str}/{date_str}_品名.png' if item_png else None

        return jsonify(
            success=True,
            message=msg,
            brand_png=f'{date_str}/{date_str}.png',
            item_png=f'{date_str}/{date_str}_品名.png' if item_png else None,
            brand_png_url=brand_url,
            item_png_url=item_url,
            total=total
        )
    except Exception as e:
        log_path = os.path.join(tempfile.gettempdir(), 'sales_web_error.log')
        try:
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(f'[{datetime.now()}] {traceback.format_exc()}\n')
        except Exception:
            pass
        return jsonify(success=False, message=f'{type(e).__name__}: {str(e)}')
    finally:
        try:
            os.remove(temp_path)
        except Exception:
            pass


@app.route('/download/<path:filename>')
def download(filename):
    """下载生成的 PNG 文件。?dl=1 触发下载，否则内联显示"""
    # 安全检查：只允许从 OUTPUT_DIR 下载 PNG
    filepath = os.path.join(OUTPUT_DIR, filename)
    filepath = os.path.normpath(filepath)

    # 防止路径穿越攻击
    if not filepath.startswith(os.path.normpath(OUTPUT_DIR)):
        return jsonify(success=False, message='非法路径'), 403

    if not os.path.exists(filepath):
        return jsonify(success=False, message='文件不存在'), 404

    # ?dl=1 → 触发下载，否则内联显示（给 <img> 用）
    as_dl = request.args.get('dl') == '1'
    return send_file(filepath, mimetype='image/png', as_attachment=as_dl)


@app.route('/health')
def health():
    """健康检查（Render 用）"""
    return jsonify(status='ok', output_dir=OUTPUT_DIR, font_path=FONT_PATH)


# ==================== 启动 ====================
if __name__ == '__main__':
    # 确保输出目录存在
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print(f"[启动] 销售汇总网页端 (Render)")
    print(f"[配置] PORT={PORT}, OUTPUT_DIR={OUTPUT_DIR}")
    print(f"[配置] FONT_PATH={FONT_PATH}")

    # Render 要求监听 0.0.0.0
    app.run(host='0.0.0.0', port=PORT, debug=False, threaded=True)
