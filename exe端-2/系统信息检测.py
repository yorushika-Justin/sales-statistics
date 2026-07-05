"""
销售汇总工具 — 系统信息检测脚本
================================

用途：
    在目标 PC 上运行，收集系统环境信息，
    用于排查 exe 或网页端运行时的问题。

用法：
    python 系统信息检测.py

输出：
    1. 控制台打印完整报告
    2. 自动生成 系统信息报告.txt 文件（同目录）

检测内容：
    - 操作系统信息
    - 磁盘空间（C/D 盘）
    - 字体文件（微软雅黑）
    - 输出目录权限
    - Python 环境（如果已安装）
    - pythonw.exe 路径（开机自启用）
    - 依赖包版本
    - 网络端口占用（网页端用）
    - exe 文件列表
    - 网页端依赖包（Flask/xlrd/openpyxl/Pillow）
    - 网页端文件（汇总脚本.py、web_server.py、config.json 等）
"""

import sys
import os
import platform
import subprocess
import ctypes
from datetime import datetime


def print_separator(title: str = ""):
    """打印分隔线"""
    if title:
        print(f"\n{'='*60}")
        print(f"  {title}")
        print(f"{'='*60}")
    else:
        print(f"{'-'*60}")


def check_os_info():
    """检查操作系统信息"""
    print_separator("操作系统信息")
    info = {
        "系统": platform.system(),
        "版本": platform.version(),
        "发行版": platform.platform(),
        "机器类型": platform.machine(),
        "处理器": platform.processor(),
        "计算机名": platform.node(),
        "用户名": os.getenv("USERNAME") or os.getenv("USER") or "未知",
    }
    for k, v in info.items():
        print(f"  {k}: {v}")
    return info


def check_disk_space():
    """检查磁盘空间"""
    import shutil as _shutil
    print_separator("磁盘空间")
    results = {}
    for drive in ["C:\\", "D:\\"]:
        try:
            if os.path.exists(drive):
                total, used, free = _shutil.disk_usage(drive)
                total_gb = total / (1024**3)
                used_gb = used / (1024**3)
                free_gb = free / (1024**3)
                print(f"  {drive} 盘:")
                print(f"    总空间: {total_gb:.1f} GB")
                print(f"    已使用: {used_gb:.1f} GB")
                print(f"    可用:   {free_gb:.1f} GB")
                results[drive] = {"total": total_gb, "used": used_gb, "free": free_gb}
            else:
                print(f"  {drive} 盘: 不存在")
                results[drive] = None
        except Exception as e:
            print(f"  {drive} 盘: 检测失败 - {e}")
            results[drive] = None
    return results


def check_font():
    """检查字体文件"""
    print_separator("字体检测")
    font_path = r"C:\Windows\Fonts\msyh.ttc"
    exists = os.path.exists(font_path)
    print(f"  微软雅黑字体: {font_path}")
    print(f"  状态: {'✓ 存在' if exists else '✗ 不存在'}")

    # 列出其他中文字体
    fonts_dir = r"C:\Windows\Fonts"
    chinese_fonts = []
    if os.path.exists(fonts_dir):
        for f in os.listdir(fonts_dir):
            if any(name in f.lower() for name in ["msyh", "simsun", "simhei", "kaiti", "fangsong"]):
                chinese_fonts.append(f)
    if chinese_fonts:
        print(f"  其他中文字体: {', '.join(chinese_fonts[:5])}...")

    return {"msyh_exists": exists, "chinese_fonts": chinese_fonts}


def check_output_dir():
    """检查输出目录"""
    print_separator("输出目录检测")
    output_dir = r"D:\销售汇总"

    # 检查目录是否存在
    exists = os.path.exists(output_dir)
    print(f"  输出目录: {output_dir}")
    print(f"  状态: {'✓ 存在' if exists else '✗ 不存在（运行时自动创建）'}")

    # 检查 D 盘是否存在
    d_drive_exists = os.path.exists("D:\\")
    print(f"  D 盘: {'✓ 存在' if d_drive_exists else '✗ 不存在'}")

    # 检查写入权限
    can_write = False
    if d_drive_exists:
        try:
            test_dir = os.path.join(output_dir, "_test_write")
            os.makedirs(test_dir, exist_ok=True)
            test_file = os.path.join(test_dir, "test.txt")
            with open(test_file, "w") as f:
                f.write("test")
            os.remove(test_file)
            os.rmdir(test_dir)
            can_write = True
            print(f"  写入权限: ✓ 正常")
        except Exception as e:
            print(f"  写入权限: ✗ 失败 - {e}")
    else:
        print(f"  写入权限: ✗ D 盘不存在")

    return {"output_dir_exists": exists, "d_drive_exists": d_drive_exists, "can_write": can_write}


def check_python():
    """检查 Python 环境"""
    print_separator("Python 环境")
    info = {
        "python_path": sys.executable,
        "python_version": sys.version,
        "platform": sys.platform,
    }
    for k, v in info.items():
        print(f"  {k}: {v}")

    # 检测 pythonw.exe 路径（开机自启用）
    python_dir = os.path.dirname(sys.executable)
    pythonw_path = os.path.join(python_dir, "pythonw.exe")
    if os.path.exists(pythonw_path):
        info["pythonw_path"] = pythonw_path
        print(f"  pythonw.exe: ✓ {pythonw_path}")
    else:
        # 尝试从 PATH 中查找
        import shutil as _shutil
        pythonw_in_path = _shutil.which("pythonw")
        if pythonw_in_path:
            info["pythonw_path"] = pythonw_in_path
            print(f"  pythonw.exe: ✓ {pythonw_in_path} (从 PATH)")
        else:
            info["pythonw_path"] = None
            print(f"  pythonw.exe: ✗ 未找到")

    # 检查依赖包
    print_separator("依赖包版本")
    packages = ["xlrd", "openpyxl", "Pillow", "tkinterdnd2", "Flask", "pyinstaller"]
    pkg_status = {}
    for pkg in packages:
        try:
            if pkg == "Pillow":
                import PIL
                version = PIL.__version__
            elif pkg == "tkinterdnd2":
                import tkinterdnd2
                version = getattr(tkinterdnd2, "__version__", "已安装")
            else:
                mod = __import__(pkg)
                version = getattr(mod, "__version__", "已安装")
            print(f"  {pkg}: ✓ {version}")
            pkg_status[pkg] = {"installed": True, "version": version}
        except ImportError:
            print(f"  {pkg}: ✗ 未安装")
            pkg_status[pkg] = {"installed": False, "version": None}
        except Exception as e:
            print(f"  {pkg}: ? 检测异常 - {e}")
            pkg_status[pkg] = {"installed": None, "version": None}

    return {**info, "packages": pkg_status}


def check_port():
    """检查端口占用（网页端用）"""
    print_separator("端口检测（网页端）")
    ports = [5000, 5001]
    port_status = {}
    for port in ports:
        try:
            import socket
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(2)
            result = sock.connect_ex(('127.0.0.1', port))
            sock.close()
            if result == 0:
                print(f"  端口 {port}: ✗ 已被占用")
                port_status[port] = "occupied"
            else:
                print(f"  端口 {port}: ✓ 可用")
                port_status[port] = "available"
        except Exception as e:
            print(f"  端口 {port}: ? 检测异常 - {e}")
            port_status[port] = "error"
    return port_status


def check_exe_files():
    """检查 exe 文件"""
    print_separator("exe 文件检测")
    exe_dir = os.path.dirname(os.path.abspath(__file__))
    exe_files = []

    for f in os.listdir(exe_dir):
        if f.endswith(".exe"):
            fpath = os.path.join(exe_dir, f)
            size_mb = os.path.getsize(fpath) / (1024**2)
            mtime = datetime.fromtimestamp(os.path.getmtime(fpath)).strftime("%Y-%m-%d %H:%M")
            exe_files.append({"name": f, "size_mb": size_mb, "mtime": mtime})
            print(f"  {f}: {size_mb:.1f} MB (修改时间: {mtime})")

    if not exe_files:
        print("  未找到 exe 文件")

    return exe_files


def check_web_deps():
    """检查网页端依赖包（Flask + 核心库）"""
    print_separator("网页端依赖检测")
    web_packages = {
        "flask": "Flask",
        "xlrd": "xlrd",
        "openpyxl": "openpyxl",
        "PIL": "Pillow",
    }
    results = {}
    all_ok = True
    for import_name, display_name in web_packages.items():
        try:
            mod = __import__(import_name)
            version = getattr(mod, "__version__", "已安装")
            print(f"  {display_name}: ✓ {version}")
            results[display_name] = {"installed": True, "version": version}
        except ImportError:
            print(f"  {display_name}: ✗ 未安装")
            results[display_name] = {"installed": False, "version": None}
            all_ok = False
        except Exception as e:
            print(f"  {display_name}: ? 检测异常 - {e}")
            results[display_name] = {"installed": None, "version": None}
            all_ok = False

    if all_ok:
        print("\n  ✓ 网页端依赖完整，可直接运行")
    else:
        print("\n  ⚠ 缺少依赖，安装命令:")
        print("    pip install Flask xlrd==1.2.0 openpyxl Pillow")

    return results


def check_web_files():
    """检查网页端文件是否就位"""
    print_separator("网页端文件检测")

    # 当前脚本所在目录（exe端-2）
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # 项目根目录（上级目录）
    project_root = os.path.dirname(script_dir)

    files_to_check = {
        "汇总脚本.py": os.path.join(project_root, "汇总脚本.py"),
        "网页端-2/web_server.py": os.path.join(project_root, "网页端-2", "web_server.py"),
        "网页端-2/config.json": os.path.join(project_root, "网页端-2", "config.json"),
        "网页端-2/启动网页端.vbs": os.path.join(project_root, "网页端-2", "启动网页端.vbs"),
        "网页端-2/启动网页端.bat": os.path.join(project_root, "网页端-2", "启动网页端.bat"),
        "网页端-2/开机自启.bat": os.path.join(project_root, "网页端-2", "开机自启.bat"),
    }

    results = {}
    for name, path in files_to_check.items():
        exists = os.path.exists(path)
        status = "✓ 存在" if exists else "✗ 不存在"
        print(f"  {name}: {status}")
        results[name] = exists

    # 读取 config.json 内容（如果存在）
    config_path = files_to_check.get("网页端-2/config.json")
    if config_path and os.path.exists(config_path):
        try:
            import json
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            print(f"\n  config.json 配置:")
            print(f"    output_dir: {config.get('output_dir', '未设置')}")
            print(f"    font_path:  {config.get('font_path', '未设置')}")
            print(f"    port:       {config.get('port', '未设置')}")
            results["config"] = config
        except Exception as e:
            print(f"\n  config.json 读取失败: {e}")
            results["config"] = None

    return results


def generate_report():
    """生成完整报告"""
    print("\n" + "="*60)
    print("  销售汇总工具 — 系统信息检测报告")
    print(f"  检测时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)

    report = {}
    report["检测时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    report["os"] = check_os_info()
    report["disk"] = check_disk_space()
    report["font"] = check_font()
    report["output_dir"] = check_output_dir()
    report["python"] = check_python()
    report["port"] = check_port()
    report["exe"] = check_exe_files()
    report["web_deps"] = check_web_deps()
    report["web_files"] = check_web_files()

    # 汇总
    print_separator("检测汇总")
    issues = []

    if not report["font"]["msyh_exists"]:
        issues.append("微软雅黑字体缺失")
    if not report["output_dir"]["d_drive_exists"]:
        issues.append("D 盘不存在")
    if not report["output_dir"]["can_write"]:
        issues.append("输出目录无写入权限")

    for pkg, status in report["python"]["packages"].items():
        if not status["installed"]:
            issues.append(f"依赖包 {pkg} 未安装")

    # 检查网页端依赖
    web_deps_ok = all(s["installed"] for s in report["web_deps"].values())
    if not web_deps_ok:
        issues.append("网页端依赖不完整（Flask/xlrd/openpyxl/Pillow）")

    # 检查关键文件
    if not report["web_files"].get("汇总脚本.py", False):
        issues.append("汇总脚本.py 不存在（网页端核心依赖）")
    if not report["web_files"].get("网页端-2/web_server.py", False):
        issues.append("网页端-2/web_server.py 不存在")

    if issues:
        print("  发现问题:")
        for issue in issues:
            print(f"    ⚠ {issue}")
    else:
        print("  ✓ 所有检测项正常")

    # 保存报告
    report_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "系统信息报告.txt")
    try:
        with open(report_file, "w", encoding="utf-8") as f:
            f.write("销售汇总工具 — 系统信息检测报告\n")
            f.write(f"检测时间: {report['检测时间']}\n")
            f.write("="*60 + "\n\n")

            f.write("【操作系统】\n")
            for k, v in report["os"].items():
                f.write(f"  {k}: {v}\n")

            f.write("\n【磁盘空间】\n")
            for drive, info in report["disk"].items():
                if info:
                    f.write(f"  {drive}: 总 {info['total']:.1f} GB, 已用 {info['used']:.1f} GB, 可用 {info['free']:.1f} GB\n")
                else:
                    f.write(f"  {drive}: 不存在或检测失败\n")

            f.write("\n【字体】\n")
            f.write(f"  微软雅黑: {'存在' if report['font']['msyh_exists'] else '缺失'}\n")

            f.write("\n【输出目录】\n")
            f.write(f"  D 盘: {'存在' if report['output_dir']['d_drive_exists'] else '缺失'}\n")
            f.write(f"  输出目录: {'存在' if report['output_dir']['output_dir_exists'] else '不存在'}\n")
            f.write(f"  写入权限: {'正常' if report['output_dir']['can_write'] else '失败'}\n")

            f.write("\n【Python 环境】\n")
            f.write(f"  路径: {report['python']['python_path']}\n")
            f.write(f"  版本: {report['python']['python_version']}\n")
            if report['python'].get('pythonw_path'):
                f.write(f"  pythonw.exe: {report['python']['pythonw_path']}\n")
            else:
                f.write(f"  pythonw.exe: 未找到\n")

            f.write("\n【依赖包】\n")
            for pkg, status in report["python"]["packages"].items():
                if status["installed"]:
                    f.write(f"  {pkg}: {status['version']}\n")
                else:
                    f.write(f"  {pkg}: 未安装\n")

            f.write("\n【端口状态】\n")
            for port, status in report["port"].items():
                f.write(f"  端口 {port}: {status}\n")

            f.write("\n【exe 文件】\n")
            for exe in report["exe"]:
                f.write(f"  {exe['name']}: {exe['size_mb']:.1f} MB ({exe['mtime']})\n")

            # 新增：网页端依赖
            f.write("\n【网页端依赖】\n")
            for pkg, status in report["web_deps"].items():
                if status["installed"]:
                    f.write(f"  {pkg}: {status['version']}\n")
                else:
                    f.write(f"  {pkg}: 未安装\n")

            # 新增：网页端文件
            f.write("\n【网页端文件】\n")
            for name, exists in report["web_files"].items():
                if name != "config":
                    f.write(f"  {name}: {'存在' if exists else '不存在'}\n")

            # 新增：config.json 内容
            if report["web_files"].get("config"):
                f.write("\n【config.json 配置】\n")
                config = report["web_files"]["config"]
                f.write(f"  output_dir: {config.get('output_dir', '未设置')}\n")
                f.write(f"  font_path:  {config.get('font_path', '未设置')}\n")
                f.write(f"  port:       {config.get('port', '未设置')}\n")

            if issues:
                f.write("\n【发现问题】\n")
                for issue in issues:
                    f.write(f"  ⚠ {issue}\n")
            else:
                f.write("\n【结论】所有检测项正常\n")

        print(f"\n  报告已保存: {report_file}")
    except Exception as e:
        print(f"\n  报告保存失败: {e}")

    return report


if __name__ == "__main__":
    generate_report()
    print("\n按 Enter 键退出...")
    input()
