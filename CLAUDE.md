# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

销售汇总工具 — xls/xlsx → 按品牌 sum → PNG + 按品名列明细 → PNG。**三方案并联**:

- **exe 方案**:`汇总脚本.py` + pyinstaller,34 MB 单文件,离线可分发
- **网页端方案**:`网页端/web_server.py` + Flask,Edge 多标签页并发,无需打包
- **程序端方案**:`程序端/web_server.py` + Render.com 云部署,手机端可用,无需电脑

三种方案**共用** `汇总脚本.py` 的核心函数。改 `汇总脚本.py` 影响三端。

## 硬编码常量(改路径/端口需改源码)

| 常量 | 值 | 位置 |
|---|---|---|
| `OUTPUT_DIR` | `D:\16678\相册\销售额\` | `汇总脚本.py` 顶部 |
| `FONT_PATH` | `C:\Windows\Fonts\msyh.ttc`(微软雅黑) | `汇总脚本.py` 顶部 |
| Flask 端口 | `5000`(被占可改 5001) | `网页端/web_server.py` 顶部 `PORT` 常量 |
| `vbs 启动器` | `网页端/启动网页端.vbs` | v1.1 新增,双击即静默启动(日常用) |
| `bat 启动器` | `网页端/启动网页端.bat` | v1.1 新增,命令行启动(开发/调试用) |

## 常用命令

```powershell
# === exe 方案 ===
# 开发运行
python 汇总脚本.py "<xls/xlsx 路径>"

# 打包 exe
pyinstaller --onefile --noconsole --exclude-module pygame --clean --noconfirm --name=销售汇总 汇总脚本.py
# 输出: dist\销售汇总.exe (~34 MB)

# === 网页端方案 ===
# 日常使用(v1.1 新增,推荐)
双击 网页端\启动网页端.vbs
# 静默启动,无窗口,浏览器自动打开 http://127.0.0.1:5000

# 开发/调试(看 Flask 日志)
python 网页端\web_server.py
# 或双击 网页端\启动网页端.bat
```

## 核心架构

### exe 方案(`汇总脚本.py`,~810 行)

| 函数 | 职责 |
|---|---|
| `read_xls(path, log_path)` | xlrd 读 `.xls` → 自动扫描表头定位"品牌名称"和"销售金额"列 → 从 Row 2 提取内部日期(优先)→ 按品牌 sum → 降序返回 `({品牌: 金额}, 日期字符串)`。金额千分位逗号自动清理;异常跳过行计数写入日志 |
| `read_xlsx(path, log_path)` | **v1.7 新增** openpyxl 读 `.xlsx`，逻辑同 `read_xls` |
| `read_xls_items(path, log_path)` | **v1.6 新增** 读 `.xls` → 按品名逐行提取明细(不汇总)→ 返回 `[(品名, 品牌, 条码, 数量, 金额, 库存), ...]`,保持原顺序 |
| `read_xlsx_items(path, log_path)` | **v1.7 新增** openpyxl 读 `.xlsx`，逻辑同 `read_xls_items` |
| `draw_table(data, date_str, output_path)` | Pillow 画品牌汇总表格(标题+表头+数据行+加粗总计行)→ 保存 PNG。图片宽度自适应标题和品牌名;品牌名超宽时自动截断加"…" |
| `draw_table_detail(items, date_str, output_path)` | **v1.6 新增** Pillow 画品名明细表格(7 列:行号/国际条码/品名/品牌/销售数量/销售金额/库存)→ 保存 PNG。品名/品牌/条码超宽自动截断;总计行统计总数量+总金额 |
| `_extract_date_from_xls(sh)` | 从 xls sheet Row 2 提取日期范围;`start == end`(当天导出当天)→ 直接返回 `start`;否则**结束日期 -1 天**为实际最后完整统计日(数据滞后性);同一天→`YYYY-MM-DD`,跨天→`YYYY-MM-DD至YYYY-MM-DD`(v1.5 修复 start==end 边界 bug) |
| `_extract_date_from_xlsx(sh)` | **v1.7 新增** 从 xlsx sheet Row 3 提取日期(openpyxl 1-indexed),逻辑同 `_extract_date_from_xls` |
| `launch_gui()` | tkinterdnd2 创建拖拽窗口(拖放区注册在 root),双击 exe 时启动 |
| `extract_date_from_filename()` | 从 xls 文件名提取 `YYYY-MM-DD`(现仅为 fallback);失败用今天 |
| `show_error(msg, log_path, parent)` | tkinter 弹窗 + 写 `汇总错误.log`(仅在有活跃异常时写 traceback)。GUI 模式传入 parent 避免创建第二个 Tk root |
| `main()` | 有 argv → 直接处理;无 argv → 启动 GUI |

### 网页端方案(`网页端/web_server.py`,~450 行)

复用 `汇总脚本.py` 的 `read_xls` / `read_xlsx` / `draw_table` / `OUTPUT_DIR`(通过 `sys.path` 注入上级目录)。

| 组件 | 职责 |
|---|---|
| `HTML_PAGE` | 内嵌前端(纯 HTML/JS/CSS,无外部 CDN) |
| `@app.route('/')` | 返回前端页面 |
| `@app.route('/upload', POST)` | 接收 xls/xlsx → 保存到 `tempfile.gettempdir()`(文件名带 `uuid.uuid4().hex[:8]` 后缀防多标签页冲突)→ 根据后缀调用 `read_xls`/`read_xlsx` + `draw_table` + `read_xls_items`/`read_xlsx_items` + `draw_table_detail` → 生成品牌汇总+品名明细 2 个 PNG → `os.startfile` 异步打开品牌汇总 → 返回 JSON(含 `brand_png` + `item_png`)。含 `MAX_CONTENT_LENGTH`(10MB)限制 |
| `@app.route('/shutdown', POST)` | 停止服务(`os._exit(0)`) |
| `_show_popup(msg, kind)` | 统一的 tkinter 弹窗工具函数(v1.2 新增) |
| `_check_already_running()` | socket 探测 `PORT` 端口,已占则弹窗询问"是否打开浏览器" |
| `_check_output_dir()` | 检测 `OUTPUT_DIR` 是否存在 |
| `_show_error_popup(msg)` | tkinter 弹窗(静默场景下提示错误) |
| `open_browser()` | socket 轮询探测端口就绪后 `os.startfile` 打开浏览器(v1.2 改进) |
| `if __name__ == '__main__'` | 启动检查(端口+目录)→ 启动 Flask(`app.run(host='127.0.0.1', port=PORT, threaded=True)`),捕获 `OSError` |

### 程序端方案(`程序端/web_server.py` + Render.com 云部署)

适配层模式：`程序端/汇总脚本.py` 从原版导入核心函数，覆盖服务器不兼容配置（tkinter mock、OUTPUT_DIR、FONT_PATH）。

| 组件 | 职责 |
|---|---|
| `程序端/web_server.py` | Flask 入口，host=0.0.0.0，PORT 从环境变量读取 |
| `程序端/汇总脚本.py` | 适配层：mock tkinter、覆盖 OUTPUT_DIR/FONT_PATH、提供 `process_xls()` 封装 |
| `程序端/msyh.ttc` | 微软雅黑字体(19MB)，嵌入仓库（Render 无 sudo 无法 apt-get） |
| `@app.route('/')` | 返回移动端前端（响应式，内嵌 HTML/JS/CSS） |
| `@app.route('/upload', POST)` | 接收 xls/xlsx → 调用 `process_xls()` → 返回 JSON(含下载链接) |
| `@app.route('/download/<path>')` | PNG 下载路由：`?dl=1` 触发下载，否则内联显示(给 `<img>` 用) |
| `@app.route('/health')` | 健康检查(Render 用) |
| `requirements.txt` | flask、xlrd、openpyxl、Pillow |
| `Procfile` | `web: python 程序端/web_server.py` |

## 数据流

**exe 方案**:xls/xlsx → xlrd/openpyxl 读取 → 从 Row 2 提取内部日期(结束日期 -1 天为最后完整统计日,文件名仅 fallback)→ 按品牌 sum → 降序排序 → Pillow 画品牌汇总表(宽度自适应)→ 同时按品名列明细 → 2 个 PNG 输出到 `OUTPUT_DIR/<日期>/` 子文件夹 → `os.startfile()` 打开品牌汇总

**网页端方案**:Edge 浏览器拖入 xls/xlsx → POST /upload → Flask 保存临时文件(唯一名)→ 复用 `read_xls`/`read_xlsx` + `draw_table` + `read_xls_items`/`read_xlsx_items` + `draw_table_detail` → 2 个 PNG 输出到 `OUTPUT_DIR/<日期>/` 子文件夹 → `os.startfile()` 打开品牌汇总 → 返回 JSON 状态给前端

## 六种使用方式

1. 拖 xls/xlsx 到 `dist\销售汇总.exe` 图标
2. 双击 `销售汇总.exe` 打开拖拽窗口
3. 命令行 `python 汇总脚本.py <xls/xlsx 路径>`
4. 双击 `网页端\启动网页端.vbs`(v1.1 静默日常用,推荐)
5. 双击 `网页端\启动网页端.bat` 或 `python 网页端\web_server.py`(开发/调试,看 Flask 日志)
6. **手机端**：夸克打开 `https://sales-statistics-rdle.onrender.com`，选文件上传 → 预览 → 下载到手机(无需电脑)

## 关键约束

### 数据/格式
- **xlrd 1.2.0** — 最后支持 `.xls` 的版本
- **openpyxl 3.1.5** — 支持 `.xlsx`(Excel 2007+ 格式)
- **后缀分发**:根据 `.xls` / `.xlsx` 后缀自动选择 xlrd 或 openpyxl
- **表头检测**:前 20 行扫描,品牌列和金额列均为**子串匹配**(兼容"销售金额(元)"等变体)
- **过滤规则**:空品牌 / 以"合计"开头 / 品牌为"总"或"总计" / 金额为 0 的行被跳过
- **日期规则**:xls/xlsx Row 2 的"结束日期"含当天未完整数据,**实际最后完整统计日 = end_date - 1 天**;`start == end`(当天导出当天)直接返回 `start`;同一天显示 `YYYY-MM-DD`,跨天显示 `YYYY-MM-DD至YYYY-MM-DD`

### 路径/字体
- **字体硬编码** `C:\Windows\Fonts\msyh.ttc`(微软雅黑),丢失时 crash,无 fallback
- **输出路径**:固定 `D:\16678\相册\销售额\<日期>\<日期>.png`(品牌汇总)+`<日期>_品名.png`(品名明细),按日期自动创建子文件夹,同名覆盖无提示
- **PIL 隐式 import pygame** — `--exclude-module pygame` 不生效,exe 体积 34 MB 已接受

### 网页端特有
- **端口 5000 硬编码**(顶部 `PORT` 常量),被占时 tkinter 弹窗提示;改端口只需改 `PORT` 一处
- **错误日志路径分两种**:exe 模式写 xls 同目录 `汇总错误.log`;网页端写 `tempfile.gettempdir()/sales_web_error.log`
- **改 `汇总脚本.py` 后网页端必须重启服务**(`app.run` 非热加载,需 Ctrl+C 重启)
- **完全离线可用**:前端内嵌,无外部 CDN,所有 HTTP 引用均 `127.0.0.1`

### 程序端(Render 云部署)特有
- **公网地址**: `https://sales-statistics-rdle.onrender.com`
- **GitHub 仓库**: `git@github.com:yorushika-Justin/sales-statistics.git`
- **Render 免费层**: 750 小时/月,15 分钟不活跃休眠(冷启动约 30 秒),无硬性到期日
- **改代码后需推送 GitHub**，Render 自动重新部署(约 2-3 分钟)
- **不修改原版**:所有改动在 `程序端/` 目录，通过适配层复用原版逻辑
- **字体嵌入**: `程序端/msyh.ttc`(19MB)直接放入仓库,Render 无 sudo 无法 apt-get

## 依赖

### exe 方案
| 包 | 版本 | 用途 |
|---|---|---|
| Python | 3.13.12 | `D:\Python\`,无 venv |
| xlrd | 1.2.0 | 读 `.xls` |
| openpyxl | 3.1.5 | 读 `.xlsx`(v1.7 新增) |
| Pillow | 12.1.1 | 绘图 |
| tkinterdnd2 | 0.4.4.1 | GUI 拖拽窗口 (~1.4 MB) |
| pyinstaller | 6.19.0 | 打包 |

### 网页端方案(在 exe 方案基础上增加)
| 包 | 版本 | 用途 |
|---|---|---|
| Flask | 3.1.3 | Web 框架 |

## 相关文档

- `技术文档.md` — 完整技术文档(exe 方案:架构/打包/维护/限制/优化方向)
- `网页端/技术文档.md` — 网页端方案技术文档
- `程序端/技术文档.md` — 程序端方案技术文档(Render 云部署:架构/适配层/字体/下载/限制)
- `工作记录/2026-06-05_销售汇总工具开发记录.md` — 开发时间线与经验沉淀(v1.0 → v1.5)
- `对话.txt` — 完整对话流水(给其他 AI 验证用)
- `提示文档.txt` — 原始需求
- `原表格.csv` — 期望输出样例

---

**最新版本**:exe v1.7 / 网页端 v1.5 / 网页端-2 v2.0 / exe端-2 v2.0 / 程序端 v1.0(2026-07-04)。变更细节见 `技术文档.md` 11 节 + `网页端/技术文档.md` 11 节 + `网页端-2/技术文档.md` 12 节 + `exe端-2/技术文档.md` 10 节 + `程序端/技术文档.md`。

**程序端 v1.0 新增**(2026-07-04):Render.com 云部署，手机端可用。适配层模式（不修改原版）；tkinter mock；字体嵌入仓库(19MB)；/download 路由支持图片预览+下载到手机。公网地址 `https://sales-statistics-rdle.onrender.com`。详见 `程序端/技术文档.md`。

**exe端-2 v2.0 新增**(2026-07-01):独立版本，输出路径改为 `D:\销售汇总\`（不依赖特定用户名路径）。新增 `系统信息检测.py`（环境检测脚本）；改进品名明细失败时的错误提示。详见 `exe端-2/技术文档.md`。

**网页端-2 v2.0 新增**(2026-06-18):独立备份版本，输出路径改为配置文件读取。新增 `config.json`（输出路径/字体/端口）；`web_server.py` v2.0 从 config.json 读取配置；`开机自启.bat` 改为中文界面；`index.html` 同步支持 .xlsx。详见 `网页端-2/技术文档.md` + `网页端-2/工作记录.md`。

**网页端 v1.5 新增**(2026-06-11):开机自启功能。新增 `开机自启.bat` 管理脚本(注册/注销/查看状态)；使用 Windows 任务计划程序实现开机自启(登录后延迟 30 秒启动)；`web_server.py` 支持 `--silent` 参数(静默模式：不弹浏览器，端口已占时静默退出)。详见 `工作记录/2026-06-11_网页端开机自启功能.md`。

**v1.7 新增**(2026-06-08):支持 `.xlsx` 格式(Excel 2007+)。新增 `read_xlsx()` / `read_xlsx_items()` 函数(openpyxl)，入口根据后缀 `.xls` / `.xlsx` 自动分发。

**v1.6.1 修复**(2026-06-08):`draw_table_detail()` 库存列 rectangle 右边框缺失 — 3 处 `draw.rectangle` 的右边界从 `x7` 改为 `x7 + stock_col_w`(引入新变量 `x_right`)。详见 `工作记录/2026-06-05_销售汇总工具开发记录.md` 第二十一节 + `工作记录/2026-06-08_代码审查_v2.md` F1。

**v1.6.1 同期修复**(2026-06-08):`网页端/web_server.py` 内联 `__import__("datetime")` 改为顶部 import。详见 `工作记录/2026-06-05_销售汇总工具开发记录.md` 第二十二节 + `工作记录/2026-06-08_代码审查_v2.md` F5。

**v1.6.1 同期修复 2**(2026-06-08):`网页端/web_server.py` upload 路由品名失败静默吞改为对齐 on_drop/main。详见 `工作记录/2026-06-05_销售汇总工具开发记录.md` 第二十四节 + `工作记录/2026-06-08_代码审查_v3_CodeGraph.md` F16。

**exe 重打包**(2026-06-08 15:00):`dist\销售汇总.exe` 重打包为 v1.7(35 MB,36,650,968 字节, +500KB vs v1.6.1,openpyxl 依赖),冒烟测试通过(xlsx + xls 均生成 PNG 正确)。详见 `工作记录/2026-06-05_销售汇总工具开发记录.md` 第二十五节。

**v1.6.1 旧打包**(2026-06-08 14:10):已备份至 `备份_2026-06-08_v1.5_v1.2/销售汇总.exe.v1.6.1.bak`。
