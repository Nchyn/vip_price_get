import os
import json
import time
import random
import threading
from dataclasses import dataclass
from typing import Optional, Tuple, Union

import pandas as pd
from flask import Flask, request, jsonify, render_template_string
from playwright.sync_api import sync_playwright, Browser, BrowserContext, Page

# ==============================================
# 1. 配置层（所有常量集中管理，修改更便捷）
# ==============================================
@dataclass
class Config:
    BASE_PATH: str = os.path.abspath(".")
    INPUT_FILE: str = os.path.join(BASE_PATH, "商品清单.csv")
    AUTH_FILE: str = os.path.join(BASE_PATH, "vip_auth.json")
    CHROME_PATH: str = os.path.join(BASE_PATH, "chrome-win64", "chrome.exe")
    
    # 爬虫配置
    REQUEST_TIMEOUT: int = 15000
    LOGIN_WAIT_TIMEOUT: int = 600
    WAIT_MIN: int = 40
    WAIT_MAX: int = 70
    MAX_LOGS: int = 200
    
    # 浏览器配置
    USER_AGENT: str = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    BROWSER_ARGS: list = None
    
    def __post_init__(self):
        self.BROWSER_ARGS = ['--disable-blink-features=AutomationControlled', '--no-sandbox']

CONFIG = Config()

# ==============================================
# 2. 线程安全状态管理层
# ==============================================
class AppState:
    _lock = threading.Lock()
    
    def __init__(self):
        self.is_running: bool = False
        self.stop_flag: bool = False
        self.logs: list = []

    def add_log(self, msg: str):
        with self._lock:
            print(msg, flush=True)
            self.logs.append(msg)
            if len(self.logs) > CONFIG.MAX_LOGS:
                self.logs.pop(0)

# 全局单例状态
state = AppState()

# ==============================================
# 3. 工具层（公共方法）
# ==============================================
def init_csv_file() -> None:
    """初始化CSV文件（不存在则创建）"""
    if not os.path.exists(CONFIG.INPUT_FILE):
        df = pd.DataFrame(columns=["ID", "品牌", "标题", "特卖价", "原价", "折扣"])
        df.to_csv(CONFIG.INPUT_FILE, index=False, encoding="utf-8-sig")

def check_and_fix_auth_file() -> None:
    """自动检测并修复损坏的Cookie文件"""
    if os.path.exists(CONFIG.AUTH_FILE):
        try:
            # 尝试解析JSON，校验合法性
            with open(CONFIG.AUTH_FILE, 'r', encoding='utf-8') as f:
                json.load(f)
        except json.JSONDecodeError:
            # 文件损坏，自动删除
            os.remove(CONFIG.AUTH_FILE)
            state.add_log(">>> [修复] 检测到损坏的登录文件，已自动重置")

# ==============================================
# 4. 爬虫核心层
# ==============================================
class VIPCrawler:
    def __init__(self):
        self.browser: Optional[Browser] = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None

    def _init_browser(self) -> None:
        """初始化浏览器上下文，加载Cookie"""
        # 前置：修复损坏的认证文件
        check_and_fix_auth_file()
        
        playwright = sync_playwright().start()
        launch_options = {
            "headless": False,
            "args": CONFIG.BROWSER_ARGS
        }
        if os.path.exists(CONFIG.CHROME_PATH):
            launch_options["executable_path"] = CONFIG.CHROME_PATH

        self.browser = playwright.chromium.launch(**launch_options)
        
        # 加载登录态
        context_options = {"user_agent": CONFIG.USER_AGENT}
        if os.path.exists(CONFIG.AUTH_FILE):
            context_options["storage_state"] = CONFIG.AUTH_FILE
        
        self.context = self.browser.new_context(**context_options)
        self.page = self.context.new_page()
        # 防爬虫检测
        self.page.add_init_script("Object.defineProperty(navigator,'webdriver',{get:()=>undefined})")

    def _handle_login(self, pid: str) -> Optional[Page]:
        """处理登录拦截与多标签页"""
        state.add_log(">>> [风控提示] 请在浏览器中完成登录/验证")
        start_time = time.time()

        while True:
            if state.stop_flag:
                return None
            if time.time() - start_time > CONFIG.LOGIN_WAIT_TIMEOUT:
                state.add_log(">>> 登录等待超时")
                return None

            for page in self.context.pages:
                try:
                    url, title = page.url, page.title()
                    # 识别商品详情页（登录成功）
                    if "detail.vip.com" in url and pid in url and "登录" not in title:
                        page.bring_to_front()
                        page.wait_for_timeout(2000)
                        # 保存Cookie
                        self.context.storage_state(path=CONFIG.AUTH_FILE)
                        # 关闭多余登录页面
                        for p in self.context.pages:
                            if p != page and ("passport" in p.url or "login" in p.url):
                                p.close()
                        state.add_log(">>> [✅ 登录成功] 继续爬取")
                        return page
                except:
                    continue
            time.sleep(1.5)

    def _fetch_data(self, pid: str) -> Union[Tuple[str, str, str, str, str], str, None]:
        """爬取单个商品数据"""
        url = f"https://detail.vip.com/detail-1234-{pid}.html"
        try:
            self.page.goto(url, timeout=CONFIG.REQUEST_TIMEOUT, wait_until="commit")
        except:
            pass

        # 登录拦截检测
        if "passport.vip.com" in self.page.url or "detail.vip.com" not in self.page.url:
            new_page = self._handle_login(pid)
            if new_page:
                self.page = new_page
                return "NEED_RETRY"
            return None

        # 解析数据
        try:
            self.page.wait_for_selector(".sp-price", timeout=8000)
            brand = self.page.locator(".J_brandName").first.inner_text().strip() or "未知"
            title = self.page.locator(".pib-title-detail").first.inner_text().strip() or "未知"
            sale = self.page.locator(".sp-price").first.inner_text().strip()
            market = self.page.locator(".marketPrice").first.inner_text().strip() if self.page.locator(".marketPrice").count() else "N/A"
            disc = self.page.locator(".sp-discount").first.inner_text().strip() if self.page.locator(".sp-discount").count() else "N/A"
            return brand, title[:40], sale, market, disc
        except Exception as e:
            if "passport" in self.page.url:
                new_page = self._handle_login(pid)
                if new_page:
                    self.page = new_page
                    return "NEED_RETRY"
            return None

    def _wait_random_time(self):
        """随机延时防封禁"""
        wait_time = random.randint(CONFIG.WAIT_MIN, CONFIG.WAIT_MAX)
        for i in range(wait_time, 0, -1):
            if state.stop_flag:
                break
            msg = f"安全间隔: {i} 秒..."
            with state._lock:
                if state.logs and "安全间隔" in state.logs[-1]:
                    state.logs[-1] = msg
                else:
                    state.logs.append(msg)
            time.sleep(1)

    def run(self):
        """爬虫主入口"""
        state.is_running = True
        state.stop_flag = False
        init_csv_file()

        try:
            df = pd.read_csv(CONFIG.INPUT_FILE, dtype=str, encoding="utf-8-sig").fillna("")
            total = len(df)
            state.add_log(f">>> 开始任务，共 {total} 个商品")

            self._init_browser()

            for index, row in df.iterrows():
                if state.stop_flag:
                    break

                pid = str(row["ID"]).strip()
                if row.get("标题"):
                    continue

                state.add_log(f"[{index+1}/{total}] 正在抓取: {pid}")
                success = False

                while not success:
                    if state.stop_flag:
                        break

                    result = self._fetch_data(pid)
                    if result == "NEED_RETRY":
                        continue

                    if isinstance(result, tuple):
                        # 修复：逐个字段赋值
                        df.at[index, "品牌"], df.at[index, "标题"], df.at[index, "特卖价"], df.at[index, "原价"], df.at[index, "折扣"] = result
                        df.to_csv(CONFIG.INPUT_FILE, index=False, encoding="utf-8-sig")
                        state.add_log(f"✅ 成功: {result[0]} | {result[2]}")
                        success = True
                    else:
                        state.add_log(f"❌ 失败: {pid} 无法获取")
                        success = True

                # 延时
                if not state.stop_flag and index < total - 1:
                    self._wait_random_time()

            state.add_log(">>> 任务已完成")

        except Exception as e:
            state.add_log(f">>> 致命错误: {str(e)}")
        finally:
            # 资源释放
            if self.browser:
                self.browser.close()
            state.is_running = False

# ==============================================
# 5. Flask Web服务层
# ==============================================
app = Flask(__name__)
HTML_TPL = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8"><title>唯品会爬虫 Pro 重构版</title>
    <script src="https://cdn.jsdelivr.net/npm/handsontable/dist/handsontable.full.min.js"></script>
    <link rel="stylesheet" href://cdn.jsdelivr.net/npm/handsontable/dist/handsontable.full.min.css">
    <style>
        body { font-family: sans-serif; background: #f0f2f5; margin: 0; display: flex; flex-direction: column; height: 100vh; }
        .nav { background: #001529; color: white; padding: 12px 20px; display: flex; align-items: center; justify-content: space-between; }
        .main { display: flex; flex: 1; padding: 15px; gap: 15px; min-height: 0; }
        .panel { background: white; border-radius: 8px; flex: 1; display: flex; flex-direction: column; padding: 15px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
        #hot-container { flex: 1; overflow: hidden; width: 100%; }
        .log-box { flex: 1; background: #1e1e1e; color: #a9d18e; padding: 10px; font-family: monospace; font-size: 12px; overflow-y: auto; border-radius: 4px; }
        button { border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; font-weight: bold; margin-right: 8px; }
        .btn-run { background: #52c41a; color: white; }
        .btn-stop { background: #ff4d4f; color: white; }
        .btn-copy { background: #faad14; color: white; }
        .btn-save { background: #1890ff; color: white; }
    </style>
</head>
<body>
    <div class="nav">
        <div>
            <span style="font-size:18px; font-weight:bold; margin-right:20px;">VIP 采集 Pro 重构版</span>
            <button class="btn-run" onclick="ctl('start')">▶ 开始任务</button>
            <button class="btn-stop" onclick="ctl('stop')">⏹ 停止</button>
            <button class="btn-copy" onclick="copyToClipboard()">📋 复制(Excel格式)</button>
        </div>
        <div id="status-area">状态检测中...</div>
    </div>
    <div class="main">
        <div class="panel" style="flex: 2;">
            <div style="margin-bottom:10px; display:flex; justify-content:space-between; align-items:center;">
                <b>表格编辑器 (支持 Ctrl+V 粘贴 ID)</b>
                <button class="btn-save" onclick="saveTable()">💾 保存修改</button>
            </div>
            <div id="hot-container"></div>
        </div>
        <div class="panel">
            <b style="margin-bottom:10px;">控制台日志</b>
            <div id="logs" class="log-box"></div>
        </div>
    </div>
    <script>
        let hot;
        const container = document.getElementById('hot-container');
        async function sync() {
            try {
                const r = await fetch('/api/get_data');
                const res = await r.json();
                document.getElementById('logs').innerHTML = res.logs.join('<br>');
                document.getElementById('logs').scrollTop = 1000000;
                document.getElementById('status-area').innerText = res.is_running ? "● 正在抓取" : "○ 已停止";
                if(!hot) {
                    hot = new Handsontable(container, {
                        data: res.data, colHeaders: res.columns, rowHeaders: true,
                        width: '100%', height: '100%', licenseKey: 'non-commercial-and-evaluation',
                        stretchH: 'all', contextMenu: true
                    });
                } else if(res.is_running || hot.isListening() === false) {
                    hot.loadData(res.data);
                }
            } catch(e) {}
        }
        async function saveTable() {
            await fetch('/api/save_data', {
                method: 'POST', headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({columns: hot.getColHeader(), data: hot.getData()})
            });
            alert('数据已保存到 CSV');
        }
        async function ctl(cmd) { await fetch('/api/' + cmd, {method: 'POST'}); }
        function copyToClipboard() {
            const fullText = [hot.getColHeader().join('\\t'), ...hot.getData().map(r => r.join('\\t'))].join('\\n');
            const el = document.createElement('textarea'); el.value = fullText;
            document.body.appendChild(el); el.select(); document.execCommand('copy');
            document.body.removeChild(el); alert('复制成功，可直接粘贴到 Excel');
        }
        setInterval(sync, 3000);
        sync();
    </script>
</body>
</html>
"""

@app.route("/")
def index():
    return render_template_string(HTML_TPL)

@app.route("/api/get_data")
def api_get_data():
    init_csv_file()
    df = pd.read_csv(CONFIG.INPUT_FILE, dtype=str, encoding="utf-8-sig").fillna("")
    return jsonify({
        "columns": df.columns.tolist(),
        "data": df.values.tolist(),
        "is_running": state.is_running,
        "logs": state.logs
    })

@app.route("/api/save_data", methods=["POST"])
def api_save_data():
    data = request.json
    pd.DataFrame(data['data'], columns=data['columns']).to_csv(CONFIG.INPUT_FILE, index=False, encoding="utf-8-sig")
    return jsonify({"ok": True})

@app.route("/api/start", methods=["POST"])
def start():
    if not state.is_running:
        crawler = VIPCrawler()
        threading.Thread(target=crawler.run, daemon=True).start()
    return jsonify({"ok": True})

@app.route("/api/stop", methods=["POST"])
def stop():
    state.stop_flag = True
    return jsonify({"ok": True})

# ==============================================
# 启动程序
# ==============================================
if __name__ == "__main__":
    import logging
    logging.getLogger('werkzeug').disabled = True
    init_csv_file()
    app.run(host="0.0.0.0", port=5000)
