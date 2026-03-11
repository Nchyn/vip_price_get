import os
import time
import random
import sys
import subprocess
import pandas as pd
from playwright.sync_api import sync_playwright

# --- 兼容性设置：修复 Windows 乱码并解决输出延迟 ---
if sys.platform.startswith('win'):
    import io
    # 强制标准输出不缓冲，实时显示
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', line_buffering=True)

# 重新定义一个带自动刷新的 print 函数，确保每一行都即时显示
def print_now(*args, **kwargs):
    kwargs['flush'] = True
    print(*args, **kwargs)

# --- 路径配置区 ---
def get_base_path():
    if hasattr(sys, '_MEIPASS'):
        return os.path.dirname(sys.executable)
    return os.path.abspath(".")

BASE_PATH = get_base_path()
INPUT_FILE = os.path.join(BASE_PATH, "商品清单.xlsx")
AUTH_FILE = os.path.join(BASE_PATH, "vip_auth.json")
CHROME_PATH = os.path.join(BASE_PATH, "chrome-win64", "chrome.exe")

def ensure_env():
    """环境自检：检查Excel模板与本地浏览器"""
    if not os.path.exists(INPUT_FILE):
        print_now(f">>> 首次运行，正在为您创建模板: {INPUT_FILE}")
        df_init = pd.DataFrame(columns=['商品ID', '品牌', '标题', '特卖价', '原价', '折扣'])
        df_init.to_excel(INPUT_FILE, index=False)
        print_now("[OK] 模板创建成功，请在表格中填写【商品ID】后重新运行程序。")
        input("按回车退出..."); sys.exit()

    if not os.path.exists(CHROME_PATH):
        print_now(f"[!] 错误：未在预期路径找到浏览器内核！")
        print_now(f"预期路径: {CHROME_PATH}")
        print_now("请确保 'chrome-win64' 文件夹与本程序放在一起。")
        input("按回车退出..."); sys.exit()

def manual_login(p):
    print_now("\n" + "!"*40)
    print_now(" 未检测到登录状态，请在弹出的浏览器中完成登录！")
    print_now("!"*40 + "\n")
    
    browser = p.chromium.launch(headless=False, executable_path=CHROME_PATH)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://passport.vip.com/login")
    
    input(">>> 完成登录并进入首页后，请回到这里按回车[Enter]继续...")
    
    context.storage_state(path=AUTH_FILE)
    print_now(f"[OK] 登录状态已保存")
    browser.close()

def fetch_vip_data(page, product_id):
    url = f"https://detail.vip.com/detail-1234-{product_id}.html"
    try:
        page.goto(url, wait_until="commit", timeout=15000)
        page.wait_for_selector(".sp-price", timeout=10000)

        brand = "未知品牌"
        try:
            brand_el = page.locator(".pib-title-class.J_brandName").first
            if brand_el.is_visible():
                brand = brand_el.inner_text().strip()
        except: pass

        title = "未知标题"
        title_selectors = [".pib-title-detail", ".pms-product-title", ".detail-product-title", ".pro-title"]
        for selector in title_selectors:
            target = page.locator(selector).first
            if target.is_visible():
                title = target.inner_text().strip()
                break
        
        sale_price = page.locator(".sp-price").first.inner_text().strip()
        try:
            market_price = page.locator(".marketPrice").first.inner_text().strip()
        except: market_price = "N/A"
        try:
            discount = page.locator(".sp-discount").first.inner_text().strip()
        except: discount = "N/A"

        return brand, title, sale_price, market_price, discount
    except:
        return None

def main():
    ensure_env()

    df = pd.read_excel(INPUT_FILE, dtype={'商品ID': str})
    df = df.dropna(subset=['商品ID'])

    target_cols = ['品牌', '标题', '特卖价', '原价', '折扣']
    for col in target_cols:
        if col not in df.columns:
            df[col] = None
    df[target_cols] = df[target_cols].astype(object)

    print_now(f"[*] 准备处理 {len(df)} 条数据...")

    with sync_playwright() as p:
        if not os.path.exists(AUTH_FILE):
            manual_login(p)

        browser = p.chromium.launch(headless=False, executable_path=CHROME_PATH)
        try:
            context = browser.new_context(storage_state=AUTH_FILE)
            context.route("**/*.{png,jpg,jpeg,webp,gif}", lambda route: route.abort())
            page = context.new_page()

            for index, row in df.iterrows():
                pid = str(row['商品ID']).strip()
                
                if pd.notna(row.get('特卖价')) and str(row['特卖价']).strip() not in ["", "None", "nan", "N/A"]:
                    print_now(f"[{index+1}/{len(df)}] ID: {pid} 已跳过")
                    continue

                print_now(f"[{index+1}/{len(df)}] 正在查询 ID: {pid} ...", end="\r")
                res = fetch_vip_data(page, pid)
                
                if res:
                    brand, title, sale_p, market_p, disc = res
                    # 这里加一点空格覆盖掉上面那个 `正在查询`，防止长度不够出现错乱字符
                    print_now(f"[{index+1}/{len(df)}] ID: {pid} | {brand} | {title[:25]}...{' '*10}")
                    
                    df.at[index, '品牌'] = brand
                    df.at[index, '标题'] = title
                    df.at[index, '特卖价'] = sale_p
                    df.at[index, '原价'] = market_p
                    df.at[index, '折扣'] = disc
                    
                    # --- 修改点：每次成功获取后，立刻写入 Excel ---
                    try:
                        df.to_excel(INPUT_FILE, index=False)
                    except Exception as e:
                        print_now(f"\n[!] 写入Excel报错: 请检查 {INPUT_FILE} 是否正在被其他软件打开")
                else:
                    print_now(f"[{index+1}/{len(df)}] ID: {pid} [!] 抓取失败{' '*20}")

                if index < len(df) - 1:
                    wait_time = random.uniform(60, 120)
                    for i in range(int(wait_time), 0, -1):
                        # 冷却提示使用进度条风格或简单字符
                        print_now(f"    - 等待中... 剩余 {i} 秒  ", end="\r")
                        time.sleep(1)
                    print_now(" " * 40, end="\r")

        finally:
            try:
                # 保险起见，脚本退出时再整体保存一次
                df.to_excel(INPUT_FILE, index=False)
                print_now(f"\n[OK] 数据最终同步成功: {INPUT_FILE}")
            except:
                print_now(f"\n[!] 保存失败，请检查 Excel 是否被占用")
            browser.close()

    input("\n程序运行完毕，按回车键关闭...")

if __name__ == "__main__":
    main()
