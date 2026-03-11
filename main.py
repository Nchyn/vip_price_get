import os
import time
import random
import sys
import pandas as pd
from playwright.sync_api import sync_playwright

# ==============================
# Windows 输出修复
# ==============================

if sys.platform.startswith('win'):
    import io
    import msvcrt
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)

def print_now(*args, **kwargs):
    kwargs['flush'] = True
    print(*args, **kwargs)

# ==============================
# 路径配置
# ==============================

def get_base_path():
    if hasattr(sys, '_MEIPASS'):
        return os.path.dirname(sys.executable)
    return os.path.abspath(".")

BASE_PATH = get_base_path()

INPUT_FILE = os.path.join(BASE_PATH, "商品清单.csv")
AUTH_FILE = os.path.join(BASE_PATH, "vip_auth.json")
CHROME_PATH = os.path.join(BASE_PATH, "chrome-win64", "chrome.exe")

# ==============================
# UA池
# ==============================

UA_POOL = [
"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36",
"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/121.0 Safari/537.36",
"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/120.0 Safari/537.36"
]

# ==============================
# 环境检查
# ==============================

def ensure_env():

    if not os.path.exists(INPUT_FILE):

        print_now(">>> 首次运行，创建 CSV 模板")

        df = pd.DataFrame(columns=["ID","品牌","标题","特卖价","原价","折扣"])

        df.to_csv(INPUT_FILE,index=False,encoding="utf-8-sig")

        print_now("模板创建成功，请填写 ID 列后重新运行")

        input()
        sys.exit()

    if not os.path.exists(CHROME_PATH):

        print_now("未找到浏览器内核")
        print_now(CHROME_PATH)

        input()
        sys.exit()

# ==============================
# 读取CSV
# ==============================

def load_product_list():

    df = pd.read_csv(INPUT_FILE,dtype=str)

    if "商品ID" in df.columns:
        df.rename(columns={"商品ID":"ID"},inplace=True)

    if "ID" not in df.columns:
        raise Exception("CSV必须包含 ID 或 商品ID")

    df=df.dropna(subset=["ID"])

    target_cols=["品牌","标题","特卖价","原价","折扣"]

    for col in target_cols:
        if col not in df.columns:
            df[col]=None

    return df

# ==============================
# 判断是否已抓取
# ==============================

def already_crawled(row):

    title=row.get("标题")
    price=row.get("特卖价")

    if pd.notna(title) and str(title).strip()!="":
        return True

    if pd.notna(price) and str(price).strip()!="":
        return True

    return False

# ==============================
# 浏览器指纹
# ==============================

def create_context(browser):

    ua=random.choice(UA_POOL)

    context=browser.new_context(
        user_agent=ua,
        viewport={
            "width":random.randint(1200,1920),
            "height":random.randint(700,1080)
        },
        locale="zh-CN",
        timezone_id="Asia/Shanghai"
    )

    context.add_init_script("""

Object.defineProperty(navigator,'webdriver',{get:()=>undefined})

""")

    return context

# ==============================
# 用户行为
# ==============================

def simulate_user(page):

    page.mouse.move(random.randint(100,1000),random.randint(100,700))
    page.mouse.wheel(0,random.randint(200,800))
    time.sleep(random.uniform(1,2))

# ==============================
# 倒计时
# ==============================

def wait_with_countdown(seconds):

    for i in range(seconds,0,-1):

        msg=f"等待 {i} 秒... (按任意键立即抓取)"

        sys.stdout.write("\r"+msg+" "*20)
        sys.stdout.flush()

        start=time.time()

        while time.time()-start<1:

            if sys.platform.startswith('win'):
                if msvcrt.kbhit():
                    msvcrt.getch()
                    sys.stdout.write("\r"+" "*80+"\r")
                    sys.stdout.flush()
                    return

            time.sleep(0.1)

    sys.stdout.write("\r"+" "*80+"\r")
    sys.stdout.flush()

# ==============================
# 登录
# ==============================

def manual_login(p):

    browser=p.chromium.launch(headless=False,executable_path=CHROME_PATH)

    context=browser.new_context()
    page=context.new_page()

    page.goto("https://passport.vip.com/login")

    input("登录完成后按回车")

    context.storage_state(path=AUTH_FILE)

    browser.close()

# ==============================
# 抓取
# ==============================

def fetch_vip_data(page,pid):

    url=f"https://detail.vip.com/detail-1234-{pid}.html"

    try:

        page.goto(url,timeout=15000)

        simulate_user(page)

        page.wait_for_selector(".sp-price",timeout=10000)

        brand="未知品牌"
        title="未知标题"

        try:
            brand=page.locator(".J_brandName").first.inner_text().strip()
        except:
            pass

        try:
            title=page.locator(".pib-title-detail").first.inner_text().strip()
        except:
            pass

        title=title[:40]

        sale=page.locator(".sp-price").first.inner_text().strip()

        try:
            market=page.locator(".marketPrice").first.inner_text().strip()
        except:
            market="N/A"

        try:
            disc=page.locator(".sp-discount").first.inner_text().strip()
        except:
            disc="N/A"

        return brand,title,sale,market,disc

    except:
        return None

# ==============================
# 主程序
# ==============================

def main():

    ensure_env()

    df=load_product_list()

    total=len(df)

    print_now(f"准备抓取 {total} 条")

    with sync_playwright() as p:

        if not os.path.exists(AUTH_FILE):
            manual_login(p)

        browser=p.chromium.launch(headless=False,executable_path=CHROME_PATH)

        context=browser.new_context(storage_state=AUTH_FILE)

        page=context.new_page()

        for index,row in df.iterrows():

            pid=str(row["ID"]).strip()

            prefix=f"[{index+1}/{total}] {pid}"

            if already_crawled(row):

                print_now(f"{prefix} | 已抓取，跳过")

                continue

            print_now(prefix,end=" ",flush=True)

            res=None

            for _ in range(3):

                res=fetch_vip_data(page,pid)

                if res:
                    break

            if res:

                brand,title,sale,market,disc=res

                df.at[index,"品牌"]=brand
                df.at[index,"标题"]=title
                df.at[index,"特卖价"]=sale
                df.at[index,"原价"]=market
                df.at[index,"折扣"]=disc

                print_now(f"| {brand} | {title} | {sale} | {market} | {disc}")

                df.to_csv(INPUT_FILE,index=False,encoding="utf-8-sig")

            else:

                print_now("| 抓取失败")

            wait=random.randint(60,120)

            wait_with_countdown(wait)

        browser.close()

    print_now("完成")

    input()

# ==============================

if __name__=="__main__":
    main()
