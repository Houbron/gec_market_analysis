import json
import os
import random
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright

try:
    from local_config import GEC_PASSWORD as LOCAL_GEC_PASSWORD
    from local_config import GEC_USERNAME as LOCAL_GEC_USERNAME
except ImportError:
    LOCAL_GEC_USERNAME = None
    LOCAL_GEC_PASSWORD = None

PAGE_URL = "https://gp.poweremarket.com/rept/sr/mp/portaladmin/login.html#/"
LISTED_PRODUCTS_PAGE_URL = "https://gp.poweremarket.com/rept/sr/dmz/gcc/greenSignProducts.html#/GreenSignProducts"
API_URLS = {
    "流出": "https://gp.poweremarket.com/rept/ma/gcc/gcc/publicHomePage/queryOutProvincialFlowDirection",
    "流入": "https://gp.poweremarket.com/rept/ma/gcc/gcc/publicHomePage/queryInProvincialFlowDirection",
}
API_URL = API_URLS["流出"]
TOP_N = 10
NAVIGATION_TIMEOUT_MS = 300000
ACTION_DELAY_MS = 1000
PAGE_TURN_DELAY_MIN_MS = 6000
PAGE_TURN_DELAY_MAX_MS = 12000
RATE_LIMIT_COOLDOWN_MS = 120000
RATE_LIMIT_TEXT = "访问过于频繁"
OUTPUT_EXCEL = "GEC绿证市场情况.xlsx"
LISTED_PRODUCTS_EXCEL = "GEC_绿证挂牌产品.xlsx"
DESKTOP_OUTPUT_DIR = Path.home() / "Desktop"
CACHE_DIR = Path("cache")
LISTED_PRODUCTS_CACHE = CACHE_DIR / "listed_products_cache.json"
EXPORT_HOME_FLOW_DATA = True
EXPORT_LISTED_PRODUCTS = True
USE_DIRECT_LISTED_PRODUCTS_PAGE = True
INCREMENTAL_LISTED_PRODUCTS = True
LISTED_PRODUCTS_CACHE_ANCHOR_COUNT = 100
LISTED_PRODUCTS_STOP_MATCH_COUNT = 20
MODE_FLOW_ONLY = "1"
MODE_FULL_LISTED_PRODUCTS = "2"
MODE_LIGHT_LISTED_PRODUCTS = "3"

# 登录配置：需要访问登录后数据时，将 LOGIN_REQUIRED 改为 True。
# 账号密码读取顺序：环境变量 -> local_config.py -> 默认账号/手动输入。
# local_config.py 已加入 .gitignore，可用于本机自动填充密码，避免提交到 GitHub。
# 图形验证码、短信验证码、登录按钮由操作者在打开的 Chrome 页面中手动完成。
LOGIN_REQUIRED = False
GEC_USERNAME = os.getenv("GEC_USERNAME") or LOCAL_GEC_USERNAME or "18600192065"
GEC_PASSWORD = os.getenv("GEC_PASSWORD") or LOCAL_GEC_PASSWORD
LOGIN_WAIT_TIMEOUT_MS = 300000
USERNAME_SELECTOR = ".login-part input[placeholder='请输入账号']"
PASSWORD_SELECTOR = ".login-part input[placeholder='请输入密码']"
PRODUCT_LIST_WAIT_MS = 3000
PRODUCT_LIST_LOAD_TIMEOUT_MS = 120000
PAGE_CHANGE_TIMEOUT_MS = 60000
LISTED_PRODUCTS_PAGE_SIZE = 48
MAX_PRODUCT_PAGES = 300

# 自定义省份映射
PROVINCE_MAP = {
    "BJ": "北京",
    "TJ": "天津",
    "HEB": "河北",
    "SX": "山西",
    "NM": "内蒙古",
    "LN": "辽宁",
    "JL": "吉林",
    "HL": "黑龙江",
    "SH": "上海",
    "JS": "江苏",
    "ZJ": "浙江",
    "AH": "安徽",
    "FJ": "福建",
    "JX": "江西",
    "SD": "山东",
    "HEN": "河南",
    "HB": "湖北",
    "HN": "湖南",
    "02": "广东",
    "03": "广西",
    "06": "海南",
    "CQ": "重庆",
    "SC": "四川",
    "05": "贵州",
    "04": "云南",
    "XZ": "西藏",
    "SN": "陕西",
    "GS": "甘肃",
    "QH": "青海",
    "NX": "宁夏",
    "XJ": "新疆",
    "TW": "台湾",
    "HK": "香港",
    "MO": "澳门"
}

def _safe_float(value, default=0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _timestamped_excel_path(output_path):
    output_path = Path(output_path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = DESKTOP_OUTPUT_DIR if DESKTOP_OUTPUT_DIR.exists() else Path.cwd()
    file_stem = output_path.stem
    file_suffix = output_path.suffix or ".xlsx"
    return str(output_dir / f"{file_stem}_{timestamp}{file_suffix}")


def _sort_key(item):
    ranking = item.get("rankingLevel")
    if ranking:
        return (0, int(ranking), -_safe_float(item.get("quantity")))
    return (1, 999, -_safe_float(item.get("quantity")))


def _create_context(browser):
    return browser.new_context()


def login_if_needed(context, page):
    page.goto(PAGE_URL, timeout=NAVIGATION_TIMEOUT_MS)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(ACTION_DELAY_MS)

    username_input = page.locator(USERNAME_SELECTOR).first
    if not username_input.is_visible(timeout=10000):
        try:
            print("未检测到账号输入框，尝试点击“账号登录”入口...")
            page.get_by_text("账号登录", exact=True).click(timeout=10000)
            page.wait_for_timeout(ACTION_DELAY_MS)
        except Exception:
            pass

    if not username_input.is_visible(timeout=10000):
        print("仍未检测到账号输入框，可能已经处于登录态或登录入口未展开。")
        return

    print("正在填写账号...")
    username_input.fill(GEC_USERNAME, timeout=60000)
    page.wait_for_timeout(ACTION_DELAY_MS)

    if GEC_PASSWORD:
        print("检测到环境变量 GEC_PASSWORD，正在填写密码...")
        page.locator(PASSWORD_SELECTOR).first.fill(GEC_PASSWORD, timeout=60000)
        page.wait_for_timeout(ACTION_DELAY_MS)
    else:
        print("未设置 GEC_PASSWORD，请在 Chrome 页面中手动输入密码。")

    print("请在打开的 Chrome 页面中输入密码、图形验证码，获取并输入短信验证码，然后点击登录。")
    print("脚本会等待账号输入框消失或页面进入登录后的 greenTradeHome。")

    deadline = datetime.now().timestamp() + LOGIN_WAIT_TIMEOUT_MS / 1000
    login_confirmed = False
    while datetime.now().timestamp() < deadline:
        if page.is_closed():
            raise RuntimeError("登录等待期间浏览器页面被关闭，已停止后续抓取。")

        try:
            if "greenTradeHome.html" in page.url:
                login_confirmed = True
                break

            if not username_input.is_visible(timeout=1000):
                login_confirmed = True
                break
        except Exception:
            login_confirmed = True
            break

        page.wait_for_timeout(ACTION_DELAY_MS)

    if not login_confirmed:
        raise RuntimeError("等待登录完成超时，请重新运行脚本并在浏览器中完成登录。")

    print("检测到登录完成，继续执行后续数据获取。")

    page.wait_for_load_state("networkidle", timeout=NAVIGATION_TIMEOUT_MS)
    page.wait_for_timeout(ACTION_DELAY_MS)

    print("已确认登录，本次运行将继续获取数据。")


def _extract_value(text, label):
    match = re.search(rf"{label}\s*[：:]\s*(.+)", text)
    if not match:
        return ""

    value = match.group(1).strip()
    return re.split(
        r"\s+(?:电源类型|电力生产年月|补贴类型|消纳方式|项目地址|过期时间|已出售|可交易|单价)[：:]",
        value,
    )[0].strip()


def _extract_owner(text):
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    skip_prefixes = (
        "电源类型",
        "电力生产年月",
        "补贴类型",
        "消纳方式",
        "项目地址",
        "过期时间",
        "可交易",
        "已出售",
        "单价",
        "立即购买",
    )
    for line in lines:
        if not line.startswith(skip_prefixes):
            return line
    return ""


def _extract_product_cards(page):
    card_items = []
    for frame in page.frames:
        try:
            frame_items = frame.evaluate(
                """() => {
                    const cards = Array.from(document.querySelectorAll(
                        ".table-wrapper .el-card, .tabel-wrapper .el-card"
                    ));
                    return cards
                        .map(card => {
                            const title = card.querySelector(".title");
                            const rows = Array.from(card.querySelectorAll("p.text-list"));
                            if (!title || rows.length === 0) {
                                return null;
                            }

                            const fields = {};
                            for (const row of rows) {
                                const labelNode = row.querySelector(".list-title");
                                if (!labelNode) {
                                    continue;
                                }

                                const label = (labelNode.innerText || "")
                                    .replace(/[：:]/g, "")
                                    .trim();
                                const values = Array.from(row.querySelectorAll(".list-content"))
                                    .map(node => (node.innerText || "").trim())
                                    .filter(Boolean);
                                fields[label] = values.join(" ");
                            }

                            const stats = {};
                            for (const item of Array.from(card.querySelectorAll(".data-list-wrapper .data-item"))) {
                                const titleNode = item.querySelector(".data-item-title");
                                const contentNode = item.querySelector(".data-item-content");
                                if (!titleNode || !contentNode) {
                                    continue;
                                }
                                stats[(titleNode.innerText || "").trim()] = (contentNode.innerText || "").trim();
                            }

                            return {
                                title: (title.innerText || "").trim(),
                                fields,
                                stats
                            };
                        })
                        .filter(Boolean);
                }"""
            )
            card_items.extend(frame_items)
        except Exception:
            continue

    products = []
    for item in card_items:
        fields = item.get("fields") or {}
        stats = item.get("stats") or {}
        products.append(
            {
                "项目业主": item.get("title", "").strip().strip('"').strip(),
                "电源类型": fields.get("电源类型", ""),
                "电力生产年月": fields.get("电力生产年月", ""),
                "补贴类型": fields.get("补贴类型", ""),
                "消纳方式": fields.get("消纳方式", ""),
                "可交易数量": stats.get("可交易", ""),
                "已出售数量": stats.get("已出售", ""),
                "单价": stats.get("单价", ""),
            }
        )
    return products


def _product_key(product):
    return tuple(
        str(product.get(column, "")).strip()
        for column in ["项目业主", "电源类型", "电力生产年月", "补贴类型", "消纳方式", "单价"]
    )


def _product_key_text(product):
    return "||".join(_product_key(product))


def load_listed_products_cache():
    if not LISTED_PRODUCTS_CACHE.exists():
        return {"products": [], "head_keys": []}

    try:
        with open(LISTED_PRODUCTS_CACHE, "r", encoding="utf-8") as file:
            return json.load(file)
    except Exception as error:
        print(f"读取挂牌产品缓存失败，将按首次运行处理：{error}")
        return {"products": [], "head_keys": []}


def save_listed_products_cache(products):
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    product_records = products.to_dict("records") if isinstance(products, pd.DataFrame) else products
    cache_data = {
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "head_keys": [_product_key_text(product) for product in product_records[:LISTED_PRODUCTS_CACHE_ANCHOR_COUNT]],
        "products": product_records,
    }

    with open(LISTED_PRODUCTS_CACHE, "w", encoding="utf-8") as file:
        json.dump(cache_data, file, ensure_ascii=False, indent=2)


def merge_with_listed_products_cache(current_df, cache_data):
    if current_df.empty or not current_df.attrs.get("stopped_by_cache_anchor"):
        return current_df

    cached_products = cache_data.get("products") or []
    if not cached_products:
        return current_df

    current_keys = set(current_df.apply(lambda row: _product_key_text(row), axis=1))
    cached_tail = [
        product
        for product in cached_products
        if _product_key_text(product) not in current_keys
    ]

    if not cached_tail:
        return current_df

    print(f"增量扫描命中缓存锚点，追加缓存中的旧产品 {len(cached_tail)} 条。")
    cached_tail_df = pd.DataFrame(cached_tail)
    return pd.concat([current_df, cached_tail_df], ignore_index=True)


def _get_first_product_key(page):
    products = _extract_product_cards(page)
    if not products:
        return None
    return _product_key(products[0])


def _get_active_page_number(page):
    for frame in page.frames:
        try:
            active = frame.locator(".el-pager li.active").first
            if active.count() == 0:
                continue
            return active.inner_text(timeout=3000).strip()
        except Exception:
            continue
    return ""


def wait_for_product_cards(page):
    print("正在等待挂牌产品卡片加载...")
    deadline = datetime.now().timestamp() + PRODUCT_LIST_LOAD_TIMEOUT_MS / 1000

    while datetime.now().timestamp() < deadline:
        for frame in page.frames:
            try:
                card_count = frame.locator(".table-wrapper .el-card, .tabel-wrapper .el-card").count()
                field_count = frame.locator(".list-title").count()
                if card_count > 0 and field_count > 0:
                    print(f"检测到挂牌产品卡片：{card_count} 个")
                    return True
            except Exception:
                continue

        page.wait_for_timeout(ACTION_DELAY_MS)

    print("等待挂牌产品卡片超时，继续尝试解析当前页面。")
    return False


def human_pause(min_ms=PAGE_TURN_DELAY_MIN_MS, max_ms=PAGE_TURN_DELAY_MAX_MS):
    delay_ms = random.randint(min_ms, max_ms)
    print(f"等待 {delay_ms}ms，降低访问频率...")
    # 用标准 sleep，避免页面最小化时浏览器定时器被节流影响等待节奏。
    time.sleep(delay_ms / 1000)


def handle_rate_limit_notice(page):
    try:
        notice = page.get_by_text(RATE_LIMIT_TEXT)
        if notice.count() == 0:
            return False

        if notice.first.is_visible(timeout=1000):
            print(f"检测到“{RATE_LIMIT_TEXT}”提示，暂停 {RATE_LIMIT_COOLDOWN_MS}ms 后继续。")
            human_pause(RATE_LIMIT_COOLDOWN_MS, RATE_LIMIT_COOLDOWN_MS)
            return True
    except Exception:
        return False

    return False


def set_listed_products_page_size(page, page_size=LISTED_PRODUCTS_PAGE_SIZE):
    if not page_size:
        return False

    option_text = f"{page_size}条/页"
    print(f"正在设置挂牌产品每页显示数量：{option_text}")

    for frame in page.frames:
        try:
            size_selector = frame.locator(".el-pagination__sizes .el-select").first
            if size_selector.count() == 0:
                continue

            size_selector.click(timeout=60000)
            page.wait_for_timeout(ACTION_DELAY_MS)
            page.locator(".el-select-dropdown__item", has_text=option_text).last.click(timeout=60000)
            page.wait_for_timeout(PRODUCT_LIST_WAIT_MS)
            wait_for_product_cards(page)
            return True
        except Exception:
            continue

    print("未能设置每页显示数量，将使用页面默认分页。")
    return False


def _go_to_next_product_page(page):
    next_button = page.locator(".btn-next").last
    if next_button.count() == 0:
        return False

    try:
        before_key = _get_first_product_key(page)
        before_page_number = _get_active_page_number(page)

        classes = next_button.get_attribute("class", timeout=10000) or ""
        disabled = next_button.get_attribute("disabled", timeout=10000)
        if "disabled" in classes or disabled is not None:
            return False

        next_button.click(timeout=60000)
        human_pause()
        deadline = datetime.now().timestamp() + PAGE_CHANGE_TIMEOUT_MS / 1000

        while datetime.now().timestamp() < deadline:
            handle_rate_limit_notice(page)
            page.wait_for_timeout(ACTION_DELAY_MS)
            after_key = _get_first_product_key(page)
            after_page_number = _get_active_page_number(page)

            if after_key and before_key and after_key != before_key:
                return True

            if after_page_number and before_page_number and after_page_number != before_page_number:
                return True

        print("点击下一页后页面内容未变化，停止继续翻页。")
        return False
    except Exception as error:
        print(f"翻页失败或已无下一页：{error}")
        return False


def navigate_to_listed_products(page):
    print("正在进入“我的绿证 -> 绿证挂牌产品”页面...")
    page.get_by_text("我的绿证", exact=True).hover(timeout=60000)
    page.wait_for_timeout(ACTION_DELAY_MS)
    page.get_by_text("绿证挂牌产品", exact=True).click(timeout=60000)
    page.wait_for_load_state("networkidle", timeout=NAVIGATION_TIMEOUT_MS)
    page.wait_for_timeout(PRODUCT_LIST_WAIT_MS)
    wait_for_product_cards(page)
    set_listed_products_page_size(page)


def go_to_direct_listed_products_page(page):
    print("正在直接进入绿证挂牌产品子页面...")
    page.goto(
        LISTED_PRODUCTS_PAGE_URL,
        wait_until="domcontentloaded",
        timeout=NAVIGATION_TIMEOUT_MS,
    )
    page.wait_for_timeout(PRODUCT_LIST_WAIT_MS)
    wait_for_product_cards(page)
    set_listed_products_page_size(page)


def fetch_listed_products():
    cache_data = load_listed_products_cache()
    cache_head_keys = set(cache_data.get("head_keys") or [])
    matched_cache_head_keys = set()
    stopped_by_cache_anchor = False

    if INCREMENTAL_LISTED_PRODUCTS and cache_head_keys:
        print(
            f"检测到挂牌产品缓存，将以上次前 {len(cache_head_keys)} 个产品为锚点，"
            f"命中任意 {LISTED_PRODUCTS_STOP_MATCH_COUNT} 个后停止继续翻页。"
        )

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=60)
        context = _create_context(browser)
        page = context.new_page()
        page.set_default_navigation_timeout(NAVIGATION_TIMEOUT_MS)
        page.set_default_timeout(60000)

        login_if_needed(context, page)
        if USE_DIRECT_LISTED_PRODUCTS_PAGE:
            go_to_direct_listed_products_page(page)
        else:
            navigate_to_listed_products(page)

        all_products = []
        seen = set()

        for page_number in range(1, MAX_PRODUCT_PAGES + 1):
            products = _extract_product_cards(page)
            print(f"第 {page_number} 页识别到 {len(products)} 个挂牌产品")

            for product in products:
                key = tuple(product.values())
                if key not in seen:
                    seen.add(key)
                    all_products.append(product)

                if INCREMENTAL_LISTED_PRODUCTS and cache_head_keys:
                    key_text = _product_key_text(product)
                    if key_text in cache_head_keys:
                        matched_cache_head_keys.add(key_text)

            if (
                INCREMENTAL_LISTED_PRODUCTS
                and cache_head_keys
                and len(matched_cache_head_keys) >= LISTED_PRODUCTS_STOP_MATCH_COUNT
            ):
                stopped_by_cache_anchor = True
                print(
                    f"已命中缓存锚点 {len(matched_cache_head_keys)} 个，停止继续翻页。"
                )
                break

            if not _go_to_next_product_page(page):
                break

        browser.close()
        products_df = pd.DataFrame(all_products)
        products_df.attrs["stopped_by_cache_anchor"] = stopped_by_cache_anchor
        products_df.attrs["cache_data"] = cache_data
        return products_df


def scrape_listed_products_from_page(page, incremental=True, reset_cache=False):
    cache_data = {"products": [], "head_keys": []} if reset_cache else load_listed_products_cache()
    cache_head_keys = set(cache_data.get("head_keys") or [])
    matched_cache_head_keys = set()
    stopped_by_cache_anchor = False

    if incremental and cache_head_keys:
        print(
            f"检测到挂牌产品缓存，将以上次前 {len(cache_head_keys)} 个产品为锚点，"
            f"命中任意 {LISTED_PRODUCTS_STOP_MATCH_COUNT} 个后停止继续翻页。"
        )
    elif reset_cache:
        print("全量模式：忽略旧缓存，遍历全部挂牌产品并重建缓存。")

    if USE_DIRECT_LISTED_PRODUCTS_PAGE:
        go_to_direct_listed_products_page(page)
    else:
        navigate_to_listed_products(page)

    all_products = []
    seen = set()

    for page_number in range(1, MAX_PRODUCT_PAGES + 1):
        products = _extract_product_cards(page)
        print(f"第 {page_number} 页识别到 {len(products)} 个挂牌产品")

        for product in products:
            key = tuple(product.values())
            if key not in seen:
                seen.add(key)
                all_products.append(product)

            if incremental and cache_head_keys:
                key_text = _product_key_text(product)
                if key_text in cache_head_keys:
                    matched_cache_head_keys.add(key_text)

        if (
            incremental
            and cache_head_keys
            and len(matched_cache_head_keys) >= LISTED_PRODUCTS_STOP_MATCH_COUNT
        ):
            stopped_by_cache_anchor = True
            print(f"已命中缓存锚点 {len(matched_cache_head_keys)} 个，停止继续翻页。")
            break

        if not _go_to_next_product_page(page):
            break

    products_df = pd.DataFrame(all_products)
    products_df.attrs["stopped_by_cache_anchor"] = stopped_by_cache_anchor
    products_df.attrs["cache_data"] = cache_data
    return products_df


def prepare_listed_products_for_export(products_df, mode=MODE_LIGHT_LISTED_PRODUCTS):
    cache_data = products_df.attrs.get("cache_data") or load_listed_products_cache()
    cached_products = cache_data.get("products") or []
    cached_keys = {_product_key_text(product) for product in cached_products}

    if mode == MODE_LIGHT_LISTED_PRODUCTS:
        scanned_df = products_df.copy()
        if scanned_df.empty:
            current_df = pd.DataFrame(cached_products)
            if not current_df.empty:
                current_df["状态"] = "上次抓取结果"
        else:
            scanned_df["状态"] = scanned_df.apply(
                lambda row: "上次抓取结果"
                if _product_key_text(row) in cached_keys
                else "本次新增",
                axis=1,
            )

            if products_df.attrs.get("stopped_by_cache_anchor") and cached_products:
                scanned_keys = set(scanned_df.apply(lambda row: _product_key_text(row), axis=1))
                cached_tail = [
                    product
                    for product in cached_products
                    if _product_key_text(product) not in scanned_keys
                ]
                if cached_tail:
                    print(f"增量扫描命中缓存锚点，追加缓存中的旧产品 {len(cached_tail)} 条。")
                    cached_tail_df = pd.DataFrame(cached_tail)
                    cached_tail_df["状态"] = "上次抓取结果"
                    current_df = pd.concat([scanned_df, cached_tail_df], ignore_index=True)
                else:
                    current_df = scanned_df
            else:
                current_df = scanned_df
    else:
        current_df = products_df.copy()
        current_df["状态"] = "当前挂牌"

    current_df["抓取时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if current_df.empty:
        print("挂牌产品数据为空，本次不更新缓存。")
    else:
        save_listed_products_cache(current_df)
        print(f"挂牌产品缓存已更新：{LISTED_PRODUCTS_CACHE}")
    return current_df


def save_listed_products_to_excel(products_df, output_path=LISTED_PRODUCTS_EXCEL, mode=MODE_LIGHT_LISTED_PRODUCTS):
    output_path = _timestamped_excel_path(output_path)
    current_df = prepare_listed_products_for_export(products_df, mode=mode)

    try:
        writer = pd.ExcelWriter(output_path, engine="openpyxl")
    except PermissionError:
        output_path = _timestamped_excel_path(output_path)
        print(f"原 Excel 文件可能正在打开，改为保存到：{output_path}")
        writer = pd.ExcelWriter(output_path, engine="openpyxl")

    with writer:
        current_df.to_excel(writer, sheet_name="当前挂牌", index=False)

        for worksheet in writer.sheets.values():
            worksheet.freeze_panes = "A2"
            for column_cells in worksheet.columns:
                max_length = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in column_cells
                )
                worksheet.column_dimensions[column_cells[0].column_letter].width = min(
                    max(max_length + 2, 12),
                    40,
                )

    return output_path


# 该函数使用 Playwright 模拟浏览器行为，访问指定页面并发送 POST 请求获取原始数据
def fetch_raw_data(year="2026", direction="流出", context=None):
    api_url = API_URLS[direction]

    if context:
        response = context.request.post(api_url, data={"years": year})
        return response.json()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=60)
        context = _create_context(browser)
        page = context.new_page()
        page.set_default_navigation_timeout(NAVIGATION_TIMEOUT_MS)

        if LOGIN_REQUIRED:
            login_if_needed(context, page)
        else:
            page.goto(PAGE_URL)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(ACTION_DELAY_MS)

        response = context.request.post(
            api_url,
            data={"years": year}
        )
        page.wait_for_timeout(ACTION_DELAY_MS)

        data = response.json()
        browser.close()
        return data


def fetch_flow_data(year="2026"):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=60)
        context = _create_context(browser)
        page = context.new_page()
        page.set_default_navigation_timeout(NAVIGATION_TIMEOUT_MS)

        if LOGIN_REQUIRED:
            login_if_needed(context, page)
        else:
            page.goto(PAGE_URL)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(ACTION_DELAY_MS)

        result = {}
        for direction in API_URLS:
            print(f"正在获取 {year} 年绿证{direction}数据...")
            result[direction] = fetch_raw_data(year, direction, context=context)
            page.wait_for_timeout(ACTION_DELAY_MS)

        browser.close()
        return result


# 该函数将原始数据转换为 pandas DataFrame 格式，方便后续分析和展示
def build_table(data, direction="流出", top_n=TOP_N):
    if not data or not data.get("data"):
        print(f"{direction}接口数据为空")
        return pd.DataFrame()

    provinces = [
        item for item in data["data"]
        if item and item.get("provinceCode") != "00"
    ]

    provinces = sorted(provinces, key=_sort_key)
    top_items = provinces[:top_n]

    rows = []

    for index, item in enumerate(top_items, start=1):
        province_code = item.get("provinceCode")
        province_name = PROVINCE_MAP.get(province_code, province_code)

        quantity_map = {
            q["month"]: q["quotaValue"]
            for q in (item.get("flowQuantity") or [])
        }

        price_map = {
            p["month"]: p["quotaValue"]
            for p in (item.get("flowPrice") or [])
        }

        row = {
            "排名": int(item.get("rankingLevel") or index),
            "省份": province_name,
            "年度累计量": _safe_float(item.get("quantity")),
        }

        # 按月份顺序添加电量和电价数据
        for month in sorted(quantity_map.keys()):
            month_label = month.split("-")[-1]
            row[f"{month_label}月数量"] = quantity_map.get(month)
            row[f"{month_label}月平均价格"] = price_map.get(month)

        rows.append(row)

    df = pd.DataFrame(rows)
    return df


def save_to_excel(out_df, in_df, output_path=OUTPUT_EXCEL):
    output_path = _timestamped_excel_path(output_path)
    try:
        writer = pd.ExcelWriter(output_path, engine="openpyxl")
    except PermissionError:
        output_path = _timestamped_excel_path(output_path)
        print(f"原 Excel 文件可能正在打开，改为保存到：{output_path}")
        writer = pd.ExcelWriter(output_path, engine="openpyxl")

    with writer:
        out_df.to_excel(writer, sheet_name="流出前10", index=False)
        in_df.to_excel(writer, sheet_name="流入前10", index=False)

        for sheet_name in ["流出前10", "流入前10"]:
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes = "A2"

            for column_cells in worksheet.columns:
                max_length = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in column_cells
                )
                worksheet.column_dimensions[column_cells[0].column_letter].width = min(
                    max(max_length + 2, 10),
                    20,
                )

                header = column_cells[0].value
                if header == "年度累计量" or header.endswith("月数量") or header.endswith("月平均价格"):
                    for cell in column_cells[1:]:
                        cell.number_format = "#,##0.00"

    return output_path


def save_market_report_to_excel(out_df, in_df, products_df, output_path=OUTPUT_EXCEL, mode=MODE_LIGHT_LISTED_PRODUCTS):
    output_path = _timestamped_excel_path(output_path)
    products_export_df = prepare_listed_products_for_export(products_df, mode=mode)

    try:
        writer = pd.ExcelWriter(output_path, engine="openpyxl")
    except PermissionError:
        output_path = _timestamped_excel_path(output_path)
        print(f"原 Excel 文件可能正在打开，改为保存到：{output_path}")
        writer = pd.ExcelWriter(output_path, engine="openpyxl")

    with writer:
        out_df.to_excel(writer, sheet_name="流出前10", index=False)
        in_df.to_excel(writer, sheet_name="流入前10", index=False)
        products_export_df.to_excel(writer, sheet_name="挂牌产品", index=False)

        for sheet_name, worksheet in writer.sheets.items():
            worksheet.freeze_panes = "A2"
            for column_cells in worksheet.columns:
                max_length = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in column_cells
                )
                max_width = 40 if sheet_name == "挂牌产品" else 20
                min_width = 12 if sheet_name == "挂牌产品" else 10
                worksheet.column_dimensions[column_cells[0].column_letter].width = min(
                    max(max_length + 2, min_width),
                    max_width,
                )

                header = column_cells[0].value
                if header == "年度累计量" or str(header).endswith("月数量") or str(header).endswith("月平均价格"):
                    for cell in column_cells[1:]:
                        cell.number_format = "#,##0.00"

    return output_path


def fetch_market_report_data(year="2026", include_listed_products=False, incremental=True, reset_cache=False):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=60)
        context = _create_context(browser)
        page = context.new_page()
        page.set_default_navigation_timeout(NAVIGATION_TIMEOUT_MS)
        page.set_default_timeout(60000)

        if include_listed_products or LOGIN_REQUIRED:
            login_if_needed(context, page)
        else:
            page.goto(PAGE_URL)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(ACTION_DELAY_MS)

        flow_data = {}
        if EXPORT_HOME_FLOW_DATA:
            for direction in API_URLS:
                print(f"正在获取 {year} 年绿证{direction}数据...")
                flow_data[direction] = fetch_raw_data(year, direction, context=context)
                page.wait_for_timeout(ACTION_DELAY_MS)

        products_df = pd.DataFrame()
        if include_listed_products:
            products_df = scrape_listed_products_from_page(
                page,
                incremental=incremental,
                reset_cache=reset_cache,
            )

        browser.close()
        return flow_data, products_df


def choose_run_mode():
    if len(sys.argv) > 1:
        choice = sys.argv[1].strip()
        if choice in {MODE_FLOW_ONLY, MODE_FULL_LISTED_PRODUCTS, MODE_LIGHT_LISTED_PRODUCTS}:
            return choice

    print("\n请选择运行模式：")
    print("1. 仅获取流入/流出前十统计")
    print("2. 获取当前全部挂牌绿证统计")
    print("3. 轻量化获取")

    while True:
        try:
            choice = input("请输入选项编号（1/2/3，默认 3）：").strip() or MODE_LIGHT_LISTED_PRODUCTS
        except EOFError:
            print("当前运行环境无法交互输入，自动使用默认选项 3。")
            return MODE_LIGHT_LISTED_PRODUCTS

        if choice in {MODE_FLOW_ONLY, MODE_FULL_LISTED_PRODUCTS, MODE_LIGHT_LISTED_PRODUCTS}:
            return choice
        print("输入无效，请输入 1、2 或 3。")


if __name__ == "__main__":
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 300)

    year = "2026"
    run_mode = choose_run_mode()
    include_listed_products = run_mode in {MODE_FULL_LISTED_PRODUCTS, MODE_LIGHT_LISTED_PRODUCTS}
    incremental = run_mode == MODE_LIGHT_LISTED_PRODUCTS
    reset_cache = run_mode == MODE_FULL_LISTED_PRODUCTS

    flow_data, products_df = fetch_market_report_data(
        year,
        include_listed_products=include_listed_products,
        incremental=incremental,
        reset_cache=reset_cache,
    )

    if EXPORT_HOME_FLOW_DATA:
        out_df = build_table(flow_data["流出"], "流出", TOP_N)
        in_df = build_table(flow_data["流入"], "流入", TOP_N)

        print(f"\n{year} 年全国绿证流出量前 {TOP_N} 省份各月流出量及价格：")
        print(out_df.to_string(index=False))

        print(f"\n{year} 年全国绿证流入量前 {TOP_N} 省份各月流入量及价格：")
        print(in_df.to_string(index=False))

        if include_listed_products:
            output_path = save_market_report_to_excel(
                out_df,
                in_df,
                products_df,
                mode=run_mode,
            )
        else:
            output_path = save_to_excel(out_df, in_df)
        print(f"\nExcel 文件已生成：{output_path}")
