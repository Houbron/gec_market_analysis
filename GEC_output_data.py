import os
from datetime import datetime

import pandas as pd
from playwright.sync_api import sync_playwright

PAGE_URL = "https://gp.poweremarket.com/rept/sr/mp/portaladmin/login.html#/"
API_URLS = {
    "流出": "https://gp.poweremarket.com/rept/ma/gcc/gcc/publicHomePage/queryOutProvincialFlowDirection",
    "流入": "https://gp.poweremarket.com/rept/ma/gcc/gcc/publicHomePage/queryInProvincialFlowDirection",
}
API_URL = API_URLS["流出"]
TOP_N = 10
NAVIGATION_TIMEOUT_MS = 120000
ACTION_DELAY_MS = 1000
OUTPUT_EXCEL = "GEC_2026_绿证流入流出前10.xlsx"

# 登录配置：需要访问登录后数据时，将 LOGIN_REQUIRED 改为 True。
# 密码建议通过环境变量 GEC_PASSWORD 传入；没有设置时，由操作者在 Chrome 页面中手动输入。
# 图形验证码、短信验证码、登录按钮由操作者在打开的 Chrome 页面中手动完成。
LOGIN_REQUIRED = False
GEC_USERNAME = os.getenv("GEC_USERNAME", "18600192065")
GEC_PASSWORD = os.getenv("GEC_PASSWORD")
LOGIN_WAIT_TIMEOUT_MS = 300000
USERNAME_SELECTOR = ".login-part input[placeholder='请输入账号']"
PASSWORD_SELECTOR = ".login-part input[placeholder='请输入密码']"

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
    print("脚本会等待账号输入框消失，或等待你在终端按 Enter 确认已登录。")

    try:
        username_input.wait_for(state="hidden", timeout=LOGIN_WAIT_TIMEOUT_MS)
        print("检测到登录区域已消失，继续执行后续数据获取。")
    except Exception:
        input("未自动检测到登录完成。如果你已经登录成功，请按 Enter 继续；否则请先在浏览器中完成登录。")

    page.wait_for_load_state("networkidle", timeout=NAVIGATION_TIMEOUT_MS)
    page.wait_for_timeout(ACTION_DELAY_MS)

    print("已确认登录，本次运行将继续获取数据。")


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
    try:
        writer = pd.ExcelWriter(output_path, engine="openpyxl")
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = output_path.replace(".xlsx", f"_{timestamp}.xlsx")
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


if __name__ == "__main__":
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 300)

    year = "2026"
    flow_data = fetch_flow_data(year)

    out_df = build_table(flow_data["流出"], "流出", TOP_N)
    in_df = build_table(flow_data["流入"], "流入", TOP_N)

    print(f"\n{year} 年全国绿证流出量前 {TOP_N} 省份各月流出量及价格：")
    print(out_df.to_string(index=False))

    print(f"\n{year} 年全国绿证流入量前 {TOP_N} 省份各月流入量及价格：")
    print(in_df.to_string(index=False))

    output_path = save_to_excel(out_df, in_df)
    print(f"\nExcel 文件已生成：{output_path}")
