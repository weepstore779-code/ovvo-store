# -*- coding: utf-8 -*-
"""OVVO storehouse — نظام إدارة المخزون والمستودعات"""
import base64
import hashlib
import hmac
import html as htmllib
import json
import os
import re
import secrets
import shutil
from datetime import datetime
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl.utils import get_column_letter
from PIL import Image, ImageOps

# =========================
# إعدادات الصفحة
# =========================
st.set_page_config(page_title="OVVO storehouse", page_icon="📦", layout="wide")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_NAME = os.path.join(BASE_DIR, "OVVO.xlsx")
ORDERS_FILE = os.path.join(BASE_DIR, "OVVO_orders.xlsx")
BANNER_PATH = os.path.join(BASE_DIR, "assets", "ovvo_inventory_banner.png")
USERS_FILE = os.path.join(BASE_DIR, "ovvo_users.json")
BACKUP_DIR = os.path.join(BASE_DIR, "backups")
IMAGES_DIR = os.path.join(BASE_DIR, "assets", "products")
MAX_BACKUPS = 30
SHEET_NAME = "المخزون"

COLUMNS = [
    "الباركود التلقائي", "وصف", "وصف 2", "الكلفة النهائية", "سعر البيع",
    "سعر البيع 2", "الكمية", "ضريبة", "نوع الصنف", "الماركة", "المنشاة",
    "الموسم", "سنة الشراء", "الطلبية", "محجب او غير محجب",
    "كت او كم او نص كم", "فاتورة رقم", "اللون", "كود اللون تلقائي",
    "المقاس", "كود المقاس تلقائي", "كود النوع تلقائي", "رقم الموديل",
    "كود الصنف",
]
NUM_COLUMNS = ["الكلفة النهائية", "سعر البيع", "سعر البيع 2", "الكمية"]
TEXT_COLUMNS = [c for c in COLUMNS if c not in NUM_COLUMNS]
HIDDEN_COLUMNS = ["الكلفة النهائية"]

PRIORITY_COLUMNS = ["الباركود التلقائي", "وصف", "وصف 2", "اللون",
                    "المقاس", "الكمية", "سعر البيع", "ضريبة"]
VISIBLE_COLUMNS = PRIORITY_COLUMNS + [
    c for c in COLUMNS if c not in HIDDEN_COLUMNS and c not in PRIORITY_COLUMNS
]

SEARCH_COLUMNS = ["الباركود التلقائي", "وصف", "وصف 2", "نوع الصنف", "الماركة",
                  "رقم الموديل", "اللون", "المقاس", "كود الصنف"]
ROLE_LABELS = {"admin": "مدير", "viewer": "مشاهد"}
STATUS_OPTIONS = ["الكل", "متوفر", "منخفض", "نفد"]
AVAILABLE_NOT_OUT = "المتوفر (عدا النافد)"

# ---- التوصيل ----
ORDER_COLS = ["رقم الطلب", "التاريخ", "الباركود", "الوصف", "اللون", "المقاس",
              "الكمية", "سعر البيع", "المبلغ المطلوب", "الحالة", "أُعيد للمخزون"]
ST_OUT = "قيد التوصيل"
ST_DELIVERED = "تم التوصيل"
ST_RET_WAIT = "مرتجع - بانتظار الاستلام"
ST_CAN_WAIT = "ملغي - بانتظار الاستلام"
ST_RET_DONE = "مرتجع - تم الاستلام"
ST_CAN_DONE = "ملغي - تم الاستلام"
OPEN_STATUSES = [ST_OUT, ST_DELIVERED, ST_RET_WAIT, ST_CAN_WAIT]

# =========================
# الأكواد الثابتة (لون / مقاس / نوع)
# =========================
CODES_FILE = os.path.join(BASE_DIR, "ovvo_codes.json")  # الأكواد الجديدة اللي بتتولد لحالها

DEFAULT_COLORS = {
    "اسود": "01", "ابيض": "02", "احمر": "03", "ازرق": "04", "كحلي": "05",
    "سماوي": "06", "اخضر": "07", "زيتي": "08", "بيج": "09", "بني": "10",
    "رمادي": "11", "زهري": "12", "وردي": "13", "بنفسجي": "14", "اصفر": "15",
    "برتقالي": "16", "عنابي": "17", "ذهبي": "18", "فضي": "19", "كريمي": "20",
    "موف": "21", "جينز": "24", "ليموني": "28",
}
DEFAULT_TYPES = {
    "بلوزة": "101", "قميص": "202", "فستان": "303", "بنطلون": "404", "جينز": "505",
    "شورت": "606", "طقم بنطلون": "707", "طقم ترينج": "808", "طقم تنورة": "909",
    "جاكيت": "1001", "فيست": "2002", "بليزر": "3003", "احذية": "4004",
    "شنتة": "5005", "جزدان": "6006",
}
SIZES = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL",
         "34", "36", "38", "40", "42", "44", "46", "48", "50", "52", "56", "58", "60",
         "فري سايز"]
FREE_SIZE_CODE = "99"
FREE_ALIASES = {"free", "free size", "freesize", "فري سايز", "فري"}


def norm_key(s):
    """توحيد الكتابة العربية (أ/ا، ة/ه، ى/ي، بنطالون/بنطلون، حذف ال التعريف) للمقارنة فقط."""
    s = re.sub(r"[أإآ]", "ا", str(s).strip())
    s = s.replace("ة", "ه").replace("ى", "ي").replace("بنطالون", "بنطلون")
    s = re.sub(r"\s+", " ", s).lower()
    if s.startswith("ال") and len(s) > 4:
        s = s[2:]
    return s


def load_extra_codes():
    try:
        with open(CODES_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
    except Exception:
        d = {}
    d.setdefault("colors", {})
    d.setdefault("types", {})
    return d


def save_extra_codes(d):
    with open(CODES_FILE, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)


def known_names(kind, extra_names=()):
    """أسماء الألوان أو الأنواع المعروفة (الثابتة + الجديدة + أي اسم إضافي)."""
    base = DEFAULT_COLORS if kind == "colors" else DEFAULT_TYPES
    names = list({**base, **load_extra_codes()[kind]})
    seen = {norm_key(n) for n in names}
    for n in extra_names:
        if n and norm_key(n) not in seen:
            names.append(n)
            seen.add(norm_key(n))
    return names


def get_code(kind, name):
    """يرجّع الكود الثابت. إذا الاسم جديد بياخد كود جديد ويتثبت للأبد."""
    name = str(name).strip()
    if not name:
        return ""
    base = DEFAULT_COLORS if kind == "colors" else DEFAULT_TYPES
    extra = load_extra_codes()
    table = {**base, **extra[kind]}
    key = norm_key(name)
    for n, c in table.items():
        if norm_key(n) == key:
            return c
    biggest = max(int(c) for c in table.values())
    new = f"{biggest + 1:02d}" if kind == "colors" else str(biggest + 1001)
    extra[kind][name] = new
    save_extra_codes(extra)
    return new


def size_code(size):
    s = str(size).strip()
    if not s:
        return ""
    if norm_key(s) in FREE_ALIASES:
        return FREE_SIZE_CODE
    return re.sub(r"\s+", "", s.upper())  # المقاس نفسه هو الكود


def next_model(data):
    """رقم موديل جديد من 3 أرقام ما بتكرر."""
    nums = pd.to_numeric(data["رقم الموديل"], errors="coerce").dropna()
    nums = nums[(nums >= 0) & (nums <= 998)]
    return f"{int(nums.max()) + 1:03d}" if len(nums) else "001"


def build_barcode(model, color_c, size_c, type_c):
    return f"{model}{color_c}{size_c}{type_c}"


@st.cache_resource(show_spinner=False)
def load_banner():
    if os.path.exists(BANNER_PATH):
        try:
            with open(BANNER_PATH, "rb") as f:
                return base64.b64encode(f.read()).decode("utf-8")
        except Exception:
            return ""
    return ""


image_base64 = load_banner()

# =========================
# التصميم (CSS)
# =========================
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800&family=Marcellus&display=swap');

:root {
    --navy: #0D1B2E; --navy2: #14263F;
    --maroon: #7A1229; --maroon2: #55091B;
    --orange: #E8830C;
    --bg: #FAF7F3; --card: #FFFFFF; --border: #EEE7DE;
    --text: #1E2633; --muted: #6B7280;
    --dark: #0D1B2E; --beige: #E8D9C3; --cream: #FAF7F3;
    --shadow: 0 6px 22px rgba(13,27,46,.07);
}

.stApp, .stApp :is(p, label, li, h1, h2, h3, h4, input, textarea, button, a,
    [data-testid="stMetricValue"], [data-testid="stMetricLabel"]) {
    font-family: 'Cairo', sans-serif;
}
.stApp { background: var(--bg); color: var(--text); }
[data-testid="stAppViewContainer"], [data-testid="stSidebar"] {
    direction: rtl; text-align: right;
}
[data-testid="stHeader"] { background: transparent; }
#MainMenu, footer { visibility: hidden; }
.block-container { padding-top: 1.2rem; max-width: 1300px; }
h1, h2, h3, h4 { color: var(--navy) !important; font-weight: 800 !important; }

/* ---------- الشريط الجانبي ---------- */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, var(--navy) 0%, #08121F 100%);
    position: relative;
}
[data-testid="stSidebar"]::after {
    content: ""; position: absolute; left: 0; right: 0; bottom: 0; height: 170px;
    pointer-events: none; z-index: 0;
    background:
        linear-gradient(158deg, transparent 52%, var(--orange) 52%, var(--orange) 60%, transparent 60%),
        linear-gradient(162deg, transparent 38%, var(--maroon2) 38%);
}
[data-testid="stSidebarContent"] { position: relative; z-index: 1; }
[data-testid="stSidebar"] * { color: #F3EDE4; }
[data-testid="stSidebar"] [role="radiogroup"] { gap: 6px; }
[data-testid="stSidebar"] [role="radiogroup"] label {
    padding: 11px 14px; border-radius: 12px; width: 100%; margin: 0; transition: .15s;
}
[data-testid="stSidebar"] [role="radiogroup"] label > div:first-child { display: none; }
[data-testid="stSidebar"] [role="radiogroup"] label p { font-weight: 600; font-size: 15.5px; }
[data-testid="stSidebar"] [role="radiogroup"] label:hover { background: rgba(255,255,255,.07); }
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {
    background: linear-gradient(90deg, var(--maroon), var(--maroon2));
    box-shadow: 0 8px 18px rgba(122,18,41,.45);
}
[data-testid="stSidebar"] .stButton > button {
    background: transparent; border: 1px solid rgba(232,217,195,.55); color: #F3EDE4;
}
[data-testid="stSidebar"] .stButton > button:hover {
    background: var(--orange); border-color: var(--orange); color: #fff;
}
.ovvo-brand {
    background: var(--navy); border-radius: 16px; padding: 18px 12px;
    margin-bottom: 12px; text-align: center;
}
[data-testid="stSidebar"] .ovvo-brand { background: transparent; padding-bottom: 6px; }
.ovvo-brand img { width: 150px; height: auto; max-width: 80%; }
.ovvo-brand-name {
    font-family: 'Marcellus', serif; font-size: 30px; color: #F3EDE4;
    letter-spacing: 1px; line-height: 1.2; margin-top: 4px; direction: ltr;
}
.ovvo-brand-name span { letter-spacing: 5px; font-size: 21px; margin-left: 6px; }
.user-chip {
    background: rgba(255,255,255,.06); border-radius: 12px;
    padding: 9px 12px; margin-bottom: 12px; font-size: 14px;
}
.tagline {
    text-align: center; margin: 26px 0 14px; letter-spacing: 4px; font-size: 12.5px;
    color: #E8D9C3 !important; line-height: 1.9; direction: ltr;
}
.tagline div { width: 38px; height: 2px; background: var(--orange); margin: 8px auto 0; }

/* ---------- الشريط العلوي ---------- */
.topbar {
    display: flex; align-items: center; justify-content: space-between; direction: ltr;
    background: #fff; border: 1px solid var(--border); border-radius: 18px;
    padding: 12px 24px; box-shadow: var(--shadow); margin-bottom: 22px;
}
.tb-brand { display: flex; align-items: center; gap: 12px; }
.tb-brand img { height: 56px; width: auto; }
.tb-name { font-family: 'Marcellus', serif; font-size: 34px; color: var(--navy); line-height: 1; }
.tb-name span { letter-spacing: 6px; font-size: 22px; color: #555; margin-left: 10px; }
.tb-right { direction: rtl; text-align: right; }
.tb-title { font-weight: 800; font-size: 18px; color: var(--navy); }
.tb-sub { font-size: 12.5px; color: var(--muted); }

/* ---------- عناوين الصفحات ---------- */
.page-head { margin-bottom: 18px; border-right: 5px solid var(--maroon); padding-right: 14px; }
.page-head h1 { margin: 0; font-size: 28px; }
.page-head p { margin: 4px 0 0; color: var(--muted); font-size: 15px; }

/* ---------- الرئيسية ---------- */
.hello { padding: 14px 6px; }
.hello-s { font-size: 24px; font-weight: 700; color: var(--navy); }
.hello-b { font-family: 'Marcellus', serif; font-size: 50px; color: var(--maroon); line-height: 1.15; direction: ltr; text-align: right; }
.hello-d { color: var(--muted); font-size: 15px; margin-top: 6px; }
.banner { position: relative; border-radius: 18px; overflow: hidden; box-shadow: var(--shadow); background: var(--navy); }
.banner img { width: 100%; height: auto; display: block; }
.banner-fallback { width: 100%; height: 215px; background: linear-gradient(120deg, var(--navy), var(--maroon2)); }
.banner .chip {
    position: absolute; left: 18px; top: 50%; transform: translateY(-50%);
    background: rgba(13,27,46,.93); color: #fff; border-radius: 14px; padding: 16px 22px;
}
.chip-t { font-weight: 800; font-size: 18px; color: #fff; }
.chip-s { font-size: 13.5px; color: #E8D9C3; }

.kpis { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin: 20px 0; }
.kpi {
    background: #fff; border: 1px solid var(--border); border-radius: 16px; padding: 18px;
    box-shadow: var(--shadow); display: flex; align-items: center; gap: 15px;
}
.kpi-ic { width: 58px; height: 58px; border-radius: 50%; display: flex; align-items: center;
    justify-content: center; flex-shrink: 0; }
.kpi-ic svg { width: 28px; height: 28px; stroke: #fff; fill: none; stroke-width: 1.8;
    stroke-linecap: round; stroke-linejoin: round; }
.kpi-l { color: var(--muted); font-size: 14px; }
.kpi-v { font-size: 30px; font-weight: 800; color: var(--navy); line-height: 1.15; }
.kpi-s { font-size: 12.5px; color: var(--muted); }
@media (max-width: 900px) { .kpis { grid-template-columns: repeat(2, 1fr); } }

.panel {
    background: #fff; border: 1px solid var(--border); border-radius: 18px;
    padding: 18px 20px; box-shadow: var(--shadow); margin-bottom: 14px;
}
.panel-h { display: flex; align-items: center; gap: 10px; font-weight: 800; font-size: 18px;
    color: var(--navy); margin-bottom: 14px; }
.panel-h .count { background: var(--maroon); color: #fff; border-radius: 999px;
    font-size: 12.5px; padding: 1px 11px; }
.ok-note { color: #1E7B3A; font-weight: 700; padding: 10px 0; }
.foot-note { color: var(--muted); font-size: 12.5px; margin-top: 10px; }
.tbl-wrap { overflow-x: auto; }
.tbl { width: 100%; border-collapse: separate; border-spacing: 0; }
.tbl th { background: #F6F1EA; color: #444; font-weight: 700; font-size: 13.5px;
    padding: 10px 12px; text-align: right; white-space: nowrap; }
.tbl th:first-child { border-radius: 0 10px 10px 0; }
.tbl th:last-child { border-radius: 10px 0 0 10px; }
.tbl td { padding: 10px 12px; border-bottom: 1px solid #F1ECE5; font-size: 14px; color: #333;
    vertical-align: middle; }
.tbl tr:last-child td { border-bottom: none; }
.pname { display: flex; align-items: center; gap: 11px; }
.pn { font-weight: 700; color: var(--navy); }
.pm { font-size: 12px; color: var(--muted); }
.thumb { width: 42px; height: 42px; border-radius: 10px; object-fit: cover; background: #F1ECE5;
    flex-shrink: 0; }
.thumb.ph { display: flex; align-items: center; justify-content: center; font-weight: 800;
    color: var(--maroon); font-size: 18px; }
.mono { direction: ltr; text-align: right; font-family: 'Courier New', monospace; color: #555; }
.pill { display: inline-block; padding: 3px 16px; border-radius: 999px; font-size: 12.5px;
    font-weight: 700; color: #fff; }
.pill.ok { background: #2E8B4A; } .pill.warn { background: #EF9A1F; } .pill.bad { background: #D8453B; }
.mini-row { display: flex; justify-content: space-between; align-items: center;
    padding: 11px 0; border-bottom: 1px dashed #EADFD1; font-size: 14.5px; }
.mini-row:last-child { border-bottom: none; }
.mini-row b { color: var(--maroon); font-size: 17px; }

/* ---------- العناصر العامة ---------- */
[data-testid="stMetric"] {
    background: #fff; border: 1px solid var(--border); border-radius: 14px;
    padding: 14px 16px; box-shadow: var(--shadow);
}
[data-testid="stMetricLabel"] { color: var(--muted); }
[data-testid="stMetricValue"] { color: var(--navy); font-weight: 800; }

.stButton > button, .stDownloadButton > button,
[data-testid="stFormSubmitButton"] > button {
    border-radius: 12px; font-weight: 700; border: 1px solid var(--navy); background: #fff;
    color: var(--navy); transition: .15s;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    border-color: var(--maroon); color: var(--maroon);
}
button[kind="primary"], [data-testid="stBaseButton-primary"],
[data-testid="stBaseButton-primaryFormSubmit"] {
    background: linear-gradient(90deg, var(--maroon), var(--maroon2)) !important;
    color: #fff !important; border-color: var(--maroon2) !important;
    box-shadow: 0 6px 14px rgba(122,18,41,.28);
}
button[kind="primary"]:hover, [data-testid="stBaseButton-primary"]:hover,
[data-testid="stBaseButton-primaryFormSubmit"]:hover { filter: brightness(1.12); color: #fff !important; }
[data-baseweb="input"], [data-baseweb="select"] > div, textarea {
    border-radius: 12px !important; background: #fff !important;
}
input, textarea { text-align: right; }
[data-testid="stForm"] {
    background: #fff; border: 1px solid var(--border); border-radius: 18px;
    box-shadow: var(--shadow);
}
[data-testid="stExpander"] { background: #fff; border: 1px solid var(--border); border-radius: 14px; }
[data-testid="stDataFrame"] {
    border: 1px solid var(--border); border-radius: 14px; overflow: hidden;
    box-shadow: var(--shadow);
}
[data-testid="stAlert"] { border-radius: 14px; }
.stTabs [data-baseweb="tab"] { font-weight: 700; }
.stTabs [aria-selected="true"] { color: var(--maroon) !important; }
.stTabs [data-baseweb="tab-highlight"] { background: var(--maroon) !important; }

@media print {
    [data-testid="stSidebar"], [data-testid="stHeader"], .stButton,
    .stDownloadButton, iframe, .topbar { display: none !important; }
    .stApp { background: #fff; }
}
</style>
""",
    unsafe_allow_html=True,
)


def flash(msg, kind="success"):
    st.session_state["_flash"] = (kind, msg)


def show_flash():
    item = st.session_state.pop("_flash", None)
    if item:
        kind, msg = item
        {"success": st.success, "error": st.error,
         "warning": st.warning, "info": st.info}[kind](msg)


def page_header(title, subtitle=""):
    sub = f"<p>{subtitle}</p>" if subtitle else ""
    st.markdown(f'<div class="page-head"><h1>{title}</h1>{sub}</div>',
                unsafe_allow_html=True)


def to_excel_bytes(data, sheet="تقرير"):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name=sheet)
        ws = writer.sheets[sheet]
        ws.sheet_view.rightToLeft = True
        for i, col in enumerate(data.columns, 1):
            lengths = [len(str(col))] + [len(str(v)) for v in data[col].head(200)]
            ws.column_dimensions[get_column_letter(i)].width = min(max(lengths) + 3, 40)
    return buf.getvalue()


def download_buttons(data, base_name, key):
    stamp = datetime.now().strftime("%Y-%m-%d")
    c1, c2, _ = st.columns([1, 1, 3])
    c1.download_button(
        "⬇️ تنزيل Excel", to_excel_bytes(data), f"{base_name}_{stamp}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"{key}_xlsx", use_container_width=True,
    )
    c2.download_button(
        "⬇️ تنزيل CSV", data.to_csv(index=False).encode("utf-8-sig"),
        f"{base_name}_{stamp}.csv", "text/csv",
        key=f"{key}_csv", use_container_width=True,
    )


def hash_password(password, salt=None):
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), bytes.fromhex(salt), 310000
    ).hex()
    return {"salt": salt, "password_hash": digest}


def verify_password(password, saved):
    try:
        calculated = hash_password(password, saved["salt"])["password_hash"]
        return hmac.compare_digest(calculated, saved["password_hash"])
    except Exception:
        return False


def load_users():
    if not os.path.exists(USERS_FILE):
        return {}
    try:
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_users(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)


# شعار OvvO Store (مأخوذ من التصميم، خلفية شفافة)
LOGO_TOP_B64 = "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"
LOGO_SIDE_B64 = "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+Utc/4c/HG/7/SPXrlzZnJycVP1+H71eDyKCn/3gB0etWrUKT33qU3/Vm5xc2Ov1kocMyFrril5ZVufKdM3IyAgUewADHtDliSwzAeXFLok3ZZbAGANrrY+mmSeHhoYGSq3e8mVZdk15wsMwHP3Upz71yTRN0WjU0O8DRx75eBz7jGOkPFkvf/nLccUVV2DD6tUHtdvt07LMgOAvyvXXX0/XXXdT8Z18L9zrXnc2xsbGMDExwS972cteaYzBd77zHWoMDXWGhoYwOTn532XAURYIG40GlFJV8EdF6myw2CAicGJ9WdlPHav4ED7boxEEAdavG4EiHoNzMFkGW5zHQcORZRlK90EpBeMsvvaNr9MTnvREev7zn3/Xueeei5NPPhk33XSTPPOZz5TxkZHdS1pts9kEnGuFYYhGowGbpkeWsUur1aqyAnvssQduueUWAECSJGi1WidkWeZV8IMAExMT32Vm3HDDDTjwwANRazTGC77JHGNMFSNJlh04Pj5+V3n9y+8AAENDQxgeHsazn/vce056/vOvee6pp9767Gc/25x44oly3HHHyRlnnCHPOumka5YvX479998fcK71YHgyerZgr3QrSktcXpTbb7+9cg1IRWi324iiaFk/y1db5zBv3jxMdpIqF0mlLJRSVR5ZMaM1Z85tnljj36s9rLF27VqsX7/+u2WazTmHjRs3nv3tb3/77//2pX+FpO8Q1RinnXYKfvi97zxt7wMff8yTnvQknHPOOQgbjRu6nd4NgQ59DrOfwuQWX/rSl3DYYe9DnguCgHDEoYdh0aJF+4+Nj9/6jGc8A1/96lfR7/cxf9G8w5Je/7rJyd5d9913HxYvPgjdboZ6I8Suu+6KLMswb96ivbq9/l3G0TRrinIGRzFK2FqLPEmWzhtqj80dnoPhoSZ6vRxhFODmm28GEdmtlWZLILRaLfR6PTSbzXiiM5nUajWEYYhbbrll77vvvAtfuPDCJf9yzjlrLrnkEhxy2KHzO53OijiOfaUxiv7UaDSQpinCev3aVqs1d3R8bLNSCvMWzD9t08jGSxuNBjZs2AAAGB4eXrpy5cqfzZs3r3L52u22T3NOTvpdaXJyl+W777776Ojor51zqlxwEInKAlq/31eNRsN2u10VNxorFi9ejD322APz589HnucoUm0YnxhDo9HAyMgIGo0G7r77blx9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"


def logo_html(variant="top"):
    data = LOGO_SIDE_B64 if variant == "side" else LOGO_TOP_B64
    return f'<img src="data:image/png;base64,{data}" alt="OvvO Store">'


def brand_block():
    st.markdown('<div class="ovvo-brand">' + logo_html("side") + '</div>',
                unsafe_allow_html=True)


def top_bar():
    st.markdown(
        '<div class="topbar"><div class="tb-brand">' + logo_html("top") + '</div>'
        '<div class="tb-right"><div class="tb-title">المستودع الإلكتروني</div>'
        '<div class="tb-sub">إدارة المخزون عبر الويب</div></div></div>',
        unsafe_allow_html=True,
    )


ICONS = {
    "box": '<svg viewBox="0 0 24 24"><path d="M21 8l-9-5-9 5 9 5 9-5z"/><path d="M3 8v8l9 5 9-5V8"/><path d="M12 13v8"/></svg>',
    "list": '<svg viewBox="0 0 24 24"><rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 4h6v3H9z"/><path d="M9 12h6M9 16h4"/></svg>',
    "warn": '<svg viewBox="0 0 24 24"><path d="M12 3l10 18H2L12 3z"/><path d="M12 10v5M12 18v.5"/></svg>',
    "stop": '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M5.6 5.6l12.8 12.8"/></svg>',
}


def kpi_card(icon, label, value, sub, color):
    return (f'<div class="kpi"><div class="kpi-ic" style="background:{color}">{ICONS[icon]}</div>'
            f'<div><div class="kpi-l">{label}</div><div class="kpi-v">{value}</div>'
            f'<div class="kpi-s">{sub}</div></div></div>')


@st.cache_data(show_spinner=False)
def thumb_b64(path, mtime):
    try:
        img = Image.open(path).convert("RGB")
        img.thumbnail((96, 96))
        buf = BytesIO()
        img.save(buf, "JPEG", quality=80)
        return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()
    except Exception:
        return ""


STATUS_STYLE = {
    "متوفر": "background-color:#DFF3E4;color:#1E7B3A;font-weight:700",
    "منخفض": "background-color:#FDEBCB;color:#B26A00;font-weight:700",
    "نفد": "background-color:#F9D9D9;color:#B3261E;font-weight:700",
}


def style_status(data):
    """تلوين عمود «الحالة» في الجداول (أخضر / برتقالي / أحمر)."""
    if "الحالة" not in data.columns:
        return data
    try:
        sty = data.style
        fn = sty.map if hasattr(sty, "map") else sty.applymap
        return fn(lambda v: STATUS_STYLE.get(v, ""), subset=["الحالة"])
    except Exception:
        return data


def stock_table_html(data, limit=8):
    """جدول أنيق بصور مصغّرة وشارات حالة ملوّنة (للصفحة الرئيسية)."""
    esc = htmllib.escape
    cls = {"متوفر": "ok", "منخفض": "warn", "نفد": "bad"}
    rows = []
    for _, r in data.head(limit).iterrows():
        p = get_model_image(r["رقم الموديل"])
        uri = thumb_b64(p, os.path.getmtime(p)) if p else ""
        desc = str(r["وصف"]) or "—"
        thumb = (f'<img class="thumb" src="{uri}" alt="">' if uri
                 else f'<div class="thumb ph">{esc(desc[:1])}</div>')
        meta = " / ".join(x for x in [str(r["اللون"]), str(r["المقاس"])] if x)
        model = str(r["رقم الموديل"]) or "—"
        rows.append(
            '<tr><td><div class="pname">' + thumb
            + '<div><div class="pn">' + esc(desc) + '</div><div class="pm">' + esc(meta)
            + '</div></div></div></td><td class="mono">' + esc(str(r["الباركود التلقائي"]))
            + '</td><td><b>' + f'{float(r["الكمية"]):,.0f}' + '</b></td><td>' + esc(model)
            + '</td><td><span class="pill ' + cls.get(r["الحالة"], "ok") + '">'
            + esc(str(r["الحالة"])) + '</span></td></tr>'
        )
    head = ('<tr><th>المنتج</th><th>الرمز</th><th>الكمية</th><th>الموديل</th>'
            '<th>الحالة</th></tr>')
    return ('<div class="tbl-wrap"><table class="tbl"><thead>' + head + '</thead><tbody>'
            + "".join(rows) + '</tbody></table></div>')


def go_delivery():
    st.session_state["nav"] = "🚚 التوصيل"


def login_screen():
    users = load_users()
    _, mid, _ = st.columns([1, 1.4, 1])
    with mid:
        brand_block()
        if not users:
            st.subheader("الإعداد الأول")
            st.caption("أنشئ حساب المدير وحساب المشاهدة. المشاهد لا يعدّل البيانات.")
            with st.form("first_setup"):
                admin_name = st.text_input("اسم مستخدم المدير")
                admin_pass = st.text_input("كلمة مرور المدير (6 أحرف على الأقل)", type="password")
                viewer_name = st.text_input("اسم مستخدم المشاهدة")
                viewer_pass = st.text_input("كلمة مرور المشاهدة (6 أحرف على الأقل)", type="password")
                submitted = st.form_submit_button("إنشاء الحسابات", type="primary",
                                                  use_container_width=True)
            if submitted:
                a, v = admin_name.strip(), viewer_name.strip()
                if not a or not v:
                    st.error("أدخل اسم مستخدم للمدير وللمشاهد.")
                elif a == v:
                    st.error("يجب أن يختلف اسم المدير عن اسم المشاهد.")
                elif len(admin_pass) < 6 or len(viewer_pass) < 6:
                    st.error("كلمة المرور يجب ألا تقل عن 6 أحرف.")
                else:
                    hp_a = hash_password(admin_pass)
                    hp_v = hash_password(viewer_pass)
                    save_users({
                        a: {"salt": hp_a["salt"], "password_hash": hp_a["password_hash"], "role": "admin"},
                        v: {"salt": hp_v["salt"], "password_hash": hp_v["password_hash"], "role": "viewer"},
                    })
                    st.rerun()
            st.stop()

        st.subheader("تسجيل الدخول")
        with st.form("login_form"):
            username = st.text_input("اسم المستخدم")
            password = st.text_input("كلمة المرور", type="password")
            submitted = st.form_submit_button("دخول", type="primary", use_container_width=True)
        if submitted:
            user = users.get(username.strip())
            if user and verify_password(password, user):
                st.session_state["ovvo_username"] = username.strip()
                st.session_state["ovvo_role"] = user["role"]
                st.rerun()
            else:
                st.error("اسم المستخدم أو كلمة المرور غير صحيحة.")
        st.stop()


if "ovvo_username" not in st.session_state:
    login_screen()

current_username = st.session_state["ovvo_username"]
current_role = st.session_state["ovvo_role"]

if current_username not in load_users():
    st.session_state.clear()
    st.rerun()


def normalize(data):
    data = data.copy()
    for c in COLUMNS:
        if c not in data.columns:
            data[c] = ""
    data = data[COLUMNS]
    for c in NUM_COLUMNS:
        data[c] = pd.to_numeric(data[c], errors="coerce")
    data["الكمية"] = data["الكمية"].fillna(0)
    for c in TEXT_COLUMNS:
        data[c] = (
            data[c].fillna("").astype(str).str.strip()
            .str.replace(r"^(\d+)\.0$", r"\1", regex=True)
        )
    return data


@st.cache_data(show_spinner=False)
def _read_excel(path, mtime):
    if not os.path.isfile(path):
        return pd.DataFrame(columns=COLUMNS), ""
    try:
        raw = pd.read_excel(path, sheet_name=SHEET_NAME, dtype={
            "الباركود التلقائي": str, "رقم الموديل": str, "كود اللون تلقائي": str,
            "كود المقاس تلقائي": str, "كود النوع تلقائي": str, "كود الصنف": str,
        })
        raw = raw.dropna(how="all").reset_index(drop=True)
        return normalize(raw), ""
    except Exception as e:
        return pd.DataFrame(columns=COLUMNS), f"تعذّرت قراءة ورقة «{SHEET_NAME}» من الملف: {e}"


def load_data():
    mtime = os.path.getmtime(FILE_NAME) if os.path.isfile(FILE_NAME) else 0
    data, error = _read_excel(FILE_NAME, mtime)
    return data.copy(), error


def save_data(data):
    data = normalize(data)
    tmp = os.path.join(BASE_DIR, "_OVVO_tmp.xlsx")
    try:
        if os.path.isfile(FILE_NAME):
            os.makedirs(BACKUP_DIR, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            shutil.copy2(FILE_NAME, os.path.join(BACKUP_DIR, f"OVVO_{stamp}.xlsx"))
            old = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith("OVVO_"))
            for f in old[:-MAX_BACKUPS]:
                os.remove(os.path.join(BACKUP_DIR, f))
        data.to_excel(tmp, sheet_name=SHEET_NAME, index=False)
        os.replace(tmp, FILE_NAME)
    except PermissionError:
        st.error("تعذّر الحفظ: ملف OVVO.xlsx مفتوح في Excel. أغلقه ثم أعد المحاولة.")
        return False
    except Exception as e:
        st.error(f"تعذّر الحفظ: {e}")
        return False
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
    st.cache_data.clear()
    return True


def next_barcode(data):
    nums = pd.to_numeric(data["الباركود التلقائي"], errors="coerce").dropna()
    return str(int(nums.max()) + 1) if len(nums) else "1001"


def qty(data):
    return pd.to_numeric(data["الكمية"], errors="coerce").fillna(0)


# ---- صور المنتجات (لكل رقم موديل) ----
SIZE_ORDER = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL", "3XL", "4XL",
              "FREE", "FREE SIZE", "فري سايز"]


def size_key(s):
    u = str(s).strip().upper()
    if u in SIZE_ORDER:
        return (0, SIZE_ORDER.index(u), "")
    try:
        return (1, float(u), "")
    except ValueError:
        return (2, 0, u)


def model_image_path(model):
    model = str(model).strip()
    if not model:
        return None
    safe = re.sub(r"[^\w\-]", "_", model)
    return os.path.join(IMAGES_DIR, f"{safe}.jpg")


def get_model_image(model):
    p = model_image_path(model)
    return p if p and os.path.isfile(p) else None


def save_model_image(model, uploaded):
    p = model_image_path(model)
    if not p or uploaded is None:
        return False
    try:
        os.makedirs(IMAGES_DIR, exist_ok=True)
        img = ImageOps.exif_transpose(Image.open(uploaded))
        if img.mode in ("RGBA", "LA", "P"):
            img = img.convert("RGBA")
            bg = Image.new("RGB", img.size, (255, 255, 255))
            bg.paste(img, mask=img.split()[-1])
            img = bg
        else:
            img = img.convert("RGB")
        img.thumbnail((900, 900))
        img.save(p, "JPEG", quality=88)
        return True
    except Exception as e:
        st.error(f"تعذّر حفظ الصورة: {e}")
        return False


# ---- الباركود (Code 128 بدون أي مكتبات خارجية) ----
CODE128_PATTERNS = [
    "212222", "222122", "222221", "121223", "121322", "131222", "122213", "122312",
    "132212", "221213", "221312", "231212", "112232", "122132", "122231", "113222",
    "123122", "123221", "223211", "221132", "221231", "213212", "223112", "312131",
    "311222", "321122", "321221", "312212", "322112", "322211", "212123", "212321",
    "232121", "111323", "131123", "131321", "112313", "132113", "132311", "211313",
    "231113", "231311", "112133", "112331", "132131", "113123", "113321", "133121",
    "313121", "211331", "231131", "213113", "213311", "213131", "311123", "311321",
    "331121", "312113", "312311", "332111", "314111", "221411", "431111", "111224",
    "111422", "121124", "121421", "141122", "141221", "112214", "112412", "122114",
    "122411", "142112", "142211", "241211", "221114", "413111", "241112", "134111",
    "111242", "121142", "121241", "114212", "124112", "124211", "411212", "421112",
    "421211", "212141", "214121", "412121", "111143", "111341", "131141", "114113",
    "114311", "411113", "411311", "113141", "114131", "311141", "411131", "211412",
    "211214", "211232", "2331112",
]


def code128_widths(text):
    vals = [104]  # Start B
    for ch in text:
        o = ord(ch)
        if not 32 <= o <= 126:
            raise ValueError("unsupported character")
        vals.append(o - 32)
    chk = (104 + sum(v * i for i, v in enumerate(vals[1:], 1))) % 103
    vals += [chk, 106]
    return [int(c) for c in "".join(CODE128_PATTERNS[v] for v in vals)]


def barcode_svg(text, vertical=False):
    try:
        widths = code128_widths(text)
    except ValueError:
        return ""
    quiet, x, rects = 10, 10, []
    for i, w in enumerate(widths):
        if i % 2 == 0:
            if vertical:
                rects.append(f'<rect x="0" y="{x}" width="1" height="{w}"/>')
            else:
                rects.append(f'<rect x="{x}" y="0" width="{w}" height="1"/>')
        x += w
    total = x + quiet
    box = f"0 0 1 {total}" if vertical else f"0 0 {total} 1"
    return (f'<svg class="bars" viewBox="{box}" preserveAspectRatio="none" '
            f'shape-rendering="crispEdges" xmlns="http://www.w3.org/2000/svg">'
            f'{"".join(rects)}</svg>')


LABEL_CSS = """
*{box-sizing:border-box}
body{margin:0;padding:8px;font-family:Arial,'Cairo',sans-serif;background:#fff;color:#000}
.noprint{margin-bottom:10px;font-family:Cairo,Arial,sans-serif}
.noprint button{font-weight:700;padding:8px 18px;border-radius:10px;border:1px solid #242421;
  background:#fff;cursor:pointer;font-family:inherit}
.sheet{display:flex;flex-wrap:wrap;gap:8px}
.label{width:__W__mm;height:__H__mm;position:relative;overflow:hidden;background:#fff;
  outline:1px solid #ccc}
.lin{position:absolute;top:0;left:0;width:__DW__mm;height:__DH__mm;padding:1.2mm 1.5mm;
  overflow:hidden;text-align:center;line-height:1.15;
  -webkit-transform-origin:0 0;transform-origin:0 0}
.lin.r90{left:__W__mm;-webkit-transform:rotate(90deg);transform:rotate(90deg)}
.lin.r270{top:__H__mm;-webkit-transform:rotate(-90deg);transform:rotate(-90deg)}
.hd{font-size:__F1__mm;font-weight:800;white-space:nowrap;overflow:hidden}
.hd2{font-size:__F2__mm;font-weight:700;white-space:nowrap;overflow:hidden}
.desc{font-size:__F1__mm;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.meta{font-size:__F2__mm}
.bars{display:block;width:100%;margin:.5mm 0}
.code{font-size:__F2__mm;letter-spacing:.3mm;font-family:'Courier New',monospace}
.price{font-size:__F3__mm;font-weight:800}
.lin.vert{padding:0}
.vbar{position:absolute;top:1.5mm;bottom:1.5mm;left:1.5mm;width:40%}
.vbar .bars{position:absolute;left:0;top:0;width:72%;height:100%;margin:0}
.vcode{position:absolute;left:76%;top:0;height:100%;width:24%;
  -webkit-writing-mode:vertical-rl;writing-mode:vertical-rl;
  -webkit-transform:rotate(180deg);transform:rotate(180deg);
  font-size:__F2__mm;letter-spacing:.3mm;font-family:'Courier New',monospace;
  white-space:nowrap;text-align:center}
.vtext{position:absolute;top:1.5mm;bottom:1.5mm;left:44%;right:1.5mm;overflow:hidden}
.vin{position:absolute;left:0;right:0;top:50%;
  -webkit-transform:translateY(-50%);transform:translateY(-50%);text-align:center}
.vin .desc{white-space:normal;overflow:visible;text-overflow:clip;margin-bottom:1.2mm}
.vin .hd,.vin .hd2,.vin .meta{margin-bottom:1mm}
@media print{
  .noprint{display:none!important}
  body{padding:0}
  .sheet{display:block}
  .label{outline:none;page-break-after:always;break-after:page}
  .label:last-child{page-break-after:auto;break-after:auto}
  __PRINT__
}
"""


def labels_html(items, w=50, h=40, layout="label", show_price=True,
                show_meta=True, autoprint=False, vertical=False,
                header1="", header2="", rotate=0):
    """items: قائمة من (dict، عدد النسخ). تُرجع صفحة HTML جاهزة للطباعة.
    rotate: 0 أو 90 أو 270 (تدوير محتوى الملصق داخل الورقة)."""
    esc = htmllib.escape
    dw, dh = (h, w) if rotate in (90, 270) else (w, h)
    k = max(0.5, min(dw / 50, dh / 30))
    header1, header2 = str(header1).strip(), str(header2).strip()
    rot_cls = {90: " r90", 270: " r270"}.get(rotate, "")
    labels = []
    for it, copies in items:
        code = str(it.get("barcode", "")).strip()
        if not code:
            continue
        meta = " / ".join(x for x in [str(it.get("color", "")).strip(),
                                      str(it.get("size", "")).strip()] if x)
        price_txt = ""
        if show_price:
            p = pd.to_numeric(it.get("price"), errors="coerce")
            if not pd.isna(p) and float(p) > 0:
                price_txt = f"{float(p):,.2f}"
        svg = barcode_svg(code, vertical=vertical)
        desc_txt = esc(str(it.get("desc", "")))
        hd = ((f'<div class="hd">{esc(header1)}</div>' if header1 else "")
              + (f'<div class="hd2">{esc(header2)}</div>' if header2 else ""))
        meta_html = f'<div class="meta">{esc(meta)}</div>' if show_meta and meta else ""
        price_html = f'<div class="price">{price_txt}</div>' if price_txt else ""
        if vertical:
            inner = (
                f'<div class="lin vert{rot_cls}">'
                '<div class="vbar">' + (svg or '<div class="bars"></div>')
                + f'<div class="vcode">{esc(code)}</div></div>'
                '<div class="vtext"><div class="vin">'
                + hd + f'<div class="desc">{desc_txt}</div>' + meta_html + price_html
                + "</div></div></div>"
            )
        else:
            # ارتفاع الباركود = المتبقي بعد أسطر النص
            lines = (3.0 * k * (1 if header1 else 0) + 2.6 * k * (1 if header2 else 0)
                     + 3.0 * k + (2.6 * k if meta_html else 0) + 2.6 * k
                     + (3.6 * k if price_html else 0))
            bar_h = max(5.0, dh - 2.4 - lines * 1.22 - 1.5)
            inner = (
                f'<div class="lin{rot_cls}">' + hd
                + f'<div class="desc">{desc_txt}</div>' + meta_html
                + (svg.replace('class="bars"', f'class="bars" style="height:{bar_h:.1f}mm"')
                   if svg else f'<div class="bars" style="height:{bar_h:.1f}mm"></div>')
                + f'<div class="code">{esc(code)}</div>' + price_html + "</div>"
            )
        labels.extend([f'<div class="label">{inner}</div>'] * int(copies))

    if layout == "a4":
        print_rules = "@page{size:A4;margin:8mm} .sheet{display:flex!important;flex-wrap:wrap;gap:2mm} .label{outline:.2mm dashed #999!important;page-break-after:auto!important;break-after:auto!important;break-inside:avoid}"
    else:
        print_rules = f"@page{{size:{w}mm {h}mm;margin:0}}"

    css = (LABEL_CSS.replace("__W__", str(w)).replace("__H__", str(h))
           .replace("__DW__", str(dw)).replace("__DH__", str(dh))
           .replace("__F1__", f"{3.0 * k:.2f}").replace("__F2__", f"{2.6 * k:.2f}")
           .replace("__F3__", f"{3.6 * k:.2f}").replace("__PRINT__", print_rules))
    script = ("<script>window.onload=function(){setTimeout(function(){window.print()},400)}</script>"
              if autoprint else "")
    return (
        f'<!doctype html><html><head><meta charset="utf-8"><style>{css}</style></head><body>'
        f'<div class="noprint"><button onclick="window.print()">🖨️ طباعة الملصقات</button> '
        f'<span>{len(labels)} ملصق</span></div>'
        f'<div class="sheet">{"".join(labels)}</div>{script}</body></html>'
    )


def stock_status(data, threshold):
    q = qty(data)
    return pd.Series(
        np.select([q <= 0, q <= threshold], ["نفد", "منخفض"], "متوفر"),
        index=data.index,
    )


def display_columns(role, with_status=True):
    cols = ["الباركود التلقائي", "وصف", "نوع الصنف", "الماركة", "اللون",
            "المقاس", "رقم الموديل", "الكمية"]
    cols.append("سعر البيع")
    return cols + (["الحالة"] if with_status else [])


def summary_by(data, column):
    group = data[column].replace("", "غير محدد")
    out = (
        pd.DataFrame({column: group, "q": qty(data)})
        .groupby(column)
        .agg(**{"عدد الأصناف": ("q", "size"), "عدد القطع": ("q", "sum")})
        .sort_values("عدد القطع", ascending=False)
        .reset_index()
    )
    out["عدد القطع"] = out["عدد القطع"].round(0).astype(int)
    return out


def filter_ui(data, key, threshold, hide_out_by_default=False):
    c1, c2, c3, c4, c5 = st.columns([2.2, 1, 1, 1, 1])
    query = c1.text_input("بحث", placeholder="الوصف، الباركود، الماركة، رقم الموديل…",
                          key=f"{key}_q")
    types = ["الكل"] + sorted(t for t in data["نوع الصنف"].unique() if t)
    brands = ["الكل"] + sorted(b for b in data["الماركة"].unique() if b)
    seasons = ["الكل"] + sorted(s for s in data["الموسم"].unique() if s)
    f_type = c2.selectbox("النوع", types, key=f"{key}_type")
    f_brand = c3.selectbox("الماركة", brands, key=f"{key}_brand")
    f_season = c4.selectbox("الموسم", seasons, key=f"{key}_season")
    status_opts = ([AVAILABLE_NOT_OUT] + STATUS_OPTIONS) if hide_out_by_default else STATUS_OPTIONS
    f_status = c5.selectbox("حالة المخزون", status_opts, key=f"{key}_status")

    mask = pd.Series(True, index=data.index)
    if query.strip():
        hay = data[SEARCH_COLUMNS].astype(str)
        for term in query.split():
            mask &= hay.apply(
                lambda s, t=term: s.str.contains(t, case=False, na=False, regex=False)
            ).any(axis=1)
    if f_type != "الكل":
        mask &= data["نوع الصنف"] == f_type
    if f_brand != "الكل":
        mask &= data["الماركة"] == f_brand
    if f_season != "الكل":
        mask &= data["الموسم"] == f_season
    if f_status == AVAILABLE_NOT_OUT:
        mask &= qty(data) > 0
    elif f_status != "الكل":
        mask &= stock_status(data, threshold) == f_status
    return data[mask]


df, load_error = load_data()
threshold = int(st.session_state.get("low_threshold", 5))

with st.sidebar:
    brand_block()
    st.markdown(
        f'<div class="user-chip">👤 {current_username}<br>'
        f'<span style="color:#D6C1A3;font-size:13px;">{ROLE_LABELS.get(current_role, "")}</span></div>',
        unsafe_allow_html=True,
    )
    pages = ["🏠 الرئيسية", "📦 المخزون", "🔎 البحث", "📑 التقارير والطباعة"]
    if current_role == "admin":
        pages = ["🏠 الرئيسية", "📦 المخزون", "➕ إضافة صنف", "🏷️ طباعة الباركود",
                 "🚚 التوصيل", "🔎 البحث", "📑 التقارير والطباعة", "👤 إدارة المستخدمين"]
    page = st.radio("القائمة", pages, label_visibility="collapsed", key="nav")
    st.markdown('<div class="tagline">MORE THAN JUST<br>A STORE<div></div></div>', unsafe_allow_html=True)
    if st.button("🚪 تسجيل الخروج", use_container_width=True):
        st.session_state.clear()
        st.rerun()

top_bar()
show_flash()
if load_error:
    st.error(load_error)


def require_admin():
    if current_role != "admin":
        st.error("هذه الصفحة للمدير فقط.")
        st.stop()


# =========================
# الرئيسية
# =========================
def page_home():
    esc = htmllib.escape
    bg = (f'<img src="data:image/png;base64,{image_base64}" alt="">' if image_base64
          else '<div class="banner-fallback"></div>')
    c_img, c_txt = st.columns([1.5, 1])
    c_img.markdown(
        '<div class="banner">' + bg + '</div>',
        unsafe_allow_html=True,
    )
    c_txt.markdown(
        '<div class="hello"><div class="hello-s">مرحبًا بك في</div>'
        '<div class="hello-b">OvvO Store</div>'
        f'<div class="hello-d">أهلًا {esc(current_username)}، إدارة المخزون بكل سهولة من أي مكان</div></div>',
        unsafe_allow_html=True,
    )

    if df.empty:
        st.info("لا توجد أصناف بعد. " + ("أضف أول صنف من صفحة «إضافة صنف»." if current_role == "admin" else ""))
        return

    q = qty(df)
    out_count = int((q <= 0).sum())
    low_count = int(((q > 0) & (q <= threshold)).sum())
    models_n = df.loc[df["رقم الموديل"] != "", "رقم الموديل"].nunique()

    st.markdown(
        '<div class="kpis">'
        + kpi_card("box", "إجمالي القطع", f"{q.sum():,.0f}", "في المستودع", "#7A1229")
        + kpi_card("list", "عدد الأصناف", f"{len(df):,}", f"{models_n:,} موديل", "#0D1B2E")
        + kpi_card("warn", "مخزون منخفض", str(low_count), f"الكمية ≤ {threshold}", "#E8830C")
        + kpi_card("stop", "نفد من المخزون", str(out_count), "يحتاج إعادة توريد", "#D8453B")
        + '</div>',
        unsafe_allow_html=True,
    )

    status = stock_status(df, threshold)
    need = df[status != "متوفر"].copy()
    need["الحالة"] = status[need.index]
    need = need.sort_values("الكمية")

    is_admin = current_role == "admin"
    if is_admin:
        c_side, c_list = st.columns([1, 1.8])
    else:
        c_side, c_list = None, st.container()

    with c_list:
        if need.empty:
            st.markdown(
                '<div class="panel"><div class="panel-h">أصناف تحتاج إعادة توريد</div>'
                '<div class="ok-note">✅ كل الأصناف ضمن الحد الآمن.</div></div>',
                unsafe_allow_html=True)
        else:
            foot = (f'<div class="foot-note">يعرض أول 8 من {len(need):,} — القائمة الكاملة في '
                    'صفحة «التقارير والطباعة».</div>' if len(need) > 8 else "")
            st.markdown(
                '<div class="panel"><div class="panel-h">أصناف تحتاج إعادة توريد '
                f'<span class="count">{len(need):,}</span></div>'
                + stock_table_html(need) + foot + '</div>',
                unsafe_allow_html=True)
        with st.expander("إعدادات التنبيه"):
            st.number_input(
                "يُعتبر الصنف منخفضًا عندما تكون كميته أقل من أو تساوي",
                min_value=1, max_value=1000, value=5, step=1, key="low_threshold",
            )

    if is_admin:
        with c_side:
            o = load_orders()
            amount = pd.to_numeric(o["المبلغ المطلوب"], errors="coerce").fillna(0)
            out_n = int((o["الحالة"] == ST_OUT).sum())
            pending = float(amount[o["الحالة"] == ST_OUT].sum())
            collected = float(amount[o["الحالة"] == ST_DELIVERED].sum())
            wait_n = int(o["الحالة"].isin([ST_RET_WAIT, ST_CAN_WAIT]).sum())
            st.markdown(
                '<div class="panel"><div class="panel-h">🚚 التوصيل والتحصيل</div>'
                f'<div class="mini-row"><span>طلبات قيد التوصيل</span><b>{out_n}</b></div>'
                f'<div class="mini-row"><span>كاش معلّق عند المندوب</span><b>{pending:,.2f}</b></div>'
                f'<div class="mini-row"><span>مرتجعات بانتظار الاستلام</span><b>{wait_n}</b></div>'
                f'<div class="mini-row"><span>كاش محصّل</span><b>{collected:,.2f}</b></div></div>',
                unsafe_allow_html=True)
            st.button("إدارة الطلبات والتوصيل ←", type="primary",
                      use_container_width=True, on_click=go_delivery)

    st.markdown('<div class="panel-h" style="margin-top:22px">توزيع المخزون حسب النوع</div>',
                unsafe_allow_html=True)
    by_type = summary_by(df, "نوع الصنف")
    t1, t2 = st.columns([1, 1.4])
    t1.dataframe(by_type, use_container_width=True, hide_index=True)
    try:
        t2.bar_chart(by_type.set_index("نوع الصنف")["عدد القطع"], color="#7A1229")
    except TypeError:
        t2.bar_chart(by_type.set_index("نوع الصنف")["عدد القطع"])


# =========================
# المخزون
# =========================
def render_models_view(data, can_edit=True):
    data = data.assign(_q=qty(data))
    with_model = data[data["رقم الموديل"] != ""]
    no_model = len(data) - len(with_model)

    c1, c2 = st.columns([1, 1])
    only_no_img = c1.checkbox("أظهر الموديلات بدون صورة فقط", key="mv_noimg")
    per_page = c2.selectbox("عدد الموديلات بالصفحة", [9, 12, 24, 48], index=1, key="mv_pp")

    models = sorted(with_model["رقم الموديل"].unique())
    if only_no_img:
        models = [m for m in models if not get_model_image(m)]
    if not models:
        st.info("لا توجد موديلات مطابقة.")
        return

    pages = max(1, -(-len(models) // per_page))
    pg = st.number_input("الصفحة", min_value=1, max_value=pages, value=1, step=1,
                         key="mv_page") if pages > 1 else 1
    chunk = models[(pg - 1) * per_page: pg * per_page]
    st.caption(f"{len(models):,} موديل"
               + (f" — و{no_model} صنف بدون رقم موديل (تظهر في «جدول التعديل»)" if no_model else ""))

    for row_start in range(0, len(chunk), 3):
        cols = st.columns(3)
        for col, model in zip(cols, chunk[row_start:row_start + 3]):
            g = with_model[with_model["رقم الموديل"] == model]
            colors = g["اللون"].replace("", "غير محدد")
            sizes = g["المقاس"].replace("", "غير محدد")
            with col, st.container(border=True):
                img = get_model_image(model)
                if img:
                    st.image(img, use_container_width=True)
                else:
                    st.markdown(
                        '<div style="height:180px;display:flex;align-items:center;'
                        'justify-content:center;background:#EFE6D6;border-radius:10px;'
                        'color:#8a7f6d;">لا توجد صورة</div>', unsafe_allow_html=True)
                st.markdown(f"**{g['وصف'].iloc[0]}**  \nموديل: `{model}`")
                m1, m2, m3 = st.columns(3)
                m1.metric("القطع", f"{g['_q'].sum():,.0f}")
                m2.metric("الألوان", colors.nunique())
                m3.metric("المقاسات", sizes.nunique())

                with st.expander("الألوان والمقاسات"):
                    pv = pd.pivot_table(
                        pd.DataFrame({"اللون": colors, "المقاس": sizes, "q": g["_q"]}),
                        index="اللون", columns="المقاس", values="q",
                        aggfunc="sum", fill_value=0,
                    )
                    pv = pv[sorted(pv.columns, key=size_key)]
                    pv["المجموع"] = pv.sum(axis=1)
                    pv.loc["المجموع"] = pv.sum()
                    st.dataframe(pv.astype(int), use_container_width=True)

                if can_edit:
                    with st.expander("🖼️ تغيير الصورة"):
                        up = st.file_uploader("اختر صورة", type=["png", "jpg", "jpeg", "webp"],
                                              key=f"mv_up_{model}")
                        if up and st.button("حفظ الصورة", key=f"mv_save_{model}"):
                            if save_model_image(model, up):
                                flash(f"تم حفظ صورة الموديل {model}.")
                                st.rerun()


def page_inventory():
    can_edit = current_role == "admin"
    if can_edit:
        page_header("إدارة المخزون", "عدّل الخلايا مباشرة ثم اضغط «حفظ التعديلات». تؤخذ نسخة احتياطية تلقائيًا قبل كل حفظ.")
    else:
        page_header("المخزون", "عرض فقط — لا يمكن التعديل.")
    if df.empty:
        st.info("لا توجد أصناف بعد." + (" أضف أول صنف من صفحة «إضافة صنف»." if can_edit else ""))
        return

    filtered = filter_ui(df, "inv", threshold, hide_out_by_default=True)
    st.caption(f"يعرض {len(filtered):,} من {len(df):,} صنف")

    table_label = "✏️ جدول التعديل" if can_edit else "📋 الجدول"
    mode = st.radio("طريقة العرض", ["🖼️ حسب الموديل", table_label],
                    horizontal=True, label_visibility="collapsed")
    if mode == "🖼️ حسب الموديل":
        render_models_view(filtered, can_edit)
        return

    if not can_edit:
        st.dataframe(filtered[VISIBLE_COLUMNS], use_container_width=True, hide_index=True)
        return

    view = filtered[VISIBLE_COLUMNS].copy()
    view.insert(0, "حذف", False)
    config = {
        "حذف": st.column_config.CheckboxColumn("حذف", default=False),
        "الباركود التلقائي": st.column_config.TextColumn("الباركود التلقائي"),
        "وصف": st.column_config.TextColumn("وصف"),
        "وصف 2": st.column_config.TextColumn("وصف 2"),
        "الكمية": st.column_config.NumberColumn("الكمية", min_value=0, step=1, format="%.0f"),
        "سعر البيع": st.column_config.NumberColumn("سعر البيع", min_value=0, format="%.2f"),
        "سعر البيع 2": st.column_config.NumberColumn("سعر البيع 2", min_value=0, format="%.2f"),
    }
    edited = st.data_editor(view, column_config=config, num_rows="fixed",
                            hide_index=True, use_container_width=True,
                            key="inventory_editor")

    to_delete = edited.index[edited["حذف"]]
    confirm = True
    if len(to_delete):
        confirm = st.checkbox(f"أؤكد حذف {len(to_delete)} صنف نهائيًا", key="confirm_delete")

    if st.button("💾 حفظ التعديلات", type="primary", use_container_width=True):
        if len(to_delete) and not confirm:
            st.warning("علّمت أصنافًا للحذف. فعّل خانة التأكيد أولًا.")
            return
        base = df.copy()
        changes = edited.drop(columns=["حذف"])
        base.loc[changes.index, VISIBLE_COLUMNS] = changes[VISIBLE_COLUMNS]
        base = base.drop(index=to_delete)
        base = normalize(base)
        codes = base["الباركود التلقائي"]
        codes = codes[codes != ""]
        dups = codes[codes.duplicated()].unique()
        if len(dups):
            st.error("لا يمكن الحفظ: باركود مكرر → " + "، ".join(dups[:5]))
            return
        if save_data(base):
            msg = "تم حفظ التعديلات."
            if len(to_delete):
                msg += f" وحُذف {len(to_delete)} صنف."
            flash(msg)
            st.rerun()


# =========================
# إضافة صنف
# =========================
def page_add_item():
    require_admin()
    page_header("إضافة صنف جديد", "الحقول المميّزة بـ * مطلوبة. الباقي اختياري.")
    show_just_added()
    form_id = st.session_state.get("add_form_id", 0)
    types = known_names("types", df["نوع الصنف"].unique())
    brands = sorted(b for b in df["الماركة"].unique() if b)
    none_opt = "— اختر —"

    with st.form(f"add_item_form_{form_id}"):
        st.markdown("#### بيانات الصنف")
        c1, c2 = st.columns(2)
        desc = c1.text_input("الوصف *")
        desc2 = c2.text_input("وصف 2")

        c1, c2, c3 = st.columns(3)
        type_sel = c1.selectbox("نوع الصنف *", [none_opt] + types)
        type_new = c1.text_input("أو اكتب نوعًا جديدًا")
        brand_sel = c2.selectbox("الماركة", [none_opt] + brands)
        brand_new = c2.text_input("أو اكتب ماركة جديدة")
        model = c3.text_input(
            "رقم الموديل",
            help="اتركه فارغًا لتوليد رقم جديد من 3 أرقام. "
                 "اكتب نفس الرقم لإضافة لون أو مقاس لموديل موجود.")
        barcode_in = c3.text_input("باركود يدوي (اختياري، الأفضل تركه فارغًا)")

        st.markdown("#### ملصق الباركود")
        p1, p2 = st.columns(2)
        auto_print = p1.checkbox("اطبع ملصق الباركود مباشرة بعد الحفظ", value=True)
        print_copies = p2.number_input("عدد الملصقات", min_value=1, max_value=500,
                                       value=1, step=1)

        st.markdown("#### صورة المنتج")
        img_file = st.file_uploader(
            "صورة الموديل (اختياري، وإذا الموديل إله صورة قبل رح تنستبدل)",
            type=["png", "jpg", "jpeg", "webp"], key=f"img_{form_id}")
        cam_file = st.camera_input("أو التقط صورة بالكاميرا", key=f"cam_{form_id}")

        c1, c2 = st.columns(2)
        color_sel = c1.selectbox("اللون", [none_opt] + known_names("colors"))
        color_new = c1.text_input("أو اكتب لونًا جديدًا (ياخد كود جديد تلقائيًا)")
        size_sel = c2.selectbox("المقاس", [none_opt] + SIZES)
        size_new = c2.text_input("أو اكتب مقاسًا آخر")

        st.markdown("#### الكمية والأسعار")
        c1, c2, c3, c4 = st.columns(4)
        quantity = c1.number_input("الكمية *", min_value=0, step=1, value=0)
        cost = c2.number_input("الكلفة النهائية (داخلي)", min_value=0.0, step=0.5,
                               format="%.2f",
                               help="تُحفظ في ملف Excel فقط ولا تظهر في التطبيق.")
        price = c3.number_input("سعر البيع", min_value=0.0, step=0.5, format="%.2f")
        price2 = c4.number_input("سعر البيع 2", min_value=0.0, step=0.5, format="%.2f")

        with st.expander("تفاصيل إضافية"):
            e1, e2, e3 = st.columns(3)
            blank = lambda x: x or "— اختر —"
            season = e1.selectbox("الموسم", ["", "صيفي", "شتوي"], format_func=blank)
            year = e2.text_input("سنة الشراء")
            order = e3.text_input("الطلبية")
            invoice = e1.text_input("فاتورة رقم")
            origin = e2.text_input("المنشأة")
            tax = e3.text_input("ضريبة (%)", value="16", disabled=True)
            hijab = e1.selectbox("محجب أو غير محجب", ["", "محجب", "غير محجب"],
                                 format_func=blank)
            sleeve = e2.selectbox("كت أو كم أو نص كم", ["", "كت", "كم", "نص كم"],
                                  format_func=blank)
            item_code = e3.text_input("كود الصنف")

        submitted = st.form_submit_button("حفظ الصنف", type="primary", use_container_width=True)

    if not submitted:
        return

    item_type = type_new.strip() or (type_sel if type_sel != none_opt else "")
    brand = brand_new.strip() or (brand_sel if brand_sel != none_opt else "")
    color = color_new.strip() or (color_sel if color_sel != none_opt else "")
    size = size_new.strip() or (size_sel if size_sel != none_opt else "")
    model = model.strip() or next_model(df)
    if model.isdigit():
        model = model.zfill(3)
    photo = img_file or cam_file

    if not desc.strip():
        st.error("اكتب وصف الصنف.")
        return
    if not item_type:
        st.error("اختر نوع الصنف أو اكتب نوعًا جديدًا.")
        return

    color_c = get_code("colors", color) if color else "00"
    size_c = size_code(size) if size else FREE_SIZE_CODE
    type_c = get_code("types", item_type)
    barcode = barcode_in.strip() or build_barcode(model, color_c, size_c, type_c)

    if (df["الباركود التلقائي"] == barcode).any():
        st.error(f"الباركود {barcode} مستخدم لصنف آخر (نفس الموديل واللون والمقاس). "
                 "عدّل كميته من صفحة «المخزون».")
    else:
        row = {c: "" for c in COLUMNS}
        row.update({
            "الباركود التلقائي": barcode, "وصف": desc.strip(), "وصف 2": desc2.strip(),
            "الكلفة النهائية": cost, "سعر البيع": price, "سعر البيع 2": price2,
            "الكمية": quantity, "ضريبة": tax.strip(), "نوع الصنف": item_type,
            "الماركة": brand, "المنشاة": origin.strip(), "الموسم": season.strip(),
            "سنة الشراء": year.strip(), "الطلبية": order.strip(),
            "محجب او غير محجب": hijab.strip(), "كت او كم او نص كم": sleeve.strip(),
            "فاتورة رقم": invoice.strip(), "اللون": color, "المقاس": size,
            "رقم الموديل": model, "كود الصنف": item_code.strip() or type_c,
            "كود اللون تلقائي": color_c, "كود المقاس تلقائي": size_c,
            "كود النوع تلقائي": type_c,
        })
        updated = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        if save_data(updated):
            if photo:
                save_model_image(model, photo)
            st.session_state["just_added"] = {
                "item": {"barcode": barcode, "desc": desc.strip(), "color": color,
                         "size": size, "price": price},
                "copies": int(print_copies),
                "auto": bool(auto_print),
            }
            st.session_state["add_form_id"] = form_id + 1
            st.rerun()


# =========================
# التوصيل
# =========================
def load_orders():
    if not os.path.isfile(ORDERS_FILE):
        return pd.DataFrame(columns=ORDER_COLS)
    try:
        o = pd.read_excel(ORDERS_FILE, dtype={"الباركود": str})
        for c in ORDER_COLS:
            if c not in o.columns:
                o[c] = ""
        o = o[ORDER_COLS]
        o["الباركود"] = o["الباركود"].astype(str).str.replace(r"\.0$", "", regex=True)
        o["أُعيد للمخزون"] = o["أُعيد للمخزون"].fillna(False).astype(bool)
        return o
    except Exception as e:
        st.error(f"تعذّرت قراءة ملف الطلبات: {e}")
        return pd.DataFrame(columns=ORDER_COLS)


def save_orders(o):
    try:
        o.to_excel(ORDERS_FILE, index=False)
        return True
    except PermissionError:
        st.error("ملف OVVO_orders.xlsx مفتوح في Excel. أغلقه ثم أعد المحاولة.")
    except Exception as e:
        st.error(f"تعذّر حفظ الطلبات: {e}")
    return False


def adjust_stock(barcode, delta):
    data, _ = load_data()
    idx = data.index[data["الباركود التلقائي"] == str(barcode)]
    if len(idx) == 0:
        st.error("الباركود غير موجود في المخزون.")
        return False
    i = idx[0]
    data.loc[i, "الكمية"] = float(data.loc[i, "الكمية"]) + delta
    return save_data(data)


def update_order(orders, i, status, restock=False):
    row = orders.loc[i]
    if restock and not row["أُعيد للمخزون"]:
        if not adjust_stock(row["الباركود"], float(row["الكمية"])):
            return
        orders.loc[i, "أُعيد للمخزون"] = True
    orders.loc[i, "الحالة"] = status
    if save_orders(orders):
        flash(f"طلب رقم {int(row['رقم الطلب'])}: {status}")
        st.rerun()


def page_delivery():
    require_admin()
    page_header("التوصيل", "أدخل الباركود والكمية فقط، وباقي المعلومات تظهر تلقائيًا.")
    orders = load_orders()
    tab_new, tab_open, tab_log = st.tabs(["➕ طلب جديد", "🚚 الطلبات المفتوحة", "📋 السجل"])

    with tab_new:
        with st.form("new_order", clear_on_submit=True):
            c1, c2 = st.columns([2, 1])
            barcode = c1.text_input("الباركود")
            q_ = c2.number_input("الكمية", min_value=1, value=1, step=1)
            go = st.form_submit_button("إرسال مع التوصيل", type="primary",
                                       use_container_width=True)
        if go:
            code = barcode.strip()
            hit = df[df["الباركود التلقائي"] == code]
            if not code or hit.empty:
                st.error("الباركود غير موجود.")
            else:
                item = hit.iloc[0]
                if float(item["الكمية"]) < q_:
                    st.error(f"الكمية المتوفرة {float(item['الكمية']):.0f} فقط.")
                elif adjust_stock(code, -float(q_)):
                    price = pd.to_numeric(item["سعر البيع"], errors="coerce")
                    price = 0.0 if pd.isna(price) else float(price)
                    new_id = int(orders["رقم الطلب"].max()) + 1 if len(orders) else 1
                    row = {
                        "رقم الطلب": new_id,
                        "التاريخ": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "الباركود": code, "الوصف": item["وصف"],
                        "اللون": item["اللون"], "المقاس": item["المقاس"],
                        "الكمية": int(q_), "سعر البيع": price,
                        "المبلغ المطلوب": price * q_,
                        "الحالة": ST_OUT, "أُعيد للمخزون": False,
                    }
                    orders = pd.concat([orders, pd.DataFrame([row])], ignore_index=True)
                    if save_orders(orders):
                        flash(f"تم إنشاء الطلب رقم {new_id} وخصم الكمية من المخزون.")
                        st.rerun()

    with tab_open:
        open_o = orders[orders["الحالة"].isin(OPEN_STATUSES)]
        if open_o.empty:
            st.info("لا توجد طلبات مفتوحة.")
        for i, r in open_o.iterrows():
            with st.container(border=True):
                st.markdown(
                    f"**طلب #{int(r['رقم الطلب'])}** — {r['الوصف']} "
                    f"({r['اللون']} / {r['المقاس']}) × الكمية الخارجة: **{int(r['الكمية'])}**  \n"
                    f"الباركود: {r['الباركود']} | المبلغ المطلوب الحالي: **{r['المبلغ المطلوب']:,.2f}** | الحالة: **{r['الحالة']}**"
                )

                if r["الحالة"] == ST_OUT:
                    b1, b2, b3 = st.columns(3)
                    if b1.button("✅ تم التوصيل (كامل المبلغ)", key=f"d{i}", use_container_width=True):
                        update_order(orders, i, ST_DELIVERED)
                    if b2.button("↩️ مرتجع / استلام جزئي (تعديل)", key=f"r_part_{i}", use_container_width=True):
                        st.session_state[f"editing_order_{i}"] = True
                    if b3.button("❌ ملغي بالكامل", key=f"c{i}", use_container_width=True):
                        update_order(orders, i, ST_CAN_WAIT)

                    if st.session_state.get(f"editing_order_{i}", False):
                        st.markdown("---")
                        st.warning("⚠️ **تعديل الطلب لاستلام جزئي / مرتجع معاينة:** حدد كمية القطع المباعة فعلياً، وسيتم تعديل المبلغ وإرجاع الباقي فوراً.")
                        with st.form(f"form_edit_order_{i}"):
                            sold_qty = st.number_input("عدد القطع المباعة (التي تم الاحتفاظ بها)", min_value=0, max_value=int(r["الكمية"]), value=int(r["الكمية"]))
                            submit_edit = st.form_submit_button("تحديث الحساب والمبلغ وطلب المرتجع", type="primary")

                        if submit_edit:
                            return_qty = int(r["الكمية"]) - sold_qty
                            new_amount = sold_qty * float(r["سعر البيع"])

                            orders.loc[i, "الكمية"] = sold_qty
                            orders.loc[i, "المبلغ المطلوب"] = new_amount
                            orders.loc[i, "الحالة"] = ST_RET_WAIT

                            if return_qty > 0:
                                adjust_stock(r["الباركود"], return_qty)
                                orders.loc[i, "أُعيد للمخزون"] = True

                            if save_orders(orders):
                                st.session_state[f"editing_order_{i}"] = False
                                flash(f"تم تحديث الطلب #{int(r['رقم الطلب'])} بنجاح. المتبقي كاش مطلوب: {new_amount:,.2f} وتم تسجيل مرتجع بـ {return_qty} قطع بانتظار استلامها.")
                                st.rerun()

                elif r["الحالة"] == ST_DELIVERED:
                    st.success(f"💵 تم تسليم الطلب بالكامل. الكاش المطلوب تحصيله: {r['المبلغ المطلوب']:,.2f}")
                    if st.button("↩️ تعديل إلى مرتجع (استرجاع لاحق)", key=f"r_del_{i}"):
                        update_order(orders, i, ST_RET_WAIT)
                else:
                    done = ST_RET_DONE if r["الحالة"] == ST_RET_WAIT else ST_CAN_DONE
                    st.warning("⚠️ **تنبيه موظف:** هذا الطلب يحتوي على مرتجع/إلغاء بانتظار استلام القطع من المندوب وتسجيل المطالبة.")
                    if st.checkbox("📥 الموظف استلم القطع المرتجعة وأدخلها للمستودع وإغلاق الطلب", key=f"k{i}"):
                        update_order(orders, i, done, restock=not r["أُعيد للمخزون"])

    with tab_log:
        if orders.empty:
            st.info("لا توجد طلبات بعد.")
        else:
            delivered = orders[orders["الحالة"] == ST_DELIVERED]
            m1, m2 = st.columns(2)
            m1.metric("💵 إجمالي الكاش المسلّم (تم التوصيل)", f"{delivered['المبلغ المطلوب'].sum():,.2f}")
            m2.metric("🚚 طلبات قيد التوصيل", int((orders["الحالة"] == ST_OUT).sum()))
            f = st.selectbox("الحالة", ["الكل"] + sorted(orders["الحالة"].unique()), key="ord_f")
            shown = orders if f == "الكل" else orders[orders["الحالة"] == f]
            shown = shown.drop(columns=["أُعيد للمخزون"]).iloc[::-1]
            st.dataframe(shown, use_container_width=True, hide_index=True)
            download_buttons(shown, "طلبات_التوصيل", "orders")


# =========================
# طباعة الباركود
# =========================
def label_options_ui(prefix):
    """خيارات شكل الملصق (مشتركة بين صفحة الطباعة ولوحة ما بعد الإضافة)."""
    c1, c2, c3, c4 = st.columns(4)
    w = c1.number_input("عرض الملصق (مم)", 20, 200, 50, step=1, key=f"{prefix}_w")
    h = c2.number_input("ارتفاع الملصق (مم)", 15, 200, 40, step=1, key=f"{prefix}_h")
    orient = c3.selectbox("اتجاه الباركود", ["أفقي (مثل الصورة)", "عمودي (بالطول)"],
                          key=f"{prefix}_orient")
    rot = c4.selectbox("تدوير الملصق عند الطباعة", ["بدون", "90°", "270°"],
                       key=f"{prefix}_rot",
                       help="إذا طلعت الطباعة ملفوفة أو بالعرض، جرّب 90° أو 270°.")
    d1, d2, d3, d4 = st.columns(4)
    header1 = d1.text_input("سطر علوي (اسم المحل)", "OVVO STORE", key=f"{prefix}_h1")
    header2 = d2.text_input("سطر ثاني (رقم هاتف مثلًا)", "", key=f"{prefix}_h2")
    show_price = d3.checkbox("أظهر السعر", value=True, key=f"{prefix}_price")
    show_meta = d4.checkbox("أظهر اللون والمقاس", value=True, key=f"{prefix}_meta")
    return {
        "w": int(w), "h": int(h), "vertical": orient.startswith("عمودي"),
        "rotate": 0 if rot == "بدون" else int(rot[:-1]),
        "header1": header1, "header2": header2,
        "show_price": show_price, "show_meta": show_meta,
    }


def show_just_added():
    """لوحة الملصق بعد إضافة صنف مباشرة (تُطبع تلقائيًا أول مرة)."""
    info = st.session_state.get("just_added")
    if not info:
        return
    item = info["item"]
    with st.container(border=True):
        st.success(f"تمت إضافة «{item['desc']}» — الباركود {item['barcode']}")
        with st.expander("شكل ومقاس الملصق"):
            opts = label_options_ui("add_lbl")
        components.html(
            labels_html([(item, info["copies"])], autoprint=info.get("auto", False), **opts),
            height=320, scrolling=True,
        )
        info["auto"] = False
        if st.button("إغلاق", key="close_just_added"):
            st.session_state.pop("just_added", None)
            st.rerun()


def page_barcodes():
    require_admin()
    page_header("طباعة الباركود", "اختر الأصناف وعدد الملصقات، ثم اضغط «طباعة الملصقات».")
    if df.empty:
        st.info("لا توجد أصناف بعد.")
        return

    filtered = filter_ui(df, "bc", threshold)
    st.caption(f"يعرض {len(filtered):,} من {len(df):,} صنف")
    if filtered.empty:
        st.info("لا نتائج.")
        return

    layout_label = st.selectbox("نوع الطباعة", ["ملصقات (طابعة باركود)", "ورقة A4"])
    opts = label_options_ui("bc_lbl")

    k1, k2 = st.columns(2)
    select_all = k1.checkbox("تحديد كل النتائج", value=False)
    from_qty = k2.checkbox("عدد الملصقات = كمية الصنف بالمخزون", value=False)

    view = filtered[["الباركود التلقائي", "وصف", "اللون", "المقاس", "الكمية", "سعر البيع"]].copy()
    view.insert(0, "طباعة", select_all)
    view["عدد الملصقات"] = qty(filtered).clip(lower=1).astype(int) if from_qty else 1

    edited = st.data_editor(
        view, hide_index=True, use_container_width=True,
        disabled=["الباركود التلقائي", "وصف", "اللون", "المقاس", "الكمية", "سعر البيع"],
        column_config={
            "طباعة": st.column_config.CheckboxColumn("طباعة"),
            "الكمية": st.column_config.NumberColumn("الكمية", format="%.0f"),
            "سعر البيع": st.column_config.NumberColumn("سعر البيع", format="%.2f"),
            "عدد الملصقات": st.column_config.NumberColumn(
                "عدد الملصقات", min_value=0, max_value=500, step=1, format="%d"),
        },
        key=f"bc_editor_{len(filtered)}_{int(select_all)}_{int(from_qty)}",
    )

    chosen = edited[edited["طباعة"] & (edited["عدد الملصقات"] > 0)]
    total = int(chosen["عدد الملصقات"].sum()) if len(chosen) else 0
    if not len(chosen):
        st.info("علّم صنفًا واحدًا على الأقل من عمود «طباعة».")
        return
    if total > 500:
        st.error(f"العدد الإجمالي {total} ملصق. الحد الأقصى 500 في المرة الواحدة.")
        return

    items = []
    for _, r in chosen.iterrows():
        items.append((
            {"barcode": r["الباركود التلقائي"], "desc": r["وصف"], "color": r["اللون"],
             "size": r["المقاس"], "price": r["سعر البيع"]},
            int(r["عدد الملصقات"]),
        ))
    bad = [it["barcode"] for it, _ in items if not str(it["barcode"]).strip()]
    if bad:
        st.warning("بعض الأصناف المحددة ما إلها باركود، وتم تجاهلها.")

    st.caption(f"إجمالي الملصقات: {total:,}")
    components.html(
        labels_html(items, layout="a4" if layout_label == "ورقة A4" else "label", **opts),
        height=520, scrolling=True,
    )
    st.caption("اضغط «طباعة الملصقات» داخل المعاينة، واختر الطابعة. "
               "للطابعة الحرارية: اجعل الهوامش «بدون» وحجم الورق نفس مقاس الملصق.")


# =========================
# البحث
# =========================
def page_search():
    page_header("البحث في المخزون", "اكتب أكثر من كلمة لتضييق النتائج، مثل: قميص أسود.")
    if df.empty:
        st.info("لا توجد أصناف بعد.")
        return
    result = filter_ui(df, "search", threshold)
    result = result.assign(الحالة=stock_status(result, threshold))
    st.caption(f"عدد النتائج: {len(result):,}")
    if result.empty:
        st.info("لا نتائج. جرّب كلمة أقصر أو أزل بعض الفلاتر.")
        return
    shown = result[display_columns(current_role)]
    st.dataframe(style_status(shown), use_container_width=True, hide_index=True)
    download_buttons(shown, "نتائج_البحث", "search")


# =========================
# التقارير والطباعة
# =========================
def page_reports():
    page_header("التقارير والطباعة", "حمّل أي تقرير بصيغة Excel أو CSV، أو اطبعه مباشرة.")
    if df.empty:
        st.info("لا توجد أصناف بعد.")
        return

    components.html(
        '<button onclick="window.parent.print()" style="font-family:Cairo,sans-serif;'
        'font-weight:700;padding:8px 18px;border-radius:10px;border:1px solid #242421;'
        'background:#fff;cursor:pointer;">🖨️ طباعة الصفحة</button>',
        height=50,
    )

    status = stock_status(df, threshold)
    full = df.assign(الحالة=status)
    tab1, tab2, tab3, tab4 = st.tabs(["ملخص", "تحتاج إعادة توريد", "كل المخزون", "تقرير التوصيل والحالات"])

    with tab1:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### حسب النوع")
            by_type = summary_by(df, "نوع الصنف")
            st.dataframe(by_type, use_container_width=True, hide_index=True)
        with c2:
            st.markdown("#### حسب الماركة")
            by_brand = summary_by(df, "الماركة")
            st.dataframe(by_brand, use_container_width=True, hide_index=True)
        download_buttons(by_type, "ملخص_الأنواع", "sum_type")

    with tab2:
        need = full[full["الحالة"] != "متوفر"].sort_values("الكمية")
        if need.empty:
            st.success("لا توجد أصناف منخفضة أو نافدة.")
        else:
            shown = need[display_columns(current_role)]
            st.dataframe(style_status(shown), use_container_width=True, hide_index=True)
            download_buttons(shown, "اصناف_تحتاج_توريد", "need")

    with tab3:
        shown = full[display_columns(current_role)]
        st.dataframe(style_status(shown), use_container_width=True, hide_index=True)
        download_buttons(shown, "المخزون_الكامل", "full")

    with tab4:
        st.markdown("#### تقرير حركة التوصيل والتحصيل")
        orders_rep = load_orders()
        if orders_rep.empty:
            st.info("لا توجد بيانات توصيل مسجلة حتى الآن.")
        else:
            total_orders = len(orders_rep)
            delivered_orders = len(orders_rep[orders_rep["الحالة"] == ST_DELIVERED])
            out_orders = len(orders_rep[orders_rep["الحالة"] == ST_OUT])
            cancelled_orders = len(orders_rep[orders_rep["الحالة"].str.contains("ملغي", na=False)])
            returned_orders = len(orders_rep[orders_rep["الحالة"].str.contains("مرتجع", na=False)])

            collected_cash = float(orders_rep[orders_rep["الحالة"] == ST_DELIVERED]["المبلغ المطلوب"].sum())
            pending_cash = float(orders_rep[orders_rep["الحالة"] == ST_OUT]["المبلغ المطلوب"].sum())

            r1, r2, r3, r4 = st.columns(4)
            r1.metric("📦 إجمالي الطلبات", f"{total_orders}")
            r2.metric("✅ تم التوصيل", f"{delivered_orders}")
            r3.metric("🚚 قيد التوصيل", f"{out_orders}")
            r4.metric("❌ الملغاة / المرتجعة", f"{cancelled_orders + returned_orders}")

            st.markdown("---")
            c1, c2 = st.columns(2)
            c1.metric("💵 إجمالي الكاش المُستلَم (المحصل)", f"{collected_cash:,.2f}")
            c2.metric("⏳ الكاش المعلق (عند المندوب/قيد التوصيل)", f"{pending_cash:,.2f}")

            st.markdown("#### تفاصيل السجل الكامل للحالات")
            shown_ord = orders_rep.drop(columns=["أُعيد للمخزون"]).iloc[::-1]
            st.dataframe(shown_ord, use_container_width=True, hide_index=True)
            download_buttons(shown_ord, "تقرير_حركة_التوصيل", "delivery_report")


# =========================
# إدارة المستخدمين
# =========================
def page_users():
    require_admin()
    page_header("إدارة المستخدمين", "أضف حسابات، غيّر كلمات المرور، أو احذف حسابًا.")
    users = load_users()
    st.dataframe(
        pd.DataFrame([{"المستخدم": u, "الصلاحية": ROLE_LABELS.get(i.get("role"), "")}
                      for u, i in users.items()]),
        use_container_width=True, hide_index=True,
    )

    tab_add, tab_pass, tab_del = st.tabs(["إضافة مستخدم", "تغيير كلمة المرور", "حذف مستخدم"])

    with tab_add:
        with st.form("add_user_form", clear_on_submit=True):
            name = st.text_input("اسم المستخدم")
            pw = st.text_input("كلمة المرور (6 أحرف على الأقل)", type="password")
            role_label = st.selectbox("الصلاحية", ["مشاهد", "مدير"])
            ok = st.form_submit_button("إضافة", type="primary")
        if ok:
            name = name.strip()
            if not name:
                st.error("اكتب اسم المستخدم.")
            elif name in users:
                st.error("اسم المستخدم موجود بالفعل.")
            elif len(pw) < 6:
                st.error("كلمة المرور يجب ألا تقل عن 6 أحرف.")
            else:
                hp = hash_password(pw)
                users[name] = {
                    "salt": hp["salt"],
                    "password_hash": hp["password_hash"],
                    "role": "admin" if role_label == "مدير" else "viewer"
                }
                save_users(users)
                flash(f"تمت إضافة المستخدم «{name}».")
                st.rerun()

    with tab_pass:
        sel = st.selectbox("المستخدم", list(users.keys()), key="pw_user")
        new_p = st.text_input("كلمة المرور الجديدة", type="password", key="pw_new")
        if st.button("تحديث كلمة المرور", type="primary"):
            if len(new_p) < 6:
                st.error("كلمة المرور يجب ألا تقل عن 6 أحرف.")
            else:
                hp_new = hash_password(new_p)
                users[sel]["salt"] = hp_new["salt"]
                users[sel]["password_hash"] = hp_new["password_hash"]
                save_users(users)
                flash(f"تم تحديث كلمة مرور «{sel}».")
                st.rerun()

    with tab_del:
        others = [u for u in users if u != current_username]
        if not others:
            st.info("لا يوجد مستخدمون آخرون لحذفهم.")
        else:
            target = st.selectbox("المستخدم", others, key="del_user")
            sure = st.checkbox(f"أؤكد حذف «{target}»", key="del_confirm")
            if st.button("حذف المستخدم"):
                admins = [u for u, i in users.items() if i.get("role") == "admin" and u != target]
                if not sure:
                    st.warning("فعّل خانة التأكيد أولًا.")
                elif not admins:
                    st.error("لا يمكن حذف آخر مدير.")
                else:
                    users.pop(target)
                    save_users(users)
                    flash(f"تم حذف «{target}».")
                    st.rerun()


# =========================
# التوجيه
# =========================
if page == "🏠 الرئيسية":
    page_home()
elif page == "📦 المخزون":
    page_inventory()
elif page == "➕ إضافة صنف":
    page_add_item()
elif page == "🏷️ طباعة الباركود":
    page_barcodes()
elif page == "🚚 التوصيل":
    page_delivery()
elif page == "🔎 البحث":
    page_search()
elif page == "📑 التقارير والطباعة":
    page_reports()
elif page == "👤 إدارة المستخدمين":
    page_users()
