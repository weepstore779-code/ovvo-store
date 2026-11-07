# -*- coding: utf-8 -*-
"""OVVO Store — نظام إدارة المخزون والمستودعات"""
import base64
import hashlib
import hmac
import html as htmllib
import json
import os
import re
import secrets
import shutil
from datetime import date, datetime, timedelta
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
st.set_page_config(page_title="OVVO Store", page_icon="📦", layout="wide")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_NAME = os.path.join(BASE_DIR, "OVVO.xlsx")
ORDERS_FILE = os.path.join(BASE_DIR, "OVVO_orders.xlsx")
MOVES_FILE = os.path.join(BASE_DIR, "OVVO_movements.xlsx")  # سجل حركات المخزون (تاريخ كل إدخال)
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
ROLE_LABELS = {"admin": "مدير", "viewer": "مشاهد", "accountant": "محاسب"}
ROLE_BY_LABEL = {v: k for k, v in ROLE_LABELS.items()}
STATUS_OPTIONS = ["الكل", "متوفر", "منخفض", "نفد"]
AVAILABLE_NOT_OUT = "المتوفر (عدا النافد)"

# ---- التوصيل ----
ORDER_COLS = ["رقم الطلب", "التاريخ", "المندوب", "الباركود", "الوصف", "اللون", "المقاس",
              "الكمية", "سعر البيع", "المبلغ المطلوب", "الحالة", "أُعيد للمخزون",
              "كلفة الوحدة", "تاريخ الحالة"]
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
# مجموعات مقاسات جاهزة لجدول إضافة الصنف
SIZE_SETS = {
    "حروف (S – XXL)": ["S", "M", "L", "XL", "XXL"],
    "حروف موسّعة (XS – XXXL)": ["XS", "S", "M", "L", "XL", "XXL", "XXXL"],
    "أرقام (34 – 44)": ["34", "36", "38", "40", "42", "44"],
    "أرقام (44 – 60)": ["44", "46", "48", "50", "52", "56", "58", "60"],
    "فري سايز": ["فري سايز"],
}
NO_COLOR = "— بدون لون —"
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


MODEL_START = 260  # أول رقم موديل


def next_model(data):
    """رقم موديل جديد تسلسلي: أول موديل 260، وبعدها آخر رقم + 1."""
    nums = pd.to_numeric(data["رقم الموديل"], errors="coerce").dropna()
    nums = nums[(nums >= MODEL_START) & (nums <= 998)]
    nxt = int(nums.max()) + 1 if len(nums) else MODEL_START
    return f"{nxt:03d}"


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
[data-testid="stSidebar"]::before {
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


# ============================================================
# شعار OvvO Store (base64)
# الصق هنا قيمتي LOGO_TOP_B64 و LOGO_SIDE_B64 من ملفك القديم
# بين علامتي الاقتباس. إذا تركتهم فاضيين بيظهر اسم المتجر كنص.
# ============================================================
LOGO_TOP_B64 = ""
LOGO_SIDE_B64 = ""


def logo_html(variant="top"):
    data = LOGO_SIDE_B64 if variant == "side" else LOGO_TOP_B64
    if not data:
        return '<div class="ovvo-brand-name">OVVO <span>STORE</span></div>'
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


# =========================
# سجل الحركات (تاريخ كل إدخال / تعديل / إخراج)
# =========================
MOVE_COLS = ["التاريخ", "نوع الحركة", "رقم الموديل", "الباركود", "الوصف",
             "اللون", "المقاس", "الكمية", "المستخدم", "ملاحظة"]
MV_IN = "إدخال"
MV_ADJ = "تعديل كمية"
MV_OUT = "إخراج (توصيل)"
MV_RET = "إرجاع للمخزون"
MV_DEL = "حذف صنف"


def load_moves():
    empty = pd.DataFrame(columns=MOVE_COLS)
    if not os.path.isfile(MOVES_FILE):
        return empty
    try:
        m = pd.read_excel(MOVES_FILE, dtype=str)
    except Exception as e:
        st.error(f"تعذّرت قراءة سجل الحركات: {e}")
        return empty
    for c in MOVE_COLS:
        if c not in m.columns:
            m[c] = ""
        m[c] = m[c].fillna("").astype(str).str.strip()
    m = m[MOVE_COLS].copy()
    for c in ("الباركود", "رقم الموديل"):
        m[c] = m[c].str.replace(r"^(\d+)\.0$", r"\1", regex=True)
    m["الكمية"] = pd.to_numeric(m["الكمية"], errors="coerce").fillna(0)
    return m


def log_moves(rows):
    """يضيف حركات لسجل الحركات. فشل السجل ما بيمنع حفظ المخزون، بس بيطلّع تحذير."""
    if not rows:
        return True
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    user = st.session_state.get("ovvo_username", "")
    new = []
    for r in rows:
        x = {c: "" for c in MOVE_COLS}
        x.update(r)
        x["التاريخ"] = r.get("التاريخ") or stamp
        x["المستخدم"] = r.get("المستخدم") or user
        new.append(x)
    new_df = pd.DataFrame(new, columns=MOVE_COLS)
    tmp = os.path.join(BASE_DIR, "_OVVO_moves_tmp.xlsx")
    try:
        old = load_moves()
        out = new_df if old.empty else pd.concat([old, new_df], ignore_index=True)
        out.to_excel(tmp, index=False)
        os.replace(tmp, MOVES_FILE)
        return True
    except PermissionError:
        st.warning("تم حفظ المخزون، لكن تعذّر تسجيل الحركة لأن ملف OVVO_movements.xlsx مفتوح في Excel.")
    except Exception as e:
        st.warning(f"تم حفظ المخزون، لكن تعذّر تسجيل الحركة: {e}")
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
    return False


def move_row(kind, r, delta, note=""):
    d = float(delta)
    d = int(d) if d == int(d) else round(d, 2)
    return {"نوع الحركة": kind, "رقم الموديل": r["رقم الموديل"],
            "الباركود": r["الباركود التلقائي"], "الوصف": r["وصف"],
            "اللون": r["اللون"], "المقاس": r["المقاس"], "الكمية": d, "ملاحظة": note}


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


def code128_values(text):
    """Code128 تلقائي: subset C لتتابعات الأرقام (4 فأكثر)، وB للباقي."""
    if not text:
        raise ValueError("empty")
    for ch in text:
        if not 32 <= ord(ch) <= 126:
            raise ValueError("unsupported character")
    vals, mode = [], None
    for p in [p for p in re.split(r"(\d{4,})", text) if p]:
        if p.isdigit() and len(p) >= 4:
            if len(p) % 2:                      # رقم فردي: أول خانة بـ B
                if mode is None:
                    vals.append(104)
                elif mode == "C":
                    vals.append(100)
                mode = "B"
                vals.append(ord(p[0]) - 32)
                p = p[1:]
            if mode is None:
                vals.append(105)
            elif mode == "B":
                vals.append(99)
            mode = "C"
            vals += [int(p[i:i + 2]) for i in range(0, len(p), 2)]
        else:
            if mode is None:
                vals.append(104)
            elif mode == "C":
                vals.append(100)
            mode = "B"
            vals += [ord(c) - 32 for c in p]
    chk = (vals[0] + sum(v * i for i, v in enumerate(vals[1:], 1))) % 103
    return vals + [chk, 106]


def code128_widths(text):
    try:
        vals = code128_values(text)
    except ValueError:
        return []
    return [int(c) for c in "".join(CODE128_PATTERNS[v] for v in vals)]


QUIET = 6  # وحدات بيضاء على كل جهة


def pick_module(total, avail_mm, dot_mm):
    """أكبر عرض وحدة = مضاعف صحيح لنقطة الطابعة ويتسع بالمساحة."""
    for n in (4, 3, 2):
        if total * n * dot_mm <= avail_mm:
            return n * dot_mm
    return avail_mm / total  # ما اتسع: نضغطه (الأفضل تقصير الباركود)


def barcode_svg(widths, vertical=False, style=""):
    if not widths:
        return ""
    x, rects = QUIET, []
    for i, w in enumerate(widths):
        if i % 2 == 0:
            if vertical:
                rects.append(f'<rect x="0" y="{x}" width="1" height="{w}"/>')
            else:
                rects.append(f'<rect x="{x}" y="0" width="{w}" height="1"/>')
        x += w
    total = x + QUIET
    box = f"0 0 1 {total}" if vertical else f"0 0 {total} 1"
    return (f'<svg class="bars" viewBox="{box}" preserveAspectRatio="none" '
            f'shape-rendering="crispEdges" style="{style}" '
            f'xmlns="http://www.w3.org/2000/svg">{"".join(rects)}</svg>')


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
  html,body{margin:0;padding:0}
  body{padding:0}
  .sheet{display:block}
  .label{outline:none;height:calc(__H__mm - .3mm);page-break-after:always;break-after:page}
  .label:last-child{page-break-after:auto;break-after:auto}
  __PRINT__
}
"""


def labels_html(items, w=40, h=30, layout="label", show_price=True,
                show_meta=True, autoprint=False, vertical=False,
                header1="", header2="", rotate=0, dpi=203):
    """items: قائمة من (dict، عدد النسخ). تُرجع صفحة HTML جاهزة للطباعة.
    rotate: 0 أو 90 أو 270 (تدوير محتوى الملصق داخل الورقة).
    dpi: دقة الطابعة، عشان عرض خط الباركود يكون مضاعف صحيح لنقطة الطابعة."""
    esc = htmllib.escape
    dot = 25.4 / int(dpi)
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
        widths = code128_widths(code)
        total = (sum(widths) + 2 * QUIET) if widths else 0
        desc_txt = esc(str(it.get("desc", "")))
        hd = ((f'<div class="hd">{esc(header1)}</div>' if header1 else "")
              + (f'<div class="hd2">{esc(header2)}</div>' if header2 else ""))
        meta_html = f'<div class="meta">{esc(meta)}</div>' if show_meta and meta else ""
        price_html = f'<div class="price">{price_txt}</div>' if price_txt else ""
        if vertical:
            svg = ""
            if widths:
                length = total * pick_module(total, dh - 3.0, dot)
                svg = barcode_svg(
                    widths, True,
                    f"height:{length:.3f}mm;top:50%;margin-top:-{length / 2:.3f}mm")
            inner = (
                f'<div class="lin vert{rot_cls}">'
                '<div class="vbar">' + (svg or '<div class="bars"></div>')
                + f'<div class="vcode">{esc(code)}</div></div>'
                '<div class="vtext"><div class="vin">'
                + hd + f'<div class="desc">{desc_txt}</div>' + meta_html + price_html
                + "</div></div></div>"
            )
        else:
            # ارتفاع الباركود = المتبقي بعد أسطر النص (مع هامش أمان)
            lines = (3.0 * k * (1 if header1 else 0) + 2.6 * k * (1 if header2 else 0)
                     + 3.0 * k + (2.6 * k if meta_html else 0) + 2.6 * k
                     + (3.6 * k if price_html else 0))
            bar_h = max(5.0, dh - 2.4 - lines * 1.22 - 3.0)
            if widths:
                mod = pick_module(total, dw - 3.0, dot)
                svg = barcode_svg(
                    widths, False,
                    f"width:{total * mod:.3f}mm;height:{bar_h:.1f}mm;margin:.5mm auto")
            else:
                svg = f'<div class="bars" style="height:{bar_h:.1f}mm"></div>'
            inner = (
                f'<div class="lin{rot_cls}">' + hd
                + f'<div class="desc">{desc_txt}</div>' + meta_html + svg
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
    pages = ["🏠 الرئيسية", "📦 المخزون", "🔎 البحث", "📋 كشف المواد", "📑 التقارير والطباعة"]
    if current_role == "accountant":
        pages = ["🏠 الرئيسية", "💰 كلف المنتجات", "🧮 المحاسبة", "📦 المخزون",
                 "🔎 البحث", "📋 كشف المواد", "📑 التقارير والطباعة"]
    if current_role == "admin":
        pages = ["🏠 الرئيسية", "📦 المخزون", "➕ إضافة صنف", "🏷️ طباعة الباركود",
                 "🚚 التوصيل", "🔎 البحث", "📋 كشف المواد", "📑 التقارير والطباعة",
                 "👤 إدارة المستخدمين"]
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
                f'<div class="mini-row"><span>أسطر قيد التوصيل</span><b>{out_n}</b></div>'
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
        old_q = qty(df).loc[changes.index]
        new_q = pd.to_numeric(changes["الكمية"], errors="coerce").fillna(0)
        log = []
        for i in changes.index:
            if i in to_delete:
                if old_q[i]:
                    log.append(move_row(MV_DEL, df.loc[i], -old_q[i], "حذف من جدول التعديل"))
            elif new_q[i] != old_q[i]:
                log.append(move_row(MV_ADJ, df.loc[i], new_q[i] - old_q[i],
                                    "تعديل من صفحة المخزون"))
        if save_data(base):
            log_moves(log)
            msg = "تم حفظ التعديلات."
            if len(to_delete):
                msg += f" وحُذف {len(to_delete)} صنف."
            flash(msg)
            st.rerun()


# =========================
# خيارات الملصق (مشتركة)
# =========================
def label_options_ui(prefix):
    """خيارات شكل الملصق (مشتركة بين صفحة الطباعة ولوحة ما بعد الإضافة). الافتراضي 40×30 مم."""
    c1, c2, c3, c4, c5 = st.columns(5)
    w = c1.number_input("عرض الملصق (مم)", 20, 200, 40, step=1, key=f"{prefix}_w")
    h = c2.number_input("ارتفاع الملصق (مم)", 15, 200, 30, step=1, key=f"{prefix}_h")
    orient = c3.selectbox("اتجاه الباركود", ["أفقي (مثل الصورة)", "عمودي (بالطول)"],
                          key=f"{prefix}_orient")
    rot = c4.selectbox("تدوير الملصق عند الطباعة", ["بدون", "90°", "270°"],
                       key=f"{prefix}_rot",
                       help="إذا طلعت الطباعة ملفوفة أو بالعرض، جرّب 90° أو 270°.")
    dpi = c5.selectbox("دقة الطابعة (DPI)", [203, 300], key=f"{prefix}_dpi",
                       help="أغلب طابعات الملصقات 203. بتلاقيها مكتوبة على الطابعة أو بالمواصفات.")
    d1, d2, d3, d4 = st.columns(4)
    header1 = d1.text_input("سطر علوي (اسم المحل)", "OVVO STORE", key=f"{prefix}_h1")
    header2 = d2.text_input("سطر ثاني (رقم هاتف مثلًا)", "", key=f"{prefix}_h2")
    show_price = d3.checkbox("أظهر السعر", value=True, key=f"{prefix}_price")
    show_meta = d4.checkbox("أظهر اللون والمقاس", value=True, key=f"{prefix}_meta")
    return {
        "w": int(w), "h": int(h), "vertical": orient.startswith("عمودي"),
        "rotate": 0 if rot == "بدون" else int(rot[:-1]),
        "header1": header1, "header2": header2,
        "show_price": show_price, "show_meta": show_meta, "dpi": int(dpi),
    }


def show_just_added():
    """لوحة الملصقات بعد إضافة صنف مباشرة."""
    info = st.session_state.get("just_added")
    if not isinstance(info, dict) or not info.get("items"):
        st.session_state.pop("just_added", None)
        return

    with st.container(border=True):
        st.success(f"تمت إضافة «{info['desc']}» — {len(info['items'])} لون/مقاس")
        with st.expander("شكل ومقاس الملصق"):
            opts = label_options_ui("add_lbl")
        components.html(
            labels_html(info["items"], autoprint=info.get("auto", False), **opts),
            height=320, scrolling=True,
        )
        info["auto"] = False
        if st.button("إغلاق", key="close_just_added"):
            st.session_state.pop("just_added", None)
            st.rerun()
# =========================
# إضافة صنف
# =========================
def page_add_item():
    require_admin()
    page_header("إضافة صنف جديد",
                "أدخل بيانات الصنف مرة وحدة، ثم أضف الألوان والمقاسات والكميات في الجدول.")
    show_just_added()
    form_id = st.session_state.get("add_form_id", 0)
    types = known_names("types", df["نوع الصنف"].unique())
    brands = sorted(b for b in df["الماركة"].unique() if b)
    none_opt = "— اختر —"

    def split_names(txt):
        return [x.strip() for x in txt.replace("،", ",").split(",") if x.strip()]

    def clean(v):
        return "" if v is None or pd.isna(v) else str(v).strip()

    # ---- اختيار ألوان/مقاسات الجدول (خارج الفورم عشان الجدول يتحدّث فورًا) ----
    st.markdown("#### جدول الألوان والمقاسات")
    n1, n2 = st.columns(2)
    new_colors = split_names(n1.text_input(
        "لون جديد مش موجود بالجدول؟ (افصل بينها بفاصلة)", key=f"nc_{form_id}"))
    new_sizes = split_names(n2.text_input(
        "مقاس جديد مش موجود؟ (افصل بينها بفاصلة)", key=f"ns_{form_id}"))
    all_colors = known_names("colors", new_colors)
    all_sizes = SIZES + [x for x in new_sizes if x not in SIZES]

    s1, s2 = st.columns([1, 2])
    size_set = s1.radio("مجموعة المقاسات", list(SIZE_SETS), key=f"sset_{form_id}")
    grid_sizes = s2.multiselect(
        "المقاسات بالجدول", all_sizes,
        default=[x for x in all_sizes if x in SIZE_SETS[size_set] or x in new_sizes],
        key=f"gs_{form_id}_{size_set}_{'|'.join(new_sizes)}")
    grid_colors = st.multiselect(
        "الألوان بالجدول (شيل اللي ما بدك ياه)", all_colors, default=all_colors,
        key=f"gc_{form_id}_{'|'.join(new_colors)}")
    if not grid_sizes:
        st.warning("اختر مقاسًا واحدًا على الأقل لعرض الجدول.")
        return
    grid_rows = grid_colors + [NO_COLOR]
    grid_df = pd.DataFrame(0, index=range(len(grid_rows)), columns=grid_sizes)
    grid_df.insert(0, "اللون", grid_rows)
    grid_sig = hashlib.md5("|".join(grid_rows + ["#"] + grid_sizes).encode("utf-8")).hexdigest()[:8]

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
            help="اتركه فارغًا لتوليد رقم جديد. اكتب رقم موديل موجود "
                 "لإضافة ألوان/مقاسات جديدة عليه.")

        st.markdown("#### الكميات *")
        st.caption("اكتب الكمية بالخانة المناسبة (لون × مقاس). الخانات اللي بتضل صفر "
                   "ما بتنحفظ، فما في داعي تدخل اللون أو المقاس.")
        grid_cfg = {"اللون": st.column_config.TextColumn("اللون", disabled=True)}
        for sz in grid_sizes:
            grid_cfg[sz] = st.column_config.NumberColumn(
                sz, min_value=0, step=1, default=0, format="%d")
        grid_edit = st.data_editor(
            grid_df, num_rows="fixed", hide_index=True, use_container_width=True,
            column_config=grid_cfg, key=f"grid_{form_id}_{grid_sig}")

        st.markdown("#### الأسعار")
        c1, c2, c3 = st.columns(3)
        cost = c1.number_input("الكلفة النهائية (داخلي)", min_value=0.0, step=0.5,
                               format="%.2f",
                               help="تُحفظ في ملف Excel فقط ولا تظهر في التطبيق.")
        price = c2.number_input("سعر البيع", min_value=0.0, step=0.5, format="%.2f")
        price2 = c3.number_input("سعر البيع 2", min_value=0.0, step=0.5, format="%.2f")

        st.markdown("#### ملصق الباركود")
        p1, p2 = st.columns(2)
        auto_print = p1.checkbox("اطبع ملصقات الباركود مباشرة بعد الحفظ", value=True)
        copies_from_qty = p2.checkbox(
            "عدد الملصقات = كمية كل لون/مقاس", value=True,
            help="إذا أطفيتها بيطلع ملصق واحد لكل سطر.")

        st.markdown("#### صورة المنتج")
        img_file = st.file_uploader(
            "صورة الموديل (اختياري، وإذا الموديل إله صورة قبل رح تنستبدل)",
            type=["png", "jpg", "jpeg", "webp"], key=f"img_{form_id}")
        cam_file = st.camera_input("أو التقط صورة بالكاميرا", key=f"cam_{form_id}")

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

        submitted = st.form_submit_button("حفظ الصنف", type="primary",
                                          use_container_width=True)

    if not submitted:
        return

    # ---------- التحقق ----------
    item_type = type_new.strip() or (type_sel if type_sel != none_opt else "")
    brand = brand_new.strip() or (brand_sel if brand_sel != none_opt else "")
    photo = img_file or cam_file

    if not desc.strip():
        st.error("اكتب وصف الصنف.")
        return
    if not item_type:
        st.error("اختر نوع الصنف أو اكتب نوعًا جديدًا.")
        return

    variants = []
    for _, r in grid_edit.iterrows():
        color = "" if r["اللون"] == NO_COLOR else str(r["اللون"]).strip()
        for sz in grid_sizes:
            q = pd.to_numeric(r[sz], errors="coerce")
            if pd.isna(q) or q == 0:
                continue  # صفر = اللون/المقاس مش موجود
            if q < 0 or q != int(q):
                st.error(f"الكمية لازم تكون رقم صحيح موجب (اللون: {color or '—'}، المقاس: {sz}).")
                return
            variants.append((color, sz, int(q)))

    if not variants:
        st.error("ما في كمية مدخلة. اكتب كمية أكبر من صفر بخانة وحدة على الأقل بالجدول.")
        return

    seen = set()
    for color, size, _ in variants:
        key = (norm_key(color), size_code(size))
        if key in seen:
            st.error(f"تكرار في الجدول: اللون «{color or '—'}» والمقاس «{size or '—'}». "
                     "اجمع الكميات في سطر واحد.")
            return
        seen.add(key)

    model = model.strip() or next_model(df)
    if model.isdigit():
        model = model.zfill(3)

    type_c = get_code("types", item_type)
    existing = set(df["الباركود التلقائي"])
    new_rows, label_items, clashes = [], [], []

    for color, size, q in variants:
        color_c = get_code("colors", color) if color else "00"
        size_c = size_code(size) if size else FREE_SIZE_CODE
        barcode = build_barcode(model, color_c, size_c, type_c)
        if barcode in existing:
            clashes.append(f"{color or '—'} / {size or '—'}  ({barcode})")
            continue
        row = {c: "" for c in COLUMNS}
        row.update({
            "الباركود التلقائي": barcode, "وصف": desc.strip(), "وصف 2": desc2.strip(),
            "الكلفة النهائية": cost, "سعر البيع": price, "سعر البيع 2": price2,
            "الكمية": q, "ضريبة": tax.strip(), "نوع الصنف": item_type,
            "الماركة": brand, "المنشاة": origin.strip(), "الموسم": season.strip(),
            "سنة الشراء": year.strip(), "الطلبية": order.strip(),
            "محجب او غير محجب": hijab.strip(), "كت او كم او نص كم": sleeve.strip(),
            "فاتورة رقم": invoice.strip(), "اللون": color, "المقاس": size,
            "رقم الموديل": model, "كود الصنف": item_code.strip() or type_c,
            "كود اللون تلقائي": color_c, "كود المقاس تلقائي": size_c,
            "كود النوع تلقائي": type_c,
        })
        new_rows.append(row)
        label_items.append((
            {"barcode": barcode, "desc": desc.strip(), "color": color,
             "size": size, "price": price},
            q if copies_from_qty else 1,
        ))

    if clashes:
        st.error("هاي الألوان/المقاسات موجودة أصلًا بنفس الموديل، "
                 "عدّل كمياتها من صفحة «المخزون»:\n\n- " + "\n- ".join(clashes))
        return

    # ---------- الحفظ (مرة وحدة لكل الأسطر) ----------
    updated = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
    if save_data(updated):
        note = " | ".join(x for x in [
            f"فاتورة {invoice.strip()}" if invoice.strip() else "",
            f"طلبية {order.strip()}" if order.strip() else ""] if x) or "إدخال جديد"
        log_moves([{
            "نوع الحركة": MV_IN, "رقم الموديل": model, "الباركود": r["الباركود التلقائي"],
            "الوصف": desc.strip(), "اللون": r["اللون"], "المقاس": r["المقاس"],
            "الكمية": r["الكمية"], "ملاحظة": note,
        } for r in new_rows])
        if photo:
            save_model_image(model, photo)
        st.session_state["just_added"] = {
            "desc": desc.strip(),
            "items": label_items,
            "auto": bool(auto_print),
        }
        st.session_state["add_form_id"] = form_id + 1
        total_qty = sum(v[2] for v in variants)
        flash(f"تمت إضافة الموديل {model} — {len(variants)} لون/مقاس، "
              f"إجمالي {total_qty} قطعة.")
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
        o["المندوب"] = o["المندوب"].fillna("").astype(str)
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


def adjust_stock_many(changes, note=""):
    """changes = {barcode: delta}. يتحقق من الكل أولًا ثم يحفظ مرة وحدة ويسجّل الحركة."""
    data, _ = load_data()
    rows = {}
    for code, delta in changes.items():
        idx = data.index[data["الباركود التلقائي"] == str(code)]
        if len(idx) == 0:
            st.error(f"الباركود {code} غير موجود في المخزون.")
            return False
        i = idx[0]
        if float(data.loc[i, "الكمية"]) + delta < 0:
            st.error(f"الكمية غير كافية للباركود {code}.")
            return False
        rows[i] = delta
    log = []
    for i, delta in rows.items():
        if delta:
            log.append(move_row(MV_OUT if delta < 0 else MV_RET, data.loc[i], delta, note))
        data.loc[i, "الكمية"] = float(data.loc[i, "الكمية"]) + delta
    if save_data(data):
        log_moves(log)
        return True
    return False


def adjust_stock(barcode, delta, note=""):
    return adjust_stock_many({barcode: delta}, note)


def update_order(orders, i, status, restock=False):
    row = orders.loc[i]
    if restock and not row["أُعيد للمخزون"]:
        if not adjust_stock(row["الباركود"], float(row["الكمية"]),
                            f"إرجاع طلب رقم {int(row['رقم الطلب'])}"):
            return
        orders.loc[i, "أُعيد للمخزون"] = True
    orders.loc[i, "الحالة"] = status
    if save_orders(orders):
        flash(f"طلب رقم {int(row['رقم الطلب'])}: {status}")
        st.rerun()


def bulk_update(orders, idxs, status):
    orders.loc[idxs, "الحالة"] = status
    if save_orders(orders):
        flash(f"تم تحديث {len(idxs)} سطر إلى: {status}")
        st.rerun()


def page_delivery():
    require_admin()
    page_header("التوصيل", "اجمع كل القطع اللي مع المندوب في شحنة وحدة، وبعدين اضغط «شحن».")
    orders = load_orders()
    cart = st.session_state.setdefault("ship_cart", [])
    tab_new, tab_open, tab_log = st.tabs(["➕ شحنة جديدة", "🚚 الشحنات المفتوحة", "📋 السجل"])

    # ---------------- شحنة جديدة ----------------
    with tab_new:
        with st.form("add_to_cart", clear_on_submit=True):
            c1, c2 = st.columns([2, 1])
            barcode = c1.text_input("الباركود")
            q_ = c2.number_input("الكمية", min_value=1, value=1, step=1)
            add = st.form_submit_button("➕ أضف للشحنة", type="primary",
                                        use_container_width=True)
        if add:
            code = barcode.strip()
            hit = df[df["الباركود التلقائي"] == code]
            if not code or hit.empty:
                st.error("الباركود غير موجود.")
            else:
                stock = float(hit.iloc[0]["الكمية"])
                line = next((l for l in cart if l["barcode"] == code), None)
                already = line["qty"] if line else 0
                if already + q_ > stock:
                    st.error(f"المتوفر {stock:.0f} فقط (منها {already} مضافة للشحنة).")
                else:
                    if line:
                        line["qty"] += int(q_)
                    else:
                        cart.append({"barcode": code, "qty": int(q_)})
                    st.rerun()

        if not cart:
            st.info("الشحنة فاضية. أضف باركودات من فوق (الماسح بيرسل Enter تلقائيًا).")
        else:
            st.markdown("#### محتويات الشحنة")
            lines, total_pcs, total_cash = [], 0, 0.0
            for n, l in enumerate(cart):
                hit = df[df["الباركود التلقائي"] == l["barcode"]]
                if hit.empty:
                    continue
                it = hit.iloc[0]
                price = pd.to_numeric(it["سعر البيع"], errors="coerce")
                price = 0.0 if pd.isna(price) else float(price)
                lines.append({"code": l["barcode"], "qty": l["qty"], "price": price, "item": it})
                total_pcs += l["qty"]
                total_cash += price * l["qty"]
                c1, c2, c3 = st.columns([5, 2, 1])
                c1.markdown(f"**{it['وصف']}** ({it['اللون']} / {it['المقاس']}) — `{l['barcode']}`")
                c2.markdown(f"× **{l['qty']}** = {price * l['qty']:,.2f}")
                if c3.button("🗑️", key=f"rm_{n}"):
                    cart.pop(n)
                    st.rerun()

            m1, m2 = st.columns(2)
            m1.metric("إجمالي القطع", total_pcs)
            m2.metric("الكاش المطلوب", f"{total_cash:,.2f}")

            driver = st.text_input("اسم المندوب (اختياري)",
                                   key=f"ship_driver_{st.session_state.get('ship_n', 0)}")
            b1, b2 = st.columns([3, 1])
            if b2.button("تفريغ الشحنة", use_container_width=True):
                st.session_state["ship_cart"] = []
                st.rerun()
            if b1.button("🚚 شحن مع المندوب", type="primary", use_container_width=True):
                changes = {}
                for x in lines:
                    changes[x["code"]] = changes.get(x["code"], 0) - x["qty"]
                if adjust_stock_many(changes, "شحنة مع مندوب"):
                    new_id = int(orders["رقم الطلب"].max()) + 1 if len(orders) else 1
                    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
                    rows = [{
                        "رقم الطلب": new_id, "التاريخ": stamp, "المندوب": driver.strip(),
                        "الباركود": x["code"], "الوصف": x["item"]["وصف"],
                        "اللون": x["item"]["اللون"], "المقاس": x["item"]["المقاس"],
                        "الكمية": x["qty"], "سعر البيع": x["price"],
                        "المبلغ المطلوب": x["price"] * x["qty"],
                        "الحالة": ST_OUT, "أُعيد للمخزون": False,
                    } for x in lines]
                    new_orders = pd.concat([orders, pd.DataFrame(rows)], ignore_index=True)
                    if save_orders(new_orders):
                        st.session_state["ship_cart"] = []
                        st.session_state["ship_n"] = st.session_state.get("ship_n", 0) + 1
                        flash(f"تم إنشاء الشحنة رقم {new_id} ({total_pcs} قطعة) وخصمها من المخزون.")
                        st.rerun()
                    else:  # تراجع عن الخصم إذا فشل حفظ الطلبات
                        adjust_stock_many({k: -v for k, v in changes.items()},
                                          "تراجع عن شحنة (فشل حفظ الطلبات)")

    # ---------------- الشحنات المفتوحة ----------------
    with tab_open:
        open_o = orders[orders["الحالة"].isin(OPEN_STATUSES)]
        if open_o.empty:
            st.info("لا توجد شحنات مفتوحة.")
        for oid in sorted(open_o["رقم الطلب"].unique(), reverse=True):
            grp = open_o[open_o["رقم الطلب"] == oid]
            out_idx = list(grp.index[grp["الحالة"] == ST_OUT])
            first = grp.iloc[0]
            with st.container(border=True):
                st.markdown(f"### 📦 شحنة #{int(oid)}")
                st.caption(f"{first['التاريخ']} | المندوب: {first['المندوب'] or '—'} | "
                           f"عدد القطع: {int(grp['الكمية'].sum())}")
                if out_idx:
                    cash = float(grp.loc[out_idx, "المبلغ المطلوب"].sum())
                    st.markdown(f"💵 الكاش المتوقع من المندوب: **{cash:,.2f}**")
                    a1, a2 = st.columns(2)
                    if a1.button("✅ تم توصيل كل الأسطر", key=f"alld_{oid}",
                                 use_container_width=True):
                        bulk_update(orders, out_idx, ST_DELIVERED)
                    if a2.button("❌ إلغاء كل الأسطر قيد التوصيل", key=f"allc_{oid}",
                                 use_container_width=True):
                        bulk_update(orders, out_idx, ST_CAN_WAIT)

                for i, r in grp.iterrows():
                    st.divider()
                    st.markdown(
                        f"**{r['الوصف']}** ({r['اللون']} / {r['المقاس']}) × **{int(r['الكمية'])}**  \n"
                        f"الباركود: {r['الباركود']} | المبلغ: **{r['المبلغ المطلوب']:,.2f}** | "
                        f"الحالة: **{r['الحالة']}**"
                    )

                    if r["الحالة"] == ST_OUT:
                        b1, b2, b3 = st.columns(3)
                        if b1.button("✅ تم التوصيل", key=f"d{i}", use_container_width=True):
                            update_order(orders, i, ST_DELIVERED)
                        if b2.button("↩️ مرتجع جزئي", key=f"r_part_{i}", use_container_width=True):
                            st.session_state[f"editing_order_{i}"] = True
                        if b3.button("❌ ملغي", key=f"c{i}", use_container_width=True):
                            update_order(orders, i, ST_CAN_WAIT)

                        if st.session_state.get(f"editing_order_{i}", False):
                            with st.form(f"form_edit_order_{i}"):
                                st.caption("حدد عدد القطع اللي الزبون أخدها فعلًا، والباقي بينسجل مرتجع.")
                                total_q = int(r["الكمية"])
                                sold_qty = st.number_input("القطع المباعة", min_value=0,
                                                           max_value=total_q, value=total_q)
                                ok = st.form_submit_button("تأكيد", type="primary")
                            if ok:
                                price = float(r["سعر البيع"])
                                ret_q = total_q - sold_qty
                                if sold_qty == 0:
                                    orders.loc[i, ["المبلغ المطلوب", "الحالة"]] = [0.0, ST_RET_WAIT]
                                else:
                                    orders.loc[i, ["الكمية", "المبلغ المطلوب", "الحالة"]] = [
                                        sold_qty, sold_qty * price, ST_DELIVERED]
                                    if ret_q > 0:
                                        extra = r.to_dict()
                                        extra.update({"الكمية": ret_q, "المبلغ المطلوب": 0.0,
                                                      "الحالة": ST_RET_WAIT, "أُعيد للمخزون": False})
                                        orders = pd.concat([orders, pd.DataFrame([extra])],
                                                           ignore_index=True)
                                if save_orders(orders):
                                    st.session_state[f"editing_order_{i}"] = False
                                    flash(f"تم تعديل السطر: مباع {sold_qty} ومرتجع {ret_q} (بانتظار الاستلام).")
                                    st.rerun()

                    elif r["الحالة"] == ST_DELIVERED:
                        st.success(f"💵 كاش مطلوب تحصيله: {r['المبلغ المطلوب']:,.2f}")
                        if st.button("↩️ تحويل إلى مرتجع", key=f"r_del_{i}"):
                            update_order(orders, i, ST_RET_WAIT)
                    else:
                        done = ST_RET_DONE if r["الحالة"] == ST_RET_WAIT else ST_CAN_DONE
                        st.warning("بانتظار استلام القطع من المندوب.")
                        if st.checkbox("📥 استلمت القطع وأدخلتها للمستودع", key=f"k{i}"):
                            update_order(orders, i, done, restock=not r["أُعيد للمخزون"])

    # ---------------- السجل ----------------
    with tab_log:
        if orders.empty:
            st.info("لا توجد طلبات بعد.")
        else:
            delivered = orders[orders["الحالة"] == ST_DELIVERED]
            m1, m2, m3 = st.columns(3)
            m1.metric("💵 إجمالي الكاش المسلّم", f"{delivered['المبلغ المطلوب'].sum():,.2f}")
            m2.metric("🚚 أسطر قيد التوصيل", int((orders["الحالة"] == ST_OUT).sum()))
            m3.metric("📦 عدد الشحنات", int(orders["رقم الطلب"].nunique()))
            f = st.selectbox("الحالة", ["الكل"] + sorted(orders["الحالة"].unique()), key="ord_f")
            shown = orders if f == "الكل" else orders[orders["الحالة"] == f]
            shown = shown.drop(columns=["أُعيد للمخزون", "كلفة الوحدة"]).iloc[::-1]
            st.dataframe(shown, use_container_width=True, hide_index=True)
            download_buttons(shown, "طلبات_التوصيل", "orders")


# =========================
# طباعة الباركود
# =========================
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
               "للطابعة الحرارية: الهوامش «بدون»، المقياس 100%، وحجم الورق نفس مقاس الملصق (40×30).")


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
# كشف المواد (مختصر / تفصيلي)
# =========================
def _model_sort_key(m):
    m = str(m)
    return (0, int(m), "") if m.isdigit() else (1, 0, m)


def _fmt_q(x):
    x = float(x)
    return f"{int(x):,}" if x == int(x) else f"{x:,.2f}"


def build_statement(data, moves, detailed, d_from=None, d_to=None, term="", hide_out=False,
                    breakdown=False):
    """يرجّع (blocks, flat). blocks = قائمة موديلات، flat = جدول مسطّح للتنزيل."""
    d = data.copy()
    d["_q"] = qty(d)
    d["_m"] = d["رقم الموديل"].replace("", "بدون موديل")

    mv = moves.copy()
    if not mv.empty:
        mv["_m"] = mv["رقم الموديل"].replace("", "بدون موديل")
        bar_model = dict(zip(d["الباركود التلقائي"], d["_m"]))
        mv["_m"] = [bar_model.get(b, m) for b, m in zip(mv["الباركود"], mv["_m"])]
    else:
        mv["_m"] = []

    date_filter = d_from is not None and d_to is not None
    if date_filter and not mv.empty:
        ds = mv["التاريخ"].str[:10]
        mv = mv[(ds >= d_from.isoformat()) & (ds <= d_to.isoformat())]

    words = term.lower().split()
    blocks = []
    for model in sorted(d["_m"].unique(), key=_model_sort_key):
        g = d[d["_m"] == model]
        descs = [x for x in g["وصف"] if x]
        desc = descs[0] if descs else ""
        total = float(g["_q"].sum())
        hay = " ".join([str(model), desc, *g["نوع الصنف"].unique(), *g["الماركة"].unique()]).lower()
        if words and not all(w in hay for w in words):
            continue
        if hide_out and total <= 0:
            continue

        rows = []
        if detailed:
            ev = mv[mv["_m"] == model]
            if not date_filter:
                # رصيد الأصناف قبل بدء السجل (ما إله تاريخ مسجّل)
                logged = mv.groupby("الباركود")["الكمية"].sum().to_dict() if not mv.empty else {}
                for _, r in g.iterrows():
                    opening = float(r["_q"]) - float(logged.get(r["الباركود التلقائي"], 0))
                    if abs(opening) > 1e-9:
                        rows.append({"date": "", "type": "رصيد سابق", "color": r["اللون"],
                                     "size": r["المقاس"], "qty": opening, "signed": False,
                                     "barcode": r["الباركود التلقائي"], "user": "", "note": ""})
            for _, r in ev.sort_values("التاريخ", kind="stable").iterrows():
                rows.append({"date": r["التاريخ"], "type": r["نوع الحركة"], "color": r["اللون"],
                             "size": r["المقاس"], "qty": float(r["الكمية"]), "signed": True,
                             "barcode": r["الباركود"], "user": r["المستخدم"], "note": r["ملاحظة"]})
            if not breakdown:
                grouped = {}
                for r in rows:
                    k = (r["date"], r["type"], r["user"], r["note"], r["signed"])
                    if k in grouped:
                        grouped[k]["qty"] += r["qty"]
                    else:
                        grouped[k] = {**r, "color": "", "size": "", "barcode": ""}
                rows = list(grouped.values())
            if date_filter and not rows:
                continue
        blocks.append({"model": model, "desc": desc, "total": total, "rows": rows})

    if detailed and not breakdown:
        flat_rows = [{
            "رقم الموديل": b["model"], "الوصف": b["desc"],
            "التاريخ": r["date"] or "قبل السجل", "الحركة": r["type"],
            "الكمية": int(r["qty"]) if r["qty"] == int(r["qty"]) else r["qty"],
            "المستخدم": r["user"], "ملاحظة": r["note"]}
            for b in blocks for r in b["rows"]]
        flat = pd.DataFrame(flat_rows, columns=[
            "رقم الموديل", "الوصف", "التاريخ", "الحركة", "الكمية", "المستخدم", "ملاحظة"])
    elif detailed:
        flat_rows = []
        for b in blocks:
            for r in b["rows"]:
                flat_rows.append({
                    "رقم الموديل": b["model"], "الوصف": b["desc"],
                    "التاريخ": r["date"] or "قبل السجل", "الحركة": r["type"],
                    "اللون": r["color"], "المقاس": r["size"],
                    "الكمية": int(r["qty"]) if r["qty"] == int(r["qty"]) else r["qty"],
                    "الباركود": r["barcode"], "المستخدم": r["user"], "ملاحظة": r["note"]})
        flat = pd.DataFrame(flat_rows, columns=[
            "رقم الموديل", "الوصف", "التاريخ", "الحركة", "اللون", "المقاس",
            "الكمية", "الباركود", "المستخدم", "ملاحظة"])
    else:
        flat = pd.DataFrame(
            [{"رقم الموديل": b["model"], "الوصف": b["desc"], "الكمية": int(round(b["total"]))}
             for b in blocks], columns=["رقم الموديل", "الوصف", "الكمية"])
    return blocks, flat


STATEMENT_CSS = """
body{font-family:Cairo,Tahoma,Arial,sans-serif;direction:rtl;margin:0;padding:14px;color:#1E2633;background:#fff}
.bar{display:flex;gap:12px;align-items:center;margin-bottom:12px}
.bar button{font-family:inherit;font-weight:700;padding:8px 18px;border-radius:10px;
  border:1px solid #0D1B2E;background:#fff;cursor:pointer}
.head{border-bottom:2px solid #0D1B2E;margin-bottom:10px;padding-bottom:6px}
.head h1{margin:0;font-size:20px}
.head p{margin:2px 0 0;color:#555;font-size:13px}
table{width:100%;border-collapse:collapse;font-size:13px;margin-bottom:6px}
th{background:#0D1B2E;color:#fff;padding:5px 8px;text-align:right}
td{border-bottom:1px solid #ddd;padding:4px 8px;text-align:right}
tfoot td{font-weight:800;background:#f3ece1;border-top:2px solid #0D1B2E}
tr{break-inside:avoid}
.model{margin:14px 0 4px;break-inside:auto}
.mh{background:#f3ece1;border:1px solid #d9c9ad;padding:6px 10px;font-weight:700;
  display:flex;justify-content:space-between;break-after:avoid}
.num{font-variant-numeric:tabular-nums}
.dim{color:#888}
.pos{color:#1b7a3a;font-weight:700}.neg{color:#b3261e;font-weight:700}
@media print{.bar{display:none}body{padding:0}thead{display:table-header-group}}
"""


def statement_html(blocks, detailed, subtitle="", breakdown=False):
    e = htmllib.escape
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    user = st.session_state.get("ovvo_username", "")
    kind = "تفصيلي" if detailed else "مختصر"
    grand = sum(b["total"] for b in blocks)
    out = [f'<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8">'
           f'<style>{STATEMENT_CSS}</style></head><body>',
           f'<div class="bar"><button onclick="window.print()">🖨️ طباعة الكشف</button>'
           f'<span>{len(blocks)} موديل</span></div>',
           f'<div class="head"><h1>OVVO STORE — كشف المواد ({kind})</h1>'
           f'<p>تاريخ الطباعة: {e(now)}'
           f'{" — المستخدم: " + e(user) if user else ""}'
           f'{" — " + e(subtitle) if subtitle else ""}</p></div>']

    if not blocks:
        out.append("<p>لا توجد نتائج.</p></body></html>")
        return "".join(out)

    if not detailed:
        out.append("<table><thead><tr><th>#</th><th>رقم الموديل</th><th>الوصف</th>"
                   "<th>الكمية</th></tr></thead><tbody>")
        for n, b in enumerate(blocks, 1):
            out.append(f'<tr><td class="dim">{n}</td><td>{e(str(b["model"]))}</td>'
                       f'<td>{e(b["desc"])}</td><td class="num">{_fmt_q(b["total"])}</td></tr>')
        out.append(f'</tbody><tfoot><tr><td colspan="3">الإجمالي ({len(blocks)} موديل)</td>'
                   f'<td class="num">{_fmt_q(grand)}</td></tr></tfoot></table>')
    else:
        for b in blocks:
            heads = (["التاريخ", "الحركة", "اللون", "المقاس", "الكمية", "الباركود", "المستخدم", "ملاحظة"]
                     if breakdown else ["التاريخ", "الحركة", "الكمية", "المستخدم", "ملاحظة"])
            out.append(f'<div class="model"><div class="mh"><span>موديل {e(str(b["model"]))}'
                       f' — {e(b["desc"])}</span><span>الكمية الحالية: {_fmt_q(b["total"])} حبة</span></div>'
                       '<table><thead><tr>' + "".join(f"<th>{h}</th>" for h in heads) + '</tr></thead><tbody>')
            if not b["rows"]:
                out.append(f'<tr><td colspan="{len(heads)}" class="dim">لا توجد حركات مسجّلة.</td></tr>')
            for r in b["rows"]:
                q = r["qty"]
                if r["signed"]:
                    cls = "pos" if q > 0 else "neg" if q < 0 else ""
                    qtxt = f'<span dir="ltr" class="{cls}">{"+" if q > 0 else ""}{_fmt_q(q)}</span>'
                else:
                    qtxt = f'<span dir="ltr">{_fmt_q(q)}</span>'
                if not breakdown:
                    qtxt += " حبة"
                date_td = f'<td class="num">{e(r["date"]) or "<span class=dim>قبل السجل</span>"}</td>'
                if breakdown:
                    out.append(
                        f'<tr>{date_td}<td>{e(r["type"])}</td><td>{e(r["color"]) or "—"}</td>'
                        f'<td>{e(r["size"]) or "—"}</td><td class="num">{qtxt}</td>'
                        f'<td class="num">{e(r["barcode"])}</td><td>{e(r["user"])}</td>'
                        f'<td>{e(r["note"])}</td></tr>')
                else:
                    out.append(
                        f'<tr>{date_td}<td>{e(r["type"])}</td><td class="num">{qtxt}</td>'
                        f'<td>{e(r["user"])}</td><td>{e(r["note"])}</td></tr>')
            out.append("</tbody></table></div>")
        out.append(f'<table><tfoot><tr><td>الإجمالي الحالي ({len(blocks)} موديل)</td>'
                   f'<td class="num" style="width:120px">{_fmt_q(grand)}</td></tr></tfoot></table>')
    out.append("</body></html>")
    return "".join(out)


def page_statement():
    page_header("كشف المواد",
                "كشف مختصر (رقم الموديل، الوصف، الكمية) أو تفصيلي مع تاريخ كل إدخال وحركة. اطبعه أو نزّله.")
    if df.empty:
        st.info("لا توجد أصناف بعد.")
        return

    c1, c2, c3 = st.columns([1.2, 1.8, 1.2])
    kind = c1.radio("نوع الكشف", ["مختصر", "تفصيلي"], horizontal=True, key="stm_kind")
    term = c2.text_input("بحث برقم الموديل أو الوصف", key="stm_q")
    hide_out = c3.checkbox("إخفاء الموديلات النافدة", key="stm_hide")
    detailed = kind == "تفصيلي"

    d_from = d_to = None
    subtitle = ""
    breakdown = False
    moves = load_moves()
    if detailed:
        breakdown = st.checkbox("إظهار تفاصيل الألوان والمقاسات لكل إدخال", key="stm_break")
        use_dates = st.checkbox("تحديد فترة زمنية", key="stm_dates")
        if use_dates:
            f1, f2 = st.columns(2)
            d_from = f1.date_input("من تاريخ", value=date.today() - timedelta(days=30), key="stm_from")
            d_to = f2.date_input("إلى تاريخ", value=date.today(), key="stm_to")
            if d_from > d_to:
                st.error("تاريخ البداية بعد تاريخ النهاية.")
                return
            subtitle = f"الفترة: {d_from.isoformat()} إلى {d_to.isoformat()}"
        if moves.empty:
            st.info("سجل الحركات بيبدأ من هالتحديث: الأصناف الموجودة من قبل بتظهر كـ«رصيد سابق» "
                    "بدون تاريخ، وكل إدخال أو تعديل جديد بيتسجّل بتاريخه ووقته.")

    blocks, flat = build_statement(df, moves, detailed, d_from, d_to, term, hide_out, breakdown)
    if not blocks:
        st.info("لا توجد نتائج مطابقة.")
        return
    st.caption(f"عدد الموديلات: {len(blocks):,} — إجمالي القطع الحالية: "
               f"{sum(b['total'] for b in blocks):,.0f}")
    components.html(statement_html(blocks, detailed, subtitle, breakdown), height=640, scrolling=True)
    download_buttons(flat, f"كشف_المواد_{kind}", "stmt")


# =========================
# تقرير المخزون الكامل (مختصر بالموديل / تفصيلي بالألوان والمقاسات)
# =========================
MODEL_INFO_COLUMNS = ["الوصف", "وصف 2", "نوع الصنف", "الماركة", "المنشاة", "الموسم",
                      "سنة الشراء", "الطلبية", "فاتورة رقم", "محجب او غير محجب",
                      "كت او كم او نص كم", "كود الصنف", "سعر البيع", "سعر البيع 2", "ضريبة"]


def stock_report_tables(data):
    """يرجّع (مختصر، تفصيلي). المختصر سطر لكل موديل بكل معلوماته بدون ألوان ومقاسات."""
    d = data.copy()
    d["_q"] = qty(d)
    d["_key"] = [m if m else f"بدون موديل #{b}" for m, b in zip(d["رقم الموديل"], d["الباركود التلقائي"])]
    rows = []
    for key in sorted(d["_key"].unique(), key=_model_sort_key):
        g = d[d["_key"] == key]
        r = {"رقم الموديل": key if g["رقم الموديل"].iloc[0] else "بدون موديل"}
        for c in MODEL_INFO_COLUMNS:
            col = "وصف" if c == "الوصف" else c
            vals = g[col]
            if col in NUM_COLUMNS:
                vals = vals.dropna()
                r[c] = vals.iloc[0] if len(vals) else ""
            else:
                vals = [v for v in vals if v]
                r[c] = vals[0] if vals else ""
        r["الكمية"] = int(round(float(g["_q"].sum())))
        rows.append(r)
    cols = ["رقم الموديل"] + MODEL_INFO_COLUMNS[:12] + ["سعر البيع", "سعر البيع 2", "ضريبة", "الكمية"]
    cols = list(dict.fromkeys(cols))
    summary = pd.DataFrame(rows, columns=cols)

    det = d.assign(_s=[size_key(x) for x in d["المقاس"]])
    det = det.sort_values(by=["_key", "اللون", "_s"],
                          key=lambda c: c.map(_model_sort_key) if c.name == "_key" else c,
                          kind="stable")
    det_cols = ["رقم الموديل"] + [c for c in VISIBLE_COLUMNS if c != "رقم الموديل"]
    detailed = det[det_cols].reset_index(drop=True)
    return summary, detailed


def table_print_html(data, title, subtitle=""):
    e = htmllib.escape
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    def cell(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return ""
        if isinstance(v, (int, float, np.integer, np.floating)):
            return _fmt_q(v)
        return e(str(v))

    head = "".join(f"<th>{e(str(c))}</th>" for c in data.columns)
    body = "".join(
        "<tr>" + "".join(f"<td>{cell(v)}</td>" for v in row) + "</tr>"
        for row in data.itertuples(index=False, name=None))
    return (
        f'<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8">'
        f'<style>{STATEMENT_CSS}@page{{size:landscape;margin:8mm}}'
        f'table{{font-size:11px}}</style></head><body>'
        f'<div class="bar"><button onclick="window.print()">🖨️ طباعة</button>'
        f'<span>{len(data):,} سطر</span></div>'
        f'<div class="head"><h1>OVVO STORE — {e(title)}</h1>'
        f'<p>تاريخ الطباعة: {e(now)}{" — " + e(subtitle) if subtitle else ""}</p></div>'
        f'<table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></body></html>')


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
        summary_t, detailed_t = stock_report_tables(df)
        rt1, rt2 = st.columns([1.3, 2])
        rkind = rt1.radio("نوع التقرير", ["مختصر (حسب الموديل)", "تفصيلي (بالألوان والمقاسات)"],
                          key="full_kind")
        is_short = rkind.startswith("مختصر")
        table = summary_t if is_short else detailed_t
        title = "المخزون الكامل — " + ("مختصر حسب الموديل" if is_short else "تفصيلي")
        stamp = datetime.now().strftime("%Y-%m-%d")
        b1, b2, b3, _ = st.columns([1, 1, 1, 2])
        b1.download_button(
            "⬇️ Excel", to_excel_bytes(table), f"المخزون_{'مختصر' if is_short else 'تفصيلي'}_{stamp}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="full_xlsx", use_container_width=True)
        b2.download_button(
            "⬇️ CSV", table.to_csv(index=False).encode("utf-8-sig"),
            f"المخزون_{'مختصر' if is_short else 'تفصيلي'}_{stamp}.csv", "text/csv",
            key="full_csv", use_container_width=True)
        show_print = b3.toggle("🖨️ معاينة وطباعة", key="full_print")
        if is_short:
            st.caption(f"{len(table):,} موديل — إجمالي القطع {int(table['الكمية'].sum()):,}. "
                       "كل موديل بسطر واحد بكل معلوماته بدون تجزئة للألوان والمقاسات.")
        else:
            st.caption(f"{len(table):,} سطر (لون/مقاس) بكل المعلومات المدخلة.")
        if show_print:
            components.html(table_print_html(table, title), height=640, scrolling=True)
        else:
            st.dataframe(table, use_container_width=True, hide_index=True)

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
            r1.metric("📦 إجمالي الأسطر", f"{total_orders}")
            r2.metric("✅ تم التوصيل", f"{delivered_orders}")
            r3.metric("🚚 قيد التوصيل", f"{out_orders}")
            r4.metric("❌ الملغاة / المرتجعة", f"{cancelled_orders + returned_orders}")

            st.markdown("---")
            c1, c2 = st.columns(2)
            c1.metric("💵 إجمالي الكاش المُستلَم (المحصل)", f"{collected_cash:,.2f}")
            c2.metric("⏳ الكاش المعلق (عند المندوب/قيد التوصيل)", f"{pending_cash:,.2f}")

            st.markdown("#### تفاصيل السجل الكامل للحالات")
            shown_ord = orders_rep.drop(columns=["أُعيد للمخزون", "كلفة الوحدة"]).iloc[::-1]
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
            role_label = st.selectbox("الصلاحية", ["مشاهد", "محاسب", "مدير"])
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
                    "role": ROLE_BY_LABEL[role_label]
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
# المحاسبة (دور «محاسب»)
# الصق هذا البلوك كامل قبل سطر:  # التوجيه
# =========================
LEDGER_FILE = os.path.join(BASE_DIR, "OVVO_ledger.xlsx")   # المصاريف والقبوض وأجور التوصيل (لا تُحذف، فقط تُلغى بسبب)
AUDIT_FILE = os.path.join(BASE_DIR, "OVVO_audit.xlsx")     # سجل الرقابة: من عمل إيش ومتى
ACC_ROLES = ("accountant",)   # الكلفة والأرباح للمحاسب فقط

LEDGER_COLS = ["رقم القيد", "التاريخ", "النوع", "الفئة", "المرجع", "المبلغ", "طريقة الدفع",
               "البيان", "المستخدم", "وقت الإدخال", "الحالة", "سبب الإلغاء"]
AUDIT_COLS = ["الوقت", "المستخدم", "الإجراء", "التفاصيل"]
LG_EXP, LG_FEE, LG_RCPT, LG_SUP = "مصروف", "أجرة توصيل", "قبض من مندوب", "دفع لمورد"
LG_ACTIVE, LG_VOID = "ساري", "ملغى"
PAY_CASH, PAY_BANK, PAY_DRV = "نقدي (صندوق)", "بنك / محفظة", "خصم من كاش المندوب"
EXP_CATS = ["إيجار", "رواتب وأجور", "كهرباء وماء", "تسويق وإعلان", "تغليف", "مواصلات",
            "صيانة", "رسوم وعمولات", "أخرى"]
NO_DRIVER = "بدون اسم مندوب"

# شجرة الحسابات
A_CASH, A_BANK, A_DRV = "الصندوق", "البنك والمحافظ", "ذمم المندوبين (كاش عند المندوب)"
A_INV, A_AP = "المخزون", "الموردون"
A_REV, A_VAT = "إيرادات المبيعات", "ضريبة المبيعات المستحقة"
A_COGS, A_FEE, A_ADJ = "تكلفة البضاعة المباعة", "مصروف أجور التوصيل", "فروقات وتسويات المخزون"
JCOLS = ["رقم القيد", "التاريخ", "المصدر", "الحساب", "مدين", "دائن", "البيان", "المرجع"]
SHIP_COLS = ["رقم الشحنة", "التاريخ", "المندوب", "القطع المشحونة", "قطع مُسلّمة",
             "المبيعات (شامل الضريبة)", "المبيعات الصافية", "التكلفة", "أجرة التوصيل",
             "ربح الشحنة", "الحالة"]


def require_accounting():
    if current_role not in ACC_ROLES:
        st.error("هذه الصفحة للمحاسب فقط.")
        st.stop()


def _num_or_none(v):
    v = pd.to_numeric(v, errors="coerce")
    return None if pd.isna(v) or v <= 0 else float(v)


def _cost_map(data):
    c = pd.to_numeric(data["الكلفة النهائية"], errors="coerce")
    return {b: float(x) for b, x in zip(data["الباركود التلقائي"], c) if not pd.isna(x)}


# ---------- قراءة/كتابة الجداول (مع نسخة احتياطية) ----------
def _read_table(path, cols):
    empty = pd.DataFrame(columns=cols)
    if not os.path.isfile(path):
        return empty
    try:
        t = pd.read_excel(path, dtype=str)
    except Exception as e:
        st.error(f"تعذّرت قراءة {os.path.basename(path)}: {e}")
        return empty
    for c in cols:
        if c not in t.columns:
            t[c] = ""
    t = t[cols].fillna("").astype(str)
    return t.apply(lambda s: s.str.strip())


def _write_table(t, path, prefix):
    tmp = os.path.join(BASE_DIR, "_tmp_" + os.path.basename(path))
    try:
        if os.path.isfile(path):
            os.makedirs(BACKUP_DIR, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            shutil.copy2(path, os.path.join(BACKUP_DIR, f"{prefix}_{stamp}.xlsx"))
            old = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith(prefix + "_"))
            for f in old[:-MAX_BACKUPS]:
                os.remove(os.path.join(BACKUP_DIR, f))
        t.to_excel(tmp, index=False)
        os.replace(tmp, path)
        return True
    except PermissionError:
        st.error(f"تعذّر الحفظ: ملف {os.path.basename(path)} مفتوح في Excel. أغلقه ثم أعد المحاولة.")
    except Exception as e:
        st.error(f"تعذّر الحفظ: {e}")
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
    return False


def audit(action, details):
    cur = _read_table(AUDIT_FILE, AUDIT_COLS)
    row = pd.DataFrame([{
        "الوقت": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "المستخدم": st.session_state.get("ovvo_username", ""),
        "الإجراء": action, "التفاصيل": details}], columns=AUDIT_COLS)
    _write_table(pd.concat([cur, row], ignore_index=True), AUDIT_FILE, "AUDIT")


def load_ledger():
    t = _read_table(LEDGER_FILE, LEDGER_COLS)
    t["المبلغ"] = pd.to_numeric(t["المبلغ"], errors="coerce").fillna(0.0)
    t["الحالة"] = t["الحالة"].replace("", LG_ACTIVE)
    t["المرجع"] = t["المرجع"].str.replace(r"^(\d+)\.0$", r"\1", regex=True)
    return t


def add_ledger_entry(kind, date_, cat, ref, amount, pay, memo):
    led = load_ledger()
    entry_id = f"L-{len(led) + 1:04d}"
    row = {"رقم القيد": entry_id, "التاريخ": date_.isoformat(), "النوع": kind, "الفئة": cat,
           "المرجع": str(ref), "المبلغ": round(float(amount), 2), "طريقة الدفع": pay,
           "البيان": memo.strip(),
           "المستخدم": st.session_state.get("ovvo_username", ""),
           "وقت الإدخال": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
           "الحالة": LG_ACTIVE, "سبب الإلغاء": ""}
    out = pd.concat([led, pd.DataFrame([row], columns=LEDGER_COLS)], ignore_index=True)
    if _write_table(out, LEDGER_FILE, "LEDGER"):
        audit("قيد جديد", f"{entry_id} | {kind} | {amount:,.2f} | {cat or ref} | {memo.strip()}")
        return entry_id
    return None


def void_ledger_entry(entry_id, reason):
    led = load_ledger()
    m = led["رقم القيد"] == entry_id
    if not m.any() or led.loc[m, "الحالة"].iloc[0] != LG_ACTIVE:
        st.error("القيد غير موجود أو ملغى مسبقًا.")
        return False
    led.loc[m, ["الحالة", "سبب الإلغاء"]] = [LG_VOID, reason.strip()]
    if _write_table(led, LEDGER_FILE, "LEDGER"):
        audit("إلغاء قيد", f"{entry_id} | السبب: {reason.strip()}")
        return True
    return False


# ---------- مبيعات / شحنات ----------
def sales_lines(orders, data):
    """أسطر الطلبات «تم التوصيل» مع الكلفة (لقطة وقت الشحن، وإلا الكلفة الحالية)."""
    o = orders[orders["الحالة"] == ST_DELIVERED].copy()
    o = orders[orders["الحالة"] == ST_DELIVERED].copy()
    if "تاريخ الحالة" not in o.columns:   # ملفات طلبات قديمة بدون العمود
        o["تاريخ الحالة"] = ""
    cm = _cost_map(data)
    cm = _cost_map(data)
    o["_unit_cost"] = o["الباركود"].map(cm).fillna(0.0)
    o["_cost_src"] = np.where(o["_unit_cost"] > 0, "الكلفة الحالية", "غير متوفرة")
    o["_qty"] = pd.to_numeric(o["الكمية"], errors="coerce").fillna(0)
    o["_amt"] = pd.to_numeric(o["المبلغ المطلوب"], errors="coerce").fillna(0)
    o["_cogs"] = o["_unit_cost"] * o["_qty"]
    st_date = o["تاريخ الحالة"].astype(str).str.strip()
    o["_date"] = np.where(st_date != "", st_date, o["التاريخ"].astype(str))
    o["_date"] = o["_date"].astype(str).str[:10]
    return o


def ship_summary(orders, data, ledger, vat):
    if orders.empty:
        return pd.DataFrame(columns=SHIP_COLS)
    sl = sales_lines(orders, data)
    fees = ledger[(ledger["الحالة"] == LG_ACTIVE) & (ledger["النوع"] == LG_FEE)]
    fee_map = fees.groupby("المرجع")["المبلغ"].sum().to_dict()
    rows = []
    for oid, g in orders.groupby("رقم الطلب"):
        s = sl[sl["رقم الطلب"] == oid]
        gross = float(s["_amt"].sum())
        net = gross / (1 + vat / 100)
        cogs = float(s["_cogs"].sum())
        fee = float(fee_map.get(str(int(oid)), 0.0))
        pend = g["الحالة"].isin([ST_OUT, ST_RET_WAIT, ST_CAN_WAIT]).any()
        rows.append({
            "رقم الشحنة": int(oid), "التاريخ": str(g["التاريخ"].iloc[0]),
            "المندوب": str(g["المندوب"].iloc[0]),
            "القطع المشحونة": int(pd.to_numeric(g["الكمية"], errors="coerce").fillna(0).sum()),
            "قطع مُسلّمة": int(s["_qty"].sum()),
            "المبيعات (شامل الضريبة)": round(gross, 2), "المبيعات الصافية": round(net, 2),
            "التكلفة": round(cogs, 2), "أجرة التوصيل": round(fee, 2),
            "ربح الشحنة": round(net - cogs - fee, 2),
            "الحالة": "مفتوحة" if pend else "مغلقة"})
    return pd.DataFrame(rows, columns=SHIP_COLS).sort_values("رقم الشحنة", ascending=False)


def driver_balances(orders, ledger):
    o = orders.copy()
    o["_d"] = o["المندوب"].replace("", NO_DRIVER)
    o["_a"] = pd.to_numeric(o["المبلغ المطلوب"], errors="coerce").fillna(0)
    deliv = o[o["الحالة"] == ST_DELIVERED].groupby("_d")["_a"].sum()
    out = o[o["الحالة"] == ST_OUT].groupby("_d")["_a"].sum()
    act = ledger[ledger["الحالة"] == LG_ACTIVE]
    rc = act[act["النوع"] == LG_RCPT].groupby("المرجع")["المبلغ"].sum()
    drv_of = o.groupby("رقم الطلب")["_d"].first()
    ded = {}
    for _, r in act[(act["النوع"] == LG_FEE) & (act["طريقة الدفع"] == PAY_DRV)].iterrows():
        try:
            d = drv_of.get(int(float(r["المرجع"])), NO_DRIVER)
        except ValueError:
            d = NO_DRIVER
        ded[d] = ded.get(d, 0.0) + float(r["المبلغ"])
    names = sorted(set(deliv.index) | set(out.index) | set(rc.index) | set(ded))
    t = pd.DataFrame({
        "المندوب": names,
        "مبيعات مُسلّمة": [float(deliv.get(n, 0.0)) for n in names],
        "مقبوض منه": [float(rc.get(n, 0.0)) for n in names],
        "أجور مخصومة": [float(ded.get(n, 0.0)) for n in names],
        "قيد التوصيل (لم يُسلَّم بعد)": [float(out.get(n, 0.0)) for n in names]})
    t["الرصيد المستحق عليه"] = t["مبيعات مُسلّمة"] - t["مقبوض منه"] - t["أجور مخصومة"]
    return t.round(2)


# ---------- دفتر اليومية ----------
def build_journal(data, orders, moves, ledger, vat):
    J = []

    def add(ref, date_, source, memo, lines, link=""):
        for acct, dr, cr in lines:
            dr, cr = round(float(dr), 2), round(float(cr), 2)
            if dr or cr:
                J.append({"رقم القيد": ref, "التاريخ": str(date_)[:10], "المصدر": source,
                          "الحساب": acct, "مدين": dr, "دائن": cr, "البيان": memo, "المرجع": link})

    cm = _cost_map(data)

    # 1) المبيعات (عند «تم التوصيل») + تكلفتها
    for i, r in sales_lines(orders, data).iterrows():
        amt = round(float(r["_amt"]), 2)
        net = round(amt / (1 + vat / 100), 2)
        tax = round(amt - net, 2)
        oid = int(float(r["رقم الطلب"]))
        memo = f"بيع {r['الوصف']} ({r['اللون']}/{r['المقاس']}) × {int(r['_qty'])}"
        link = f"شحنة {oid}"
        add(f"S-{oid}-{i}", r["_date"], "مبيعات", memo,
            [(A_DRV, amt, 0), (A_REV, 0, net), (A_VAT, 0, tax)], link)
        c = round(float(r["_cogs"]), 2)
        add(f"C-{oid}-{i}", r["_date"], "تكلفة المبيعات", memo, [(A_COGS, c, 0), (A_INV, 0, c)], link)

    # 2) حركات المخزون: إدخال (مشتريات) + تسويات يدوية
    for n, m in moves.reset_index(drop=True).iterrows():
        kind = m["نوع الحركة"]
        if kind not in (MV_IN, MV_ADJ, MV_DEL):
            continue
        q = float(m["الكمية"])
        unit = cm.get(m["الباركود"])
        if unit is None or q == 0:
            continue
        v = round(abs(q) * unit, 2)
        memo = f"{kind}: {m['الوصف']} ({m['اللون']}/{m['المقاس']}) {q:+g}"
        if kind == MV_IN:
            add(f"P-{n + 1}", m["التاريخ"], "إدخال مخزون", memo, [(A_INV, v, 0), (A_AP, 0, v)], m["الباركود"])
        elif q > 0:
            add(f"A-{n + 1}", m["التاريخ"], "تسوية مخزون", memo, [(A_INV, v, 0), (A_ADJ, 0, v)], m["الباركود"])
        else:
            add(f"A-{n + 1}", m["التاريخ"], "تسوية مخزون", memo, [(A_ADJ, v, 0), (A_INV, 0, v)], m["الباركود"])

    # 3) المصاريف / أجور التوصيل / القبض من المندوب / الدفع للموردين
    pay_acc = {PAY_CASH: A_CASH, PAY_BANK: A_BANK, PAY_DRV: A_DRV}
    for _, r in ledger[ledger["الحالة"] == LG_ACTIVE].iterrows():
        amt = round(float(r["المبلغ"]), 2)
        pa = pay_acc.get(r["طريقة الدفع"], A_CASH)
        k = r["النوع"]
        memo = r["البيان"] or r["الفئة"] or k
        link = r["المرجع"]
        if k == LG_EXP:
            lines, src = [(f"مصروف - {r['الفئة'] or 'أخرى'}", amt, 0), (pa, 0, amt)], "مصاريف"
        elif k == LG_FEE:
            lines, src, link = [(A_FEE, amt, 0), (pa, 0, amt)], "أجور توصيل", f"شحنة {link}"
        elif k == LG_RCPT:
            lines, src, link = [(pa, amt, 0), (A_DRV, 0, amt)], "قبض من مندوب", f"مندوب: {link}"
        elif k == LG_SUP:
            lines, src, link = [(A_AP, amt, 0), (pa, 0, amt)], "دفع لمورد", f"مورد: {link}"
        else:
            continue
        add(r["رقم القيد"], r["التاريخ"], src, memo, lines, link)

    j = pd.DataFrame(J, columns=JCOLS)
    return j.sort_values("التاريخ", kind="stable").reset_index(drop=True)


def acct_kind(a):
    if a == A_REV:
        return "إيراد"
    if a in (A_COGS, A_FEE, A_ADJ) or a.startswith("مصروف"):
        return "مصروف"
    if a in (A_VAT, A_AP):
        return "التزام"
    return "أصل"


def _period(t, col, d_from, d_to):
    if t.empty:
        return t
    ds = t[col].astype(str).str[:10]
    return t[(ds >= d_from.isoformat()) & (ds <= d_to.isoformat())]


def _bal(j, acct, credit=True):
    s = j[j["الحساب"] == acct]
    dr, cr = float(s["مدين"].sum()), float(s["دائن"].sum())
    return cr - dr if credit else dr - cr


# ---------- فحوصات الرقابة ----------
def _sl_view(sl):
    v = sl[["رقم الطلب", "الوصف", "اللون", "المقاس", "الباركود"]].copy()
    v["الكمية"], v["المبلغ"] = sl["_qty"], sl["_amt"]
    v["التكلفة"], v["مصدر الكلفة"] = sl["_cogs"], sl["_cost_src"]
    return v


def control_checks(data, orders, moves, ledger, ships, vat, days, d_from, d_to):
    checks = []
    cm = _cost_map(data)
    cost = pd.to_numeric(data["الكلفة النهائية"], errors="coerce").fillna(0)
    checks.append(("أصناف بالمخزون بدون كلفة", "عالية",
                   data[(cost <= 0) & (qty(data) > 0)][["الباركود التلقائي", "وصف", "اللون", "المقاس", "الكمية"]]))
    sl = sales_lines(orders, data)
    sl = sl[(sl["_date"] >= d_from.isoformat()) & (sl["_date"] <= d_to.isoformat())] if len(sl) else sl
    checks.append(("مبيعات مُسلّمة بدون كلفة معروفة", "عالية", _sl_view(sl[sl["_unit_cost"] <= 0])))
    checks.append(("بيع بأقل من الكلفة", "عالية",
                   _sl_view(sl[(sl["_unit_cost"] > 0) & (sl["_amt"] < sl["_cogs"])])))
    checks.append(("مبيعات بسعر صفر", "عالية", _sl_view(sl[sl["_amt"] <= 0])))
    sp = _period(ships, "التاريخ", d_from, d_to)
    checks.append(("شحنات فيها مبيعات بدون أجرة توصيل مسجّلة", "متوسطة",
                   sp[(sp["قطع مُسلّمة"] > 0) & (sp["أجرة التوصيل"] <= 0)]))
    o = orders[orders["الحالة"].isin([ST_OUT, ST_RET_WAIT, ST_CAN_WAIT])].copy()
    if len(o):
        age = (pd.Timestamp.now() - pd.to_datetime(o["التاريخ"], errors="coerce")).dt.days
        o["أيام مفتوحة"] = age
        o = o[o["أيام مفتوحة"] >= days][["رقم الطلب", "التاريخ", "المندوب", "الوصف", "الكمية",
                                         "المبلغ المطلوب", "الحالة", "أيام مفتوحة"]]
    else:
        o = pd.DataFrame(columns=["رقم الطلب", "التاريخ", "المندوب", "الوصف", "الكمية",
                                  "المبلغ المطلوب", "الحالة", "أيام مفتوحة"])
    checks.append((f"شحنات/مرتجعات مفتوحة منذ {days} يوم أو أكثر", "عالية", o))
    adj = _period(moves[moves["نوع الحركة"].isin([MV_ADJ, MV_DEL])], "التاريخ", d_from, d_to).copy()
    adj["القيمة بالكلفة"] = [round(q * cm.get(b, 0.0), 2) for q, b in zip(adj["الكمية"], adj["الباركود"])]
    checks.append(("تعديلات/حذف كميات يدوية على المخزون", "متوسطة",
                   adj[["التاريخ", "نوع الحركة", "الوصف", "اللون", "المقاس", "الكمية",
                        "القيمة بالكلفة", "المستخدم", "ملاحظة"]]))
    voided = _period(ledger[ledger["الحالة"] == LG_VOID], "التاريخ", d_from, d_to)
    checks.append(("قيود ملغاة", "متوسطة",
                   voided[["رقم القيد", "التاريخ", "النوع", "المبلغ", "البيان", "المستخدم", "سبب الإلغاء"]]))
    db = driver_balances(orders, ledger)
    checks.append(("مندوب مقبوض منه أكثر من المستحق (رصيد سالب)", "عالية",
                   db[db["الرصيد المستحق عليه"] < -0.005]))
    return checks


# =========================
# المحاسبة (دور «محاسب»)
# الصق هذا البلوك كامل قبل سطر:  # التوجيه
# =========================
LEDGER_FILE = os.path.join(BASE_DIR, "OVVO_ledger.xlsx")   # المصاريف والقبوض وأجور التوصيل (لا تُحذف، فقط تُلغى بسبب)
AUDIT_FILE = os.path.join(BASE_DIR, "OVVO_audit.xlsx")     # سجل الرقابة: من عمل إيش ومتى
ACC_ROLES = ("accountant",)   # الكلفة والأرباح للمحاسب فقط

LEDGER_COLS = ["رقم القيد", "التاريخ", "النوع", "الفئة", "المرجع", "المبلغ", "طريقة الدفع",
               "البيان", "المستخدم", "وقت الإدخال", "الحالة", "سبب الإلغاء"]
AUDIT_COLS = ["الوقت", "المستخدم", "الإجراء", "التفاصيل"]
LG_EXP, LG_FEE, LG_RCPT, LG_SUP = "مصروف", "أجرة توصيل", "قبض من مندوب", "دفع لمورد"
LG_ACTIVE, LG_VOID = "ساري", "ملغى"
PAY_CASH, PAY_BANK, PAY_DRV = "نقدي (صندوق)", "بنك / محفظة", "خصم من كاش المندوب"
EXP_CATS = ["إيجار", "رواتب وأجور", "كهرباء وماء", "تسويق وإعلان", "تغليف", "مواصلات",
            "صيانة", "رسوم وعمولات", "أخرى"]
NO_DRIVER = "بدون اسم مندوب"

# شجرة الحسابات
A_CASH, A_BANK, A_DRV = "الصندوق", "البنك والمحافظ", "ذمم المندوبين (كاش عند المندوب)"
A_INV, A_AP = "المخزون", "الموردون"
A_REV, A_VAT = "إيرادات المبيعات", "ضريبة المبيعات المستحقة"
A_COGS, A_FEE, A_ADJ = "تكلفة البضاعة المباعة", "مصروف أجور التوصيل", "فروقات وتسويات المخزون"
JCOLS = ["رقم القيد", "التاريخ", "المصدر", "الحساب", "مدين", "دائن", "البيان", "المرجع"]
SHIP_COLS = ["رقم الشحنة", "التاريخ", "المندوب", "القطع المشحونة", "قطع مُسلّمة",
             "المبيعات (شامل الضريبة)", "المبيعات الصافية", "التكلفة", "أجرة التوصيل",
             "ربح الشحنة", "الحالة"]


def require_accounting():
    if current_role not in ACC_ROLES:
        st.error("هذه الصفحة للمحاسب فقط.")
        st.stop()


def _num_or_none(v):
    v = pd.to_numeric(v, errors="coerce")
    return None if pd.isna(v) or v <= 0 else float(v)


def _cost_map(data):
    c = pd.to_numeric(data["الكلفة النهائية"], errors="coerce")
    return {b: float(x) for b, x in zip(data["الباركود التلقائي"], c) if not pd.isna(x)}


# ---------- قراءة/كتابة الجداول (مع نسخة احتياطية) ----------
def _read_table(path, cols):
    empty = pd.DataFrame(columns=cols)
    if not os.path.isfile(path):
        return empty
    try:
        t = pd.read_excel(path, dtype=str)
    except Exception as e:
        st.error(f"تعذّرت قراءة {os.path.basename(path)}: {e}")
        return empty
    for c in cols:
        if c not in t.columns:
            t[c] = ""
    t = t[cols].fillna("").astype(str)
    return t.apply(lambda s: s.str.strip())


def _write_table(t, path, prefix):
    tmp = os.path.join(BASE_DIR, "_tmp_" + os.path.basename(path))
    try:
        if os.path.isfile(path):
            os.makedirs(BACKUP_DIR, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            shutil.copy2(path, os.path.join(BACKUP_DIR, f"{prefix}_{stamp}.xlsx"))
            old = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith(prefix + "_"))
            for f in old[:-MAX_BACKUPS]:
                os.remove(os.path.join(BACKUP_DIR, f))
        t.to_excel(tmp, index=False)
        os.replace(tmp, path)
        return True
    except PermissionError:
        st.error(f"تعذّر الحفظ: ملف {os.path.basename(path)} مفتوح في Excel. أغلقه ثم أعد المحاولة.")
    except Exception as e:
        st.error(f"تعذّر الحفظ: {e}")
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
    return False


def audit(action, details):
    cur = _read_table(AUDIT_FILE, AUDIT_COLS)
    row = pd.DataFrame([{
        "الوقت": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "المستخدم": st.session_state.get("ovvo_username", ""),
        "الإجراء": action, "التفاصيل": details}], columns=AUDIT_COLS)
    _write_table(pd.concat([cur, row], ignore_index=True), AUDIT_FILE, "AUDIT")


def load_ledger():
    t = _read_table(LEDGER_FILE, LEDGER_COLS)
    t["المبلغ"] = pd.to_numeric(t["المبلغ"], errors="coerce").fillna(0.0)
    t["الحالة"] = t["الحالة"].replace("", LG_ACTIVE)
    t["المرجع"] = t["المرجع"].str.replace(r"^(\d+)\.0$", r"\1", regex=True)
    return t


def add_ledger_entry(kind, date_, cat, ref, amount, pay, memo):
    led = load_ledger()
    entry_id = f"L-{len(led) + 1:04d}"
    row = {"رقم القيد": entry_id, "التاريخ": date_.isoformat(), "النوع": kind, "الفئة": cat,
           "المرجع": str(ref), "المبلغ": round(float(amount), 2), "طريقة الدفع": pay,
           "البيان": memo.strip(),
           "المستخدم": st.session_state.get("ovvo_username", ""),
           "وقت الإدخال": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
           "الحالة": LG_ACTIVE, "سبب الإلغاء": ""}
    out = pd.concat([led, pd.DataFrame([row], columns=LEDGER_COLS)], ignore_index=True)
    if _write_table(out, LEDGER_FILE, "LEDGER"):
        audit("قيد جديد", f"{entry_id} | {kind} | {amount:,.2f} | {cat or ref} | {memo.strip()}")
        return entry_id
    return None


def void_ledger_entry(entry_id, reason):
    led = load_ledger()
    m = led["رقم القيد"] == entry_id
    if not m.any() or led.loc[m, "الحالة"].iloc[0] != LG_ACTIVE:
        st.error("القيد غير موجود أو ملغى مسبقًا.")
        return False
    led.loc[m, ["الحالة", "سبب الإلغاء"]] = [LG_VOID, reason.strip()]
    if _write_table(led, LEDGER_FILE, "LEDGER"):
        audit("إلغاء قيد", f"{entry_id} | السبب: {reason.strip()}")
        return True
    return False


# ---------- مبيعات / شحنات ----------
def sales_lines(orders, data):
    """أسطر الطلبات «تم التوصيل» مع الكلفة (لقطة وقت الشحن، وإلا الكلفة الحالية)."""
    o = orders[orders["الحالة"] == ST_DELIVERED].copy()
    cm = _cost_map(data)
    o["_unit_cost"] = o["الباركود"].map(cm).fillna(0.0)
    o["_cost_src"] = np.where(o["_unit_cost"] > 0, "الكلفة الحالية", "غير متوفرة")
    o["_qty"] = pd.to_numeric(o["الكمية"], errors="coerce").fillna(0)
    o["_amt"] = pd.to_numeric(o["المبلغ المطلوب"], errors="coerce").fillna(0)
    o["_cogs"] = o["_unit_cost"] * o["_qty"]
    st_date = o["تاريخ الحالة"].astype(str).str.strip()
    o["_date"] = np.where(st_date != "", st_date, o["التاريخ"].astype(str))
    o["_date"] = o["_date"].astype(str).str[:10]
    return o


def ship_summary(orders, data, ledger, vat):
    if orders.empty:
        return pd.DataFrame(columns=SHIP_COLS)
    sl = sales_lines(orders, data)
    fees = ledger[(ledger["الحالة"] == LG_ACTIVE) & (ledger["النوع"] == LG_FEE)]
    fee_map = fees.groupby("المرجع")["المبلغ"].sum().to_dict()
    rows = []
    for oid, g in orders.groupby("رقم الطلب"):
        s = sl[sl["رقم الطلب"] == oid]
        gross = float(s["_amt"].sum())
        net = gross / (1 + vat / 100)
        cogs = float(s["_cogs"].sum())
        fee = float(fee_map.get(str(int(oid)), 0.0))
        pend = g["الحالة"].isin([ST_OUT, ST_RET_WAIT, ST_CAN_WAIT]).any()
        rows.append({
            "رقم الشحنة": int(oid), "التاريخ": str(g["التاريخ"].iloc[0]),
            "المندوب": str(g["المندوب"].iloc[0]),
            "القطع المشحونة": int(pd.to_numeric(g["الكمية"], errors="coerce").fillna(0).sum()),
            "قطع مُسلّمة": int(s["_qty"].sum()),
            "المبيعات (شامل الضريبة)": round(gross, 2), "المبيعات الصافية": round(net, 2),
            "التكلفة": round(cogs, 2), "أجرة التوصيل": round(fee, 2),
            "ربح الشحنة": round(net - cogs - fee, 2),
            "الحالة": "مفتوحة" if pend else "مغلقة"})
    return pd.DataFrame(rows, columns=SHIP_COLS).sort_values("رقم الشحنة", ascending=False)


def driver_balances(orders, ledger):
    o = orders.copy()
    o["_d"] = o["المندوب"].replace("", NO_DRIVER)
    o["_a"] = pd.to_numeric(o["المبلغ المطلوب"], errors="coerce").fillna(0)
    deliv = o[o["الحالة"] == ST_DELIVERED].groupby("_d")["_a"].sum()
    out = o[o["الحالة"] == ST_OUT].groupby("_d")["_a"].sum()
    act = ledger[ledger["الحالة"] == LG_ACTIVE]
    rc = act[act["النوع"] == LG_RCPT].groupby("المرجع")["المبلغ"].sum()
    drv_of = o.groupby("رقم الطلب")["_d"].first()
    ded = {}
    for _, r in act[(act["النوع"] == LG_FEE) & (act["طريقة الدفع"] == PAY_DRV)].iterrows():
        try:
            d = drv_of.get(int(float(r["المرجع"])), NO_DRIVER)
        except ValueError:
            d = NO_DRIVER
        ded[d] = ded.get(d, 0.0) + float(r["المبلغ"])
    names = sorted(set(deliv.index) | set(out.index) | set(rc.index) | set(ded))
    t = pd.DataFrame({
        "المندوب": names,
        "مبيعات مُسلّمة": [float(deliv.get(n, 0.0)) for n in names],
        "مقبوض منه": [float(rc.get(n, 0.0)) for n in names],
        "أجور مخصومة": [float(ded.get(n, 0.0)) for n in names],
        "قيد التوصيل (لم يُسلَّم بعد)": [float(out.get(n, 0.0)) for n in names]})
    t["الرصيد المستحق عليه"] = t["مبيعات مُسلّمة"] - t["مقبوض منه"] - t["أجور مخصومة"]
    return t.round(2)


# ---------- دفتر اليومية ----------
def build_journal(data, orders, moves, ledger, vat):
    J = []

    def add(ref, date_, source, memo, lines, link=""):
        for acct, dr, cr in lines:
            dr, cr = round(float(dr), 2), round(float(cr), 2)
            if dr or cr:
                J.append({"رقم القيد": ref, "التاريخ": str(date_)[:10], "المصدر": source,
                          "الحساب": acct, "مدين": dr, "دائن": cr, "البيان": memo, "المرجع": link})

    cm = _cost_map(data)

    # 1) المبيعات (عند «تم التوصيل») + تكلفتها
    for i, r in sales_lines(orders, data).iterrows():
        amt = round(float(r["_amt"]), 2)
        net = round(amt / (1 + vat / 100), 2)
        tax = round(amt - net, 2)
        oid = int(float(r["رقم الطلب"]))
        memo = f"بيع {r['الوصف']} ({r['اللون']}/{r['المقاس']}) × {int(r['_qty'])}"
        link = f"شحنة {oid}"
        add(f"S-{oid}-{i}", r["_date"], "مبيعات", memo,
            [(A_DRV, amt, 0), (A_REV, 0, net), (A_VAT, 0, tax)], link)
        c = round(float(r["_cogs"]), 2)
        add(f"C-{oid}-{i}", r["_date"], "تكلفة المبيعات", memo, [(A_COGS, c, 0), (A_INV, 0, c)], link)

    # 2) حركات المخزون: إدخال (مشتريات) + تسويات يدوية
    for n, m in moves.reset_index(drop=True).iterrows():
        kind = m["نوع الحركة"]
        if kind not in (MV_IN, MV_ADJ, MV_DEL):
            continue
        q = float(m["الكمية"])
        unit = cm.get(m["الباركود"])
        if unit is None or q == 0:
            continue
        v = round(abs(q) * unit, 2)
        memo = f"{kind}: {m['الوصف']} ({m['اللون']}/{m['المقاس']}) {q:+g}"
        if kind == MV_IN:
            add(f"P-{n + 1}", m["التاريخ"], "إدخال مخزون", memo, [(A_INV, v, 0), (A_AP, 0, v)], m["الباركود"])
        elif q > 0:
            add(f"A-{n + 1}", m["التاريخ"], "تسوية مخزون", memo, [(A_INV, v, 0), (A_ADJ, 0, v)], m["الباركود"])
        else:
            add(f"A-{n + 1}", m["التاريخ"], "تسوية مخزون", memo, [(A_ADJ, v, 0), (A_INV, 0, v)], m["الباركود"])

    # 3) المصاريف / أجور التوصيل / القبض من المندوب / الدفع للموردين
    pay_acc = {PAY_CASH: A_CASH, PAY_BANK: A_BANK, PAY_DRV: A_DRV}
    for _, r in ledger[ledger["الحالة"] == LG_ACTIVE].iterrows():
        amt = round(float(r["المبلغ"]), 2)
        pa = pay_acc.get(r["طريقة الدفع"], A_CASH)
        k = r["النوع"]
        memo = r["البيان"] or r["الفئة"] or k
        link = r["المرجع"]
        if k == LG_EXP:
            lines, src = [(f"مصروف - {r['الفئة'] or 'أخرى'}", amt, 0), (pa, 0, amt)], "مصاريف"
        elif k == LG_FEE:
            lines, src, link = [(A_FEE, amt, 0), (pa, 0, amt)], "أجور توصيل", f"شحنة {link}"
        elif k == LG_RCPT:
            lines, src, link = [(pa, amt, 0), (A_DRV, 0, amt)], "قبض من مندوب", f"مندوب: {link}"
        elif k == LG_SUP:
            lines, src, link = [(A_AP, amt, 0), (pa, 0, amt)], "دفع لمورد", f"مورد: {link}"
        else:
            continue
        add(r["رقم القيد"], r["التاريخ"], src, memo, lines, link)

    j = pd.DataFrame(J, columns=JCOLS)
    return j.sort_values("التاريخ", kind="stable").reset_index(drop=True)


def acct_kind(a):
    if a == A_REV:
        return "إيراد"
    if a in (A_COGS, A_FEE, A_ADJ) or a.startswith("مصروف"):
        return "مصروف"
    if a in (A_VAT, A_AP):
        return "التزام"
    return "أصل"


def _period(t, col, d_from, d_to):
    if t.empty:
        return t
    ds = t[col].astype(str).str[:10]
    return t[(ds >= d_from.isoformat()) & (ds <= d_to.isoformat())]


def _bal(j, acct, credit=True):
    s = j[j["الحساب"] == acct]
    dr, cr = float(s["مدين"].sum()), float(s["دائن"].sum())
    return cr - dr if credit else dr - cr


# ---------- فحوصات الرقابة ----------
def _sl_view(sl):
    v = sl[["رقم الطلب", "الوصف", "اللون", "المقاس", "الباركود"]].copy()
    v["الكمية"], v["المبلغ"] = sl["_qty"], sl["_amt"]
    v["التكلفة"], v["مصدر الكلفة"] = sl["_cogs"], sl["_cost_src"]
    return v


def control_checks(data, orders, moves, ledger, ships, vat, days, d_from, d_to):
    checks = []
    cm = _cost_map(data)
    cost = pd.to_numeric(data["الكلفة النهائية"], errors="coerce").fillna(0)
    checks.append(("أصناف بالمخزون بدون كلفة", "عالية",
                   data[(cost <= 0) & (qty(data) > 0)][["الباركود التلقائي", "وصف", "اللون", "المقاس", "الكمية"]]))
    sl = sales_lines(orders, data)
    sl = sl[(sl["_date"] >= d_from.isoformat()) & (sl["_date"] <= d_to.isoformat())] if len(sl) else sl
    checks.append(("مبيعات مُسلّمة بدون كلفة معروفة", "عالية", _sl_view(sl[sl["_unit_cost"] <= 0])))
    checks.append(("بيع بأقل من الكلفة", "عالية",
                   _sl_view(sl[(sl["_unit_cost"] > 0) & (sl["_amt"] < sl["_cogs"])])))
    checks.append(("مبيعات بسعر صفر", "عالية", _sl_view(sl[sl["_amt"] <= 0])))
    sp = _period(ships, "التاريخ", d_from, d_to)
    checks.append(("شحنات فيها مبيعات بدون أجرة توصيل مسجّلة", "متوسطة",
                   sp[(sp["قطع مُسلّمة"] > 0) & (sp["أجرة التوصيل"] <= 0)]))
    o = orders[orders["الحالة"].isin([ST_OUT, ST_RET_WAIT, ST_CAN_WAIT])].copy()
    if len(o):
        age = (pd.Timestamp.now() - pd.to_datetime(o["التاريخ"], errors="coerce")).dt.days
        o["أيام مفتوحة"] = age
        o = o[o["أيام مفتوحة"] >= days][["رقم الطلب", "التاريخ", "المندوب", "الوصف", "الكمية",
                                         "المبلغ المطلوب", "الحالة", "أيام مفتوحة"]]
    else:
        o = pd.DataFrame(columns=["رقم الطلب", "التاريخ", "المندوب", "الوصف", "الكمية",
                                  "المبلغ المطلوب", "الحالة", "أيام مفتوحة"])
    checks.append((f"شحنات/مرتجعات مفتوحة منذ {days} يوم أو أكثر", "عالية", o))
    adj = _period(moves[moves["نوع الحركة"].isin([MV_ADJ, MV_DEL])], "التاريخ", d_from, d_to).copy()
    adj["القيمة بالكلفة"] = [round(q * cm.get(b, 0.0), 2) for q, b in zip(adj["الكمية"], adj["الباركود"])]
    checks.append(("تعديلات/حذف كميات يدوية على المخزون", "متوسطة",
                   adj[["التاريخ", "نوع الحركة", "الوصف", "اللون", "المقاس", "الكمية",
                        "القيمة بالكلفة", "المستخدم", "ملاحظة"]]))
    voided = _period(ledger[ledger["الحالة"] == LG_VOID], "التاريخ", d_from, d_to)
    checks.append(("قيود ملغاة", "متوسطة",
                   voided[["رقم القيد", "التاريخ", "النوع", "المبلغ", "البيان", "المستخدم", "سبب الإلغاء"]]))
    db = driver_balances(orders, ledger)
    checks.append(("مندوب مقبوض منه أكثر من المستحق (رصيد سالب)", "عالية",
                   db[db["الرصيد المستحق عليه"] < -0.005]))
    return checks


# =========================
# صفحة المحاسبة
# =========================
def page_accounting():
    require_accounting()
    page_header("المحاسبة والرقابة المالية",
                "قيود البيع والمخزون والمصاريف وأجور التوصيل، مع ذمم المندوبين وميزان المراجعة وسجل الرقابة.")
    orders, moves, ledger = load_orders(), load_moves(), load_ledger()
    today = date.today()
    c1, c2, c3 = st.columns(3)
    d_from = c1.date_input("من تاريخ", value=today.replace(day=1), key="acc_from")
    d_to = c2.date_input("إلى تاريخ", value=today, key="acc_to")
    vat = c3.number_input("نسبة ضريبة المبيعات % (الأسعار شاملة الضريبة، 0 = بدون)",
                          min_value=0.0, max_value=50.0, value=16.0, step=0.5, key="acc_vat")
    if d_from > d_to:
        st.error("تاريخ البداية بعد تاريخ النهاية.")
        return

    j_all = build_journal(df, orders, moves, ledger, vat)
    j = _period(j_all, "التاريخ", d_from, d_to)
    ships = ship_summary(orders, df, ledger, vat)
    ships_p = _period(ships, "التاريخ", d_from, d_to)

    tabs = st.tabs(["📊 الملخص", "🚚 أجور التوصيل", "🧾 المصاريف والقبوض", "👥 ذمم المندوبين",
                    "📒 دفتر اليومية", "⚖️ ميزان المراجعة والأستاذ", "🛡️ الرقابة والمراجعة"])

    # ---------------- الملخص ----------------
    with tabs[0]:
        rev, tax = _bal(j, A_REV), _bal(j, A_VAT)
        cogs, fee, adj = _bal(j, A_COGS, False), _bal(j, A_FEE, False), _bal(j, A_ADJ, False)
        om = j[j["الحساب"].str.startswith("مصروف - ")]
        other = float(om["مدين"].sum() - om["دائن"].sum())
        gross = rev - cogs
        net = gross - fee - adj - other
        a1, a2, a3, a4 = st.columns(4)
        a1.metric("صافي المبيعات (بدون ضريبة)", f"{rev:,.2f}")
        a2.metric("ضريبة المبيعات المستحقة", f"{tax:,.2f}")
        a3.metric("تكلفة البضاعة المباعة", f"{cogs:,.2f}")
        a4.metric("مجمل الربح", f"{gross:,.2f}", f"{(gross / rev * 100 if rev else 0):.1f}%")
        b1, b2, b3, b4 = st.columns(4)
        b1.metric("أجور التوصيل", f"{fee:,.2f}")
        b2.metric("مصاريف أخرى", f"{other:,.2f}")
        b3.metric("فروقات/تسويات مخزون", f"{adj:,.2f}")
        b4.metric("صافي الربح", f"{net:,.2f}")

        cm = _cost_map(df)
        stock_cost = float(sum(q * cm.get(b, 0.0) for q, b in zip(qty(df), df["الباركود التلقائي"])))
        stock_retail = float((qty(df) * pd.to_numeric(df["سعر البيع"], errors="coerce").fillna(0)).sum())
        ot = orders[orders["الحالة"] == ST_OUT]
        transit = float(sum(float(q) * cm.get(b, 0.0)
                            for q, b in zip(pd.to_numeric(ot["الكمية"], errors="coerce").fillna(0), ot["الباركود"])))
        db = driver_balances(orders, ledger)
        d1, d2, d3, d4 = st.columns(4)
        d1.metric("قيمة المخزون بالكلفة", f"{stock_cost:,.2f}")
        d2.metric("قيمة المخزون بسعر البيع", f"{stock_retail:,.2f}")
        d3.metric("بضاعة عند المندوبين (كلفة)", f"{transit:,.2f}")
        d4.metric("كاش مستحق على المندوبين", f"{db['الرصيد المستحق عليه'].sum():,.2f}")

        st.markdown("#### الأرباح حسب الشحنة (شحنات الفترة)")
        if ships_p.empty:
            st.info("لا توجد شحنات بهالفترة.")
        else:
            st.dataframe(ships_p, use_container_width=True, hide_index=True)
            download_buttons(ships_p, "ربح_الشحنات", "acc_ships")

        st.markdown("#### ربح كل طلب مُسلّم (قبل أجرة التوصيل)")
        sl = sales_lines(orders, df)
        if len(sl):
            sl = sl[(sl["_date"] >= d_from.isoformat()) & (sl["_date"] <= d_to.isoformat())]
        if len(sl):
            v = _sl_view(sl)
            v["صافي البيع"] = (sl["_amt"] / (1 + vat / 100)).round(2)
            v["الربح"] = (v["صافي البيع"] - sl["_cogs"]).round(2)
            st.dataframe(v, use_container_width=True, hide_index=True)
            download_buttons(v, "ربح_الطلبات", "acc_lines")
        else:
            st.info("لا توجد طلبات «تم التوصيل» بهالفترة.")

    # ---------------- أجور التوصيل ----------------
    with tabs[1]:
        st.markdown("#### كلفة التوصيل لكل شحنة")
        sp = ships_p
        tot_fee = float(sp["أجرة التوصيل"].sum()) if len(sp) else 0.0
        with_fee = int((sp["أجرة التوصيل"] > 0).sum()) if len(sp) else 0
        tot_net = float(sp["المبيعات الصافية"].sum()) if len(sp) else 0.0
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("إجمالي أجور التوصيل", f"{tot_fee:,.2f}")
        k2.metric("متوسط الأجرة للشحنة", f"{(tot_fee / with_fee if with_fee else 0):,.2f}")
        k3.metric("نسبة الأجور من المبيعات", f"{(tot_fee / tot_net * 100 if tot_net else 0):.1f}%")
        k4.metric("شحنات بدون أجرة مسجلة", int(((sp["أجرة التوصيل"] <= 0) & (sp["قطع مُسلّمة"] > 0)).sum()) if len(sp) else 0)
        if len(sp):
            st.dataframe(sp[["رقم الشحنة", "التاريخ", "المندوب", "قطع مُسلّمة", "المبيعات الصافية",
                             "التكلفة", "أجرة التوصيل", "ربح الشحنة", "الحالة"]],
                         use_container_width=True, hide_index=True)

        st.markdown("#### تسجيل أجرة توصيل لشحنة")
        labels = {f"#{int(r['رقم الشحنة'])} — {str(r['التاريخ'])[:10]} — {r['المندوب'] or '—'} — "
                  f"{r['القطع المشحونة']} قطعة": int(r["رقم الشحنة"]) for _, r in ships.iterrows()}
        if not labels:
            st.info("لا توجد شحنات بعد.")
        else:
            with st.form("fee_form", clear_on_submit=True):
                sel = st.selectbox("الشحنة", list(labels))
                amount = st.number_input("أجرة التوصيل", min_value=0.0, step=0.25, format="%.2f")
                pay = st.selectbox("طريقة الدفع", [PAY_DRV, PAY_CASH, PAY_BANK],
                                   help="«خصم من كاش المندوب» = المندوب أخذ أجرته من الكاش اللي معه.")
                memo = st.text_input("ملاحظة")
                ok = st.form_submit_button("تسجيل الأجرة", type="primary")
            if ok:
                oid = labels[sel]
                exist = ledger[(ledger["النوع"] == LG_FEE) & (ledger["المرجع"] == str(oid))
                               & (ledger["الحالة"] == LG_ACTIVE)]
                if amount <= 0:
                    st.error("اكتب مبلغ أكبر من صفر.")
                elif len(exist):
                    st.error(f"الشحنة #{oid} إلها أجرة مسجلة بالقيد {exist['رقم القيد'].iloc[0]}. "
                             "ألغِ القيد أولًا (مع ذكر السبب) من تبويب «المصاريف والقبوض».")
                elif add_ledger_entry(LG_FEE, today, "أجور توصيل", oid, amount, pay,
                                      memo or f"أجرة توصيل شحنة {oid}"):
                    flash(f"تم تسجيل أجرة الشحنة #{oid}.")
                    st.rerun()

    # ---------------- المصاريف والقبوض ----------------
    with tabs[2]:
        kind = st.radio("نوع القيد", [LG_EXP, LG_RCPT, LG_SUP], horizontal=True, key="lg_kind")
        drivers = sorted(set(orders["المندوب"].replace("", NO_DRIVER))) or [NO_DRIVER]
        with st.form("ledger_form", clear_on_submit=True):
            f1, f2, f3 = st.columns(3)
            d_ = f1.date_input("التاريخ", value=today)
            amount = f2.number_input("المبلغ", min_value=0.0, step=0.5, format="%.2f")
            pay = f3.selectbox("الدفع / القبض عبر", [PAY_CASH, PAY_BANK])
            cat = ref = ""
            if kind == LG_EXP:
                cat = st.selectbox("الفئة", EXP_CATS)
            elif kind == LG_RCPT:
                ref = st.selectbox("المندوب", drivers)
            else:
                ref = st.text_input("اسم المورد *")
            memo = st.text_input("البيان * (إجباري للرقابة)")
            ok = st.form_submit_button("تسجيل القيد", type="primary")
        if ok:
            if amount <= 0:
                st.error("اكتب مبلغ أكبر من صفر.")
            elif not memo.strip():
                st.error("البيان إجباري.")
            elif d_ > today:
                st.error("ما بيصير تسجل قيد بتاريخ مستقبلي.")
            elif kind == LG_SUP and not ref.strip():
                st.error("اكتب اسم المورد.")
            else:
                eid = add_ledger_entry(kind, d_, cat, ref.strip(), amount, pay, memo)
                if eid:
                    flash(f"تم تسجيل القيد {eid}.")
                    st.rerun()

        st.markdown("#### سجل القيود اليدوية")
        lp = _period(ledger, "التاريخ", d_from, d_to)
        if lp.empty:
            st.info("لا توجد قيود بهالفترة.")
        else:
            st.dataframe(lp.iloc[::-1], use_container_width=True, hide_index=True)
            download_buttons(lp, "القيود_اليدوية", "acc_ledger")
        active = ledger[ledger["الحالة"] == LG_ACTIVE]
        if len(active):
            with st.expander("إلغاء قيد (ما في حذف، الإلغاء بيبقى ظاهر بالسجل)"):
                with st.form("void_form", clear_on_submit=True):
                    opts = {f"{r['رقم القيد']} | {r['النوع']} | {r['المبلغ']:,.2f} | {r['البيان']}": r["رقم القيد"]
                            for _, r in active.iloc[::-1].iterrows()}
                    pick = st.selectbox("القيد", list(opts))
                    reason = st.text_input("سبب الإلغاء *")
                    vok = st.form_submit_button("إلغاء القيد")
                if vok:
                    if len(reason.strip()) < 3:
                        st.error("اكتب سبب الإلغاء.")
                    elif void_ledger_entry(opts[pick], reason):
                        flash("تم إلغاء القيد.")
                        st.rerun()

    # ---------------- ذمم المندوبين ----------------
    with tabs[3]:
        st.caption("المستحق = مبيعات «تم التوصيل» − ما قبضناه منه − أجوره المخصومة من الكاش. "
                   "سجّل كل تسليم كاش من تبويب «المصاريف والقبوض» (قبض من مندوب).")
        db = driver_balances(orders, ledger)
        if db.empty:
            st.info("لا توجد بيانات توصيل.")
        else:
            st.metric("إجمالي الكاش المستحق على المندوبين", f"{db['الرصيد المستحق عليه'].sum():,.2f}")
            st.dataframe(db, use_container_width=True, hide_index=True)
            download_buttons(db, "ذمم_المندوبين", "acc_drivers")

    # ---------------- دفتر اليومية ----------------
    with tabs[4]:
        if j.empty:
            st.info("لا توجد قيود بهالفترة.")
        else:
            srcs = ["الكل"] + sorted(j["المصدر"].unique())
            f_src = st.selectbox("المصدر", srcs, key="acc_src")
            jj = j if f_src == "الكل" else j[j["المصدر"] == f_src]
            m1, m2, m3 = st.columns(3)
            m1.metric("مجموع المدين", f"{jj['مدين'].sum():,.2f}")
            m2.metric("مجموع الدائن", f"{jj['دائن'].sum():,.2f}")
            m3.metric("الفرق (لازم صفر)", f"{round(jj['مدين'].sum() - jj['دائن'].sum(), 2):,.2f}")
            st.dataframe(jj, use_container_width=True, hide_index=True)
            download_buttons(jj, "دفتر_اليومية", "acc_journal")

    # ---------------- ميزان المراجعة والأستاذ ----------------
    with tabs[5]:
        if j.empty:
            st.info("لا توجد قيود بهالفترة.")
        else:
            tb = j.groupby("الحساب")[["مدين", "دائن"]].sum().reset_index()
            tb.insert(1, "النوع", tb["الحساب"].map(acct_kind))
            diff = tb["مدين"] - tb["دائن"]
            tb["رصيد مدين"] = diff.clip(lower=0).round(2)
            tb["رصيد دائن"] = (-diff).clip(lower=0).round(2)
            tot = {"الحساب": "الإجمالي", "النوع": "", "مدين": tb["مدين"].sum(), "دائن": tb["دائن"].sum(),
                   "رصيد مدين": tb["رصيد مدين"].sum(), "رصيد دائن": tb["رصيد دائن"].sum()}
            st.markdown("#### ميزان المراجعة")
            st.dataframe(pd.concat([tb, pd.DataFrame([tot])], ignore_index=True),
                         use_container_width=True, hide_index=True)
            st.caption("رصيد حساب «المخزون» هنا = حركة الفترة فقط (الرصيد الافتتاحي قبل تشغيل السجل غير مُدخل). "
                       "للقيمة الفعلية شوف «قيمة المخزون بالكلفة» بالملخص.")
            st.markdown("#### كشف حساب (أستاذ)")
            acct = st.selectbox("الحساب", sorted(j["الحساب"].unique()), key="acc_ledger_sel")
            g = j[j["الحساب"] == acct].copy()
            g["الرصيد"] = (g["مدين"] - g["دائن"]).cumsum().round(2)
            st.dataframe(g[["التاريخ", "رقم القيد", "المصدر", "البيان", "مدين", "دائن", "الرصيد"]],
                         use_container_width=True, hide_index=True)
            download_buttons(g, "كشف_حساب", "acc_acct")

    # ---------------- الرقابة والمراجعة ----------------
    with tabs[6]:
        days = int(st.number_input("اعتبر الشحنة/المرتجع متأخرًا بعد (أيام)", min_value=1, max_value=90,
                                   value=7, step=1, key="acc_days"))
        checks = control_checks(df, orders, moves, ledger, ships, vat, days, d_from, d_to)
        st.dataframe(pd.DataFrame([{
            "الفحص": n, "الأهمية": s, "العدد": len(t),
            "الحالة": "✅ سليم" if t.empty else "⚠️ يحتاج مراجعة"} for n, s, t in checks]),
            use_container_width=True, hide_index=True)
        for n, s, t in checks:
            if not t.empty:
                with st.expander(f"{n} ({len(t)})"):
                    st.dataframe(t, use_container_width=True, hide_index=True)
        st.markdown("#### سجل الرقابة (من عمل إيش ومتى)")
        au = _read_table(AUDIT_FILE, AUDIT_COLS)
        if au.empty:
            st.info("لا توجد عمليات مسجّلة بعد.")
        else:
            st.dataframe(au.iloc[::-1], use_container_width=True, hide_index=True)
            download_buttons(au, "سجل_الرقابة", "acc_audit")


# =========================
# تكاليف المنتجات (المحاسب فقط)
# =========================
def page_costs():
    require_accounting()
    page_header("تكاليف المنتجات",
                "أدخل كلفة كل موديل مرة وحدة، وبتنطبق على كل ألوانه ومقاساته. الأرباح بتنحسب لحالها لما الطلب يصير «تم التوصيل».")
    data, _ = load_data()
    if data.empty:
        st.info("لا توجد أصناف بعد.")
        return
    keys = pd.Series([m if m else f"بدون موديل #{b}" for m, b in
                      zip(data["رقم الموديل"], data["الباركود التلقائي"])], index=data.index)
    cost = pd.to_numeric(data["الكلفة النهائية"], errors="coerce").fillna(0.0)
    price = pd.to_numeric(data["سعر البيع"], errors="coerce").fillna(0.0)
    rows = []
    for k in sorted(keys.unique(), key=_model_sort_key):
        m = keys == k
        c = cost[m]
        c_pos = c[c > 0]
        rows.append({
            "الموديل": k,
            "الوصف": next((x for x in data.loc[m, "وصف"] if x), ""),
            "نوع الصنف": next((x for x in data.loc[m, "نوع الصنف"] if x), ""),
            "القطع": int(qty(data)[m].sum()),
            "سعر البيع": float(price[m].max()),
            "الكلفة": float(c_pos.iloc[0]) if len(c_pos) else 0.0,
        })
    full = pd.DataFrame(rows)

    c1, c2 = st.columns([2, 1])
    q = c1.text_input("بحث (موديل أو وصف)", key="cost_q")
    only_missing = c2.checkbox("الموديلات بدون كلفة فقط", key="cost_missing")
    view = full
    if q.strip():
        t = q.strip().lower()
        view = view[view["الموديل"].astype(str).str.lower().str.contains(t, regex=False)
                    | view["الوصف"].astype(str).str.lower().str.contains(t, regex=False)]
    if only_missing:
        view = view[view["الكلفة"] <= 0]
    st.caption(f"{len(view):,} موديل — بدون كلفة: {int((full['الكلفة'] <= 0).sum()):,}")
    if view.empty:
        st.info("لا نتائج.")
        return

    edited = st.data_editor(
        view, hide_index=True, use_container_width=True, num_rows="fixed",
        disabled=["الموديل", "الوصف", "نوع الصنف", "القطع", "سعر البيع"],
        column_config={
            "الكلفة": st.column_config.NumberColumn("الكلفة (للقطعة)", min_value=0.0, step=0.25, format="%.2f"),
            "سعر البيع": st.column_config.NumberColumn("سعر البيع", format="%.2f"),
        },
        key=f"cost_editor_{hashlib.md5((q + str(only_missing)).encode('utf-8')).hexdigest()[:8]}")

    if st.button("💾 حفظ التكاليف", type="primary", use_container_width=True):
        old = dict(zip(full["الموديل"], full["الكلفة"]))
        changes = []
        for _, r in edited.iterrows():
            new = pd.to_numeric(r["الكلفة"], errors="coerce")
            if pd.isna(new) or new < 0:
                st.error(f"كلفة غير صالحة للموديل {r['الموديل']}.")
                return
            if abs(float(new) - old[r["الموديل"]]) > 1e-9:
                changes.append((r["الموديل"], old[r["الموديل"]], round(float(new), 4)))
        if not changes:
            st.info("ما في تعديلات للحفظ.")
            return
        for k, _, new in changes:
            data.loc[keys == k, "الكلفة النهائية"] = new
        if save_data(data):
            audit("تغيير كلفة", " | ".join(f"موديل {k}: {o:,.2f} ← {n:,.2f}" for k, o, n in changes))
            flash(f"تم حفظ كلفة {len(changes)} موديل.")
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
elif page == "📋 كشف المواد":
    page_statement()
elif page == "📑 التقارير والطباعة":
    page_reports()
elif page == "👤 إدارة المستخدمين":
    page_users()
elif page == "💰 كلف المنتجات":
    page_costs()
elif page == "🧮 المحاسبة":
    page_accounting()
