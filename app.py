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
st.set_page_config(page_title="OVVO Store", page_icon="📦", layout="wide")

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
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800&display=swap');

:root {
    --dark: #242421;
    --beige: #D6C1A3;
    --cream: #F7F2E9;
    --border: #E3D7C5;
    --card: #FFFCF7;
    --ok: #4F7A5A;
    --warn: #B7791F;
    --bad: #A8403A;
}

.stApp, .stApp :is(p, label, li, h1, h2, h3, h4, input, textarea, button, a,
    [data-testid="stMetricValue"], [data-testid="stMetricLabel"]) {
    font-family: 'Cairo', sans-serif;
}
.stApp { background: var(--cream); color: #292722; }
[data-testid="stAppViewContainer"], [data-testid="stSidebar"] {
    direction: rtl; text-align: right;
}
[data-testid="stHeader"] { background: transparent; }
#MainMenu, footer { visibility: hidden; }
.block-container { padding-top: 2rem; max-width: 1250px; }

h1, h2, h3, h4 { color: var(--dark) !important; font-weight: 800 !important; }

[data-testid="stSidebar"] { background: var(--dark); }
[data-testid="stSidebar"] * { color: var(--cream); }
[data-testid="stSidebar"] [role="radiogroup"] { gap: 4px; }
[data-testid="stSidebar"] [role="radiogroup"] label {
    padding: 10px 12px; border-radius: 10px; width: 100%;
}
[data-testid="stSidebar"] [role="radiogroup"] label:hover {
    background: rgba(214,193,163,.14);
}
[data-testid="stSidebar"] .stButton > button {
    background: transparent; border: 1px solid var(--beige); color: var(--cream);
}
[data-testid="stSidebar"] .stButton > button:hover {
    background: var(--beige); color: var(--dark);
}

.ovvo-brand {
    background: rgba(255,255,255,.04); border: 1px solid rgba(214,193,163,.35);
    border-radius: 14px; padding: 16px; margin-bottom: 14px; text-align: center;
}
.ovvo-brand-name { font-size: 22px; font-weight: 800; letter-spacing: 2px; color: var(--cream); }
.ovvo-brand-sub { font-size: 11px; letter-spacing: 1px; color: var(--beige); }
.user-chip {
    background: rgba(255,255,255,.06); border-radius: 10px;
    padding: 8px 12px; margin-bottom: 10px; font-size: 14px;
}

.page-head { margin-bottom: 18px; }
.page-head h1 { margin: 0; font-size: 28px; }
.page-head p { margin: 4px 0 0; color: #6f6759; font-size: 15px; }
.hero {
    background: var(--dark); border-radius: 16px; padding: 28px 26px; height: 100%;
}
.hero-title { color: var(--cream); font-size: 26px; font-weight: 800; line-height: 1.5; }
.hero-sub { color: var(--beige); font-size: 15px; margin-top: 6px; }

[data-testid="stMetric"] {
    background: var(--card); border: 1px solid var(--border);
    border-right: 4px solid var(--beige); padding: 14px 16px; border-radius: 12px;
}
[data-testid="stMetricLabel"] { color: #6f6759; }
[data-testid="stMetricValue"] { color: var(--dark); font-weight: 800; }

.stButton > button, .stDownloadButton > button,
[data-testid="stFormSubmitButton"] > button {
    border-radius: 10px; font-weight: 700; border: 1px solid var(--dark);
}
button[kind="primary"], [data-testid="stBaseButton-primary"],
[data-testid="stBaseButton-primaryFormSubmit"] {
    background: var(--dark) !important; color: var(--cream) !important;
    border-color: var(--dark) !important;
}
button[kind="primary"]:hover, [data-testid="stBaseButton-primary"]:hover,
[data-testid="stBaseButton-primaryFormSubmit"]:hover { background: #3b3a35 !important; }
[data-baseweb="input"], [data-baseweb="select"] > div, textarea {
    border-radius: 10px !important; background: var(--card) !important;
}
input, textarea { text-align: right; }
[data-testid="stForm"] {
    background: var(--card); border: 1px solid var(--border); border-radius: 14px;
}
.stTabs [data-baseweb="tab"] { font-weight: 700; }

@media print {
    [data-testid="stSidebar"], [data-testid="stHeader"], .stButton,
    .stDownloadButton, iframe { display: none !important; }
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


def brand_block():
    st.markdown(
        """
<div class="ovvo-brand" style="background:#242421;">
    <div class="ovvo-brand-name">OVVO STORE</div>
    <div class="ovvo-brand-sub">INVENTORY MANAGEMENT</div>
</div>""",
        unsafe_allow_html=True,
    )


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
        raw = pd.read_excel(path, sheet_name=SHEET_NAME,
                            dtype={"الباركود التلقائي": str})
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
              "FREE", "FREE SIZE"]


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


def filter_ui(data, key, threshold):
    c1, c2, c3, c4, c5 = st.columns([2.2, 1, 1, 1, 1])
    query = c1.text_input("بحث", placeholder="الوصف، الباركود، الماركة، رقم الموديل…",
                          key=f"{key}_q")
    types = ["الكل"] + sorted(t for t in data["نوع الصنف"].unique() if t)
    brands = ["الكل"] + sorted(b for b in data["الماركة"].unique() if b)
    seasons = ["الكل"] + sorted(s for s in data["الموسم"].unique() if s)
    f_type = c2.selectbox("النوع", types, key=f"{key}_type")
    f_brand = c3.selectbox("الماركة", brands, key=f"{key}_brand")
    f_season = c4.selectbox("الموسم", seasons, key=f"{key}_season")
    f_status = c5.selectbox("حالة المخزون", STATUS_OPTIONS, key=f"{key}_status")

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
    if f_status != "الكل":
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
    page = st.radio("القائمة", pages, label_visibility="collapsed")
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 تسجيل الخروج", use_container_width=True):
        st.session_state.clear()
        st.rerun()

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
    st.markdown(
        '<div class="hero">'
        '<div class="hero-title">نظام إدارة المخزون والمستودعات</div>'
        f'<div class="hero-sub">أهلًا {current_username}</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    c_img_col1, c_img_col2 = st.columns([1, 1])
    with c_img_col2:
        if image_base64:
            st.markdown(
                f'<img src="data:image/png;base64,{image_base64}" '
                'style="width:100%;max-width:650px;height:auto;display:block;'
                'margin:auto;border-radius:14px;object-fit:contain;">',
                unsafe_allow_html=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)

    if df.empty:
        st.info("لا توجد أصناف بعد. " + ("أضف أول صنف من صفحة «إضافة صنف»." if current_role == "admin" else ""))
        return

    q = qty(df)
    out_count = int((q <= 0).sum())
    low_count = int(((q > 0) & (q <= threshold)).sum())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📊 إجمالي القطع", f"{q.sum():,.0f}")
    c2.metric("🏷️ عدد الأصناف", f"{len(df):,}")
    c3.metric("⚠️ مخزون منخفض", f"{low_count}")
    c4.metric("⛔ نفد من المخزون", f"{out_count}")

    st.markdown("### أصناف تحتاج إعادة توريد")
    with st.expander("إعدادات التنبيه"):
        st.number_input(
            "يُعتبر الصنف منخفضًا عندما تكون كميته أقل من أو تساوي",
            min_value=1, max_value=1000, value=5, step=1, key="low_threshold",
        )
    status = stock_status(df, threshold)
    need = df[status != "متوفر"].copy()
    if need.empty:
        st.success("كل الأصناف ضمن الحد الآمن.")
    else:
        need["الحالة"] = status[need.index]
        need = need.sort_values("الكمية")
        st.dataframe(need[display_columns(current_role)], use_container_width=True,
                     hide_index=True)

    st.markdown("### توزيع المخزون حسب النوع")
    by_type = summary_by(df, "نوع الصنف")
    t1, t2 = st.columns([1, 1.4])
    t1.dataframe(by_type, use_container_width=True, hide_index=True)
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

    filtered = filter_ui(df, "inv", threshold)
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
    types = sorted(t for t in df["نوع الصنف"].unique() if t)
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
        model = c3.text_input("رقم الموديل")
        barcode_in = c3.text_input("الباركود (اتركه فارغًا للتوليد التلقائي)")

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
        color = c1.text_input("اللون")
        size = c2.text_input("المقاس")

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
            season = e1.text_input("الموسم")
            year = e2.text_input("سنة الشراء")
            order = e3.text_input("الطلبية")
            invoice = e1.text_input("فاتورة رقم")
            origin = e2.text_input("المنشأة")
            tax = e3.text_input("ضريبة")
            hijab = e1.text_input("محجب أو غير محجب")
            sleeve = e2.text_input("كت أو كم أو نص كم")
            item_code = e3.text_input("كود الصنف")

        submitted = st.form_submit_button("حفظ الصنف", type="primary", use_container_width=True)

    if not submitted:
        return

    item_type = type_new.strip() or (type_sel if type_sel != none_opt else "")
    brand = brand_new.strip() or (brand_sel if brand_sel != none_opt else "")
    model, color, size = model.strip(), color.strip(), size.strip()
    barcode = barcode_in.strip() or next_barcode(df)
    photo = img_file or cam_file

    if not desc.strip():
        st.error("اكتب وصف الصنف.")
    elif not item_type:
        st.error("اختر نوع الصنف أو اكتب نوعًا جديدًا.")
    elif photo and not model:
        st.error("لازم تكتب رقم الموديل عشان تنربط الصورة فيه.")
    elif (df["الباركود التلقائي"] == barcode).any():
        st.error(f"الباركود {barcode} مستخدم لصنف آخر.")
    elif model and ((df["رقم الموديل"] == model) & (df["اللون"] == color) & (df["المقاس"] == size)).any():
        st.error("هذا الصنف موجود بالفعل. عدّل كميته من صفحة «المخزون».")
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
            "رقم الموديل": model, "كود الصنف": item_code.strip(),
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
    st.dataframe(shown, use_container_width=True, hide_index=True)
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
            st.dataframe(shown, use_container_width=True, hide_index=True)
            download_buttons(shown, "اصناف_تحتاج_توريد", "need")

    with tab3:
        shown = full[display_columns(current_role)]
        st.dataframe(shown, use_container_width=True, hide_index=True)
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
