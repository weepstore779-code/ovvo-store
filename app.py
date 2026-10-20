# -*- coding: utf-8 -*-
"""OVVO Store — نظام إدارة المخزون والمستودعات"""
import base64
import hashlib
import hmac
import json
import os
import secrets
import shutil
from datetime import datetime
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl.utils import get_column_letter

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
    pages = ["🏠 الرئيسية", "🔎 البحث", "📑 التقارير والطباعة"]
    if current_role == "admin":
        pages = ["🏠 الرئيسية", "📦 المخزون", "➕ إضافة صنف", "🚚 التوصيل",
                 "🔎 البحث", "📑 التقارير والطباعة", "👤 إدارة المستخدمين"]
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
def page_inventory():
    require_admin()
    page_header("إدارة المخزون", "عدّل الخلايا مباشرة ثم اضغط «حفظ التعديلات». تؤخذ نسخة احتياطية تلقائيًا قبل كل حفظ.")
    if df.empty:
        st.info("لا توجد أصناف بعد. أضف أول صنف من صفحة «إضافة صنف».")
        return

    filtered = filter_ui(df, "inv", threshold)
    st.caption(f"يعرض {len(filtered):,} من {len(df):,} صنف")

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

    if not desc.strip():
        st.error("اكتب وصف الصنف.")
    elif not item_type:
        st.error("اختر نوع الصنف أو اكتب نوعًا جديدًا.")
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
            st.session_state["add_form_id"] = form_id + 1
            flash(f"تمت إضافة «{desc.strip()}» برقم الباركود {barcode}.")
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
elif page == "🚚 التوصيل":
    page_delivery()
elif page == "🔎 البحث":
    page_search()
elif page == "📑 التقارير والطباعة":
    page_reports()
elif page == "👤 إدارة المستخدمين":
    page_users()