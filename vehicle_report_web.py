# -*- coding: utf-8 -*-
"""
Vehicle Report Importer — Web app (Streamlit) v1.1.0
ใช้ตัวอ่าน/กระทบยอด/Export ชุดเดียวกับ vehicle_report_app.py (ต้องอยู่โฟลเดอร์เดียวกัน)
รัน:  py -m streamlit run vehicle_report_web.py
"""
import os, sys, subprocess, tempfile, shutil, zipfile, time
from datetime import date, datetime

# เวลาไทยเมื่อรันบน Cloud (เซิร์ฟเวอร์เป็น UTC)
os.environ.setdefault("TZ", "Asia/Bangkok")
if hasattr(time, "tzset"):
    time.tzset()

import pandas as pd
import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import vehicle_report_app as core  # noqa: E402

WEB_VERSION = "1.3.0"
import inspect as _inspect
# ใช้ได้ทั้ง Streamlit รุ่นใหม่ (width="stretch") และรุ่นเก่า (**FULL)
FULL = ({"width": "stretch"} if "width" in _inspect.signature(st.button).parameters
        else {"use_container_width": True})
st.set_page_config(page_title="Vehicle Report Importer", page_icon="🚛", layout="wide",
                   initial_sidebar_state="expanded")

# รันบนเครื่อง Windows ของผู้ใช้ = โหมดเครื่อง (เลือกโฟลเดอร์/บันทึกลง D:) ; อื่น ๆ (Streamlit Cloud) = โหมดอัปโหลด
IS_LOCAL = os.name == "nt" or os.environ.get("VR_LOCAL") == "1"


def _secret(name, default=None):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


# รหัสผ่านชั้นที่ 2 (ถ้าตั้ง app_password ใน Secrets ของ Streamlit Cloud)
_pw = _secret("app_password")
if _pw and not st.session_state.get("auth_ok"):
    st.markdown("### 🔒 Vehicle Report Importer")
    v = st.text_input("รหัสผ่าน", type="password")
    if v:
        if v == _pw:
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("รหัสผ่านไม่ถูกต้อง")
    st.stop()

# ----------------------------------------------------------------------------- style
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');
html, body, [class*="css"], .stMarkdown, .stDataFrame, button, input { font-family: 'Sarabun', 'Tahoma', sans-serif; }
.block-container { padding-top: 1.2rem; padding-bottom: 2rem; max-width: 100%; }
.app-hero { background: linear-gradient(120deg, #0f3d68 0%, #1f6fb2 60%, #2aa7c9 100%);
  color: #fff; border-radius: 16px; padding: 18px 26px; margin-bottom: 14px;
  box-shadow: 0 6px 20px rgba(15,61,104,.18); }
.app-hero h1 { font-size: 1.55rem; margin: 0; font-weight: 700; color: #fff; }
.app-hero p { margin: 4px 0 0 0; opacity: .9; font-size: .95rem; }
.kpi { background: #fff; border: 1px solid #e6ebf1; border-radius: 14px; padding: 14px 16px;
  box-shadow: 0 2px 8px rgba(16,40,70,.05); height: 100%; }
.kpi .lbl { color: #5b6b7d; font-size: .82rem; font-weight: 600; text-transform: none; }
.kpi .val { color: #0f3d68; font-size: clamp(1rem, 1.35vw, 1.4rem); white-space: nowrap; font-weight: 700; line-height: 1.25; margin-top: 2px; }
.kpi .sub { color: #8795a6; font-size: .78rem; margin-top: 2px; }
.kpi.warn { border-left: 5px solid #e8833a; } .kpi.ok { border-left: 5px solid #2e9e6a; }
.kpi.blue { border-left: 5px solid #1f6fb2; }
.badge { display:inline-block; padding: 2px 10px; border-radius: 999px; font-size: .78rem; font-weight: 600; margin-right: 6px; }
.b-ok { background:#e3f5ec; color:#1d7a50; } .b-no { background:#fde8dc; color:#b4531a; } .b-info { background:#e5effa; color:#1f5f9c; }
.empty { border: 2px dashed #c9d6e3; border-radius: 16px; padding: 46px; text-align: center; color: #5b6b7d; background: #f8fbfe; }
section[data-testid="stSidebar"] { background: #f4f7fb; }
div[data-testid="stTabs"] button p { font-weight: 600; }
</style>
""", unsafe_allow_html=True)


# ----------------------------------------------------------------------------- helpers
def pick_folder(initial):
    """เปิดหน้าต่างเลือกโฟลเดอร์ของ Windows (ทำงานเมื่อรันบนเครื่องตัวเอง)"""
    code = ("import tkinter as tk; from tkinter import filedialog; r=tk.Tk(); r.withdraw(); "
            "r.wm_attributes('-topmost',1); "
            f"p=filedialog.askdirectory(initialdir={initial!r}, title='เลือกโฟลเดอร์'); print(p or '', end='')")
    try:
        res = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, timeout=600,
                             encoding="utf-8")
        p = res.stdout.strip()
        return os.path.normpath(p) if p else None
    except Exception as e:
        st.sidebar.error(f"เปิดหน้าต่างเลือกโฟลเดอร์ไม่ได้: {e} — พิมพ์ path เองได้")
        return None


ALLOWED_EXT = (".xls", ".xlsx")


def scan_uploads(files, dedupe, progress):
    """บันทึกไฟล์อัปโหลดลงโฟลเดอร์ชั่วคราว -> อ่าน -> ลบโฟลเดอร์ทิ้งทันที (ไม่เก็บข้อมูลบนเซิร์ฟเวอร์)"""
    tmp = tempfile.mkdtemp(prefix="vr_")
    try:
        n = 0
        for f in files:
            name = os.path.basename(f.name)
            if name.lower().endswith(".zip"):
                with zipfile.ZipFile(f) as z:
                    for info in z.infolist():
                        base = os.path.basename(info.filename)
                        if info.is_dir() or not base.lower().endswith(ALLOWED_EXT) or base.startswith(("~$", ".")):
                            continue
                        with z.open(info) as src, open(os.path.join(tmp, base), "wb") as dst:
                            shutil.copyfileobj(src, dst)
                        n += 1
            elif name.lower().endswith(ALLOWED_EXT):
                with open(os.path.join(tmp, name), "wb") as dst:
                    dst.write(f.getbuffer())
                n += 1
        ds = core.scan_folder(tmp, progress=progress, dedupe=dedupe)
        ds["summary"]["folder"] = f"อัปโหลด {len(files)} รายการ ({n} ไฟล์รายงาน)"
        return ds
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def export_bytes(ds):
    tmp = tempfile.mkdtemp(prefix="vrx_")
    try:
        fn = core.export_excel(ds, tmp)
        with open(fn, "rb") as fh:
            return os.path.basename(fn), fh.read()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def money(v):
    return f"{v:,.2f}"


def kpi(col, label, value, sub="", cls="blue"):
    col.markdown(f'<div class="kpi {cls}"><div class="lbl">{label}</div><div class="val">{value}</div>'
                 f'<div class="sub">{sub}</div></div>', unsafe_allow_html=True)


def to_df(key, rows):
    cols = core.COLS[key]
    df = pd.DataFrame(rows, columns=[c[0] for c in cols]) if rows else pd.DataFrame(columns=[c[0] for c in cols])
    for k, _, kind, _ in cols:
        if kind == "date":
            df[k] = pd.to_datetime(df[k], errors="coerce")
        elif kind in ("money", "num"):
            df[k] = pd.to_numeric(df[k], errors="coerce")
        elif kind == "km":  # แสดงเป็นข้อความ "11,048" ; ไม่มีเลขไมล์ = ช่องว่าง (ไม่ใส่ 0)
            df[k] = df[k].map(lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else f"{int(v):,}")
        elif kind == "int":
            df[k] = pd.to_numeric(df[k], errors="coerce").astype("Int64")
        else:
            df[k] = df[k].map(lambda v: "ใช่" if v is True else ("" if v is False or v is None else
                                                                 (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v))))
    return df


def col_config(key):
    cfg = {}
    for k, h, kind, w in core.COLS[key]:
        width = "small" if w <= 12 else ("medium" if w <= 24 else "large")
        if kind == "date":
            cfg[k] = st.column_config.DateColumn(h, format="DD/MM/YYYY", width=width)
        elif kind == "money":
            cfg[k] = st.column_config.NumberColumn(h, format="%,.2f", width=width)
        elif kind == "num":
            cfg[k] = st.column_config.NumberColumn(h, format="%,.4g", width=width)
        elif kind == "int":
            cfg[k] = st.column_config.NumberColumn(h, format="%d", width=width)
        elif kind == "km":
            cfg[k] = st.column_config.TextColumn(h, width=width)
        else:
            cfg[k] = st.column_config.TextColumn(h, width=width)
    return cfg


def is_issue(key, df):
    if key in ("issues", "wo_hdr", "wo_lines"):
        return df["confirm_flag"] == "No"
    if key == "files":
        return df["status"] != "PASSED"
    if key == "flags":
        return df["level"].isin(["No", "ข้าม"])
    return pd.Series(False, index=df.index)


def style_rows(key, df):
    mask = is_issue(key, df)
    dup = df["status"].astype(str).str.startswith("ซ้ำ") if key == "files" else pd.Series(False, index=df.index)

    def f(row):
        if dup.loc[row.name]:
            return ["color:#9aa5b1"] * len(row)
        if mask.loc[row.name]:
            return ["background-color:#fde8dc"] * len(row)
        return [""] * len(row)
    sty = df.style.apply(f, axis=1)
    return sty


# ----------------------------------------------------------------------------- state
ss = st.session_state
ss.setdefault("src", core.DEFAULT_SRC)
ss.setdefault("out", core.DEFAULT_OUT)
ss.setdefault("ds", None)
ss.setdefault("last_export", None)
ss.setdefault("export_blob", None)

# ----------------------------------------------------------------------------- sidebar
with st.sidebar:
    st.markdown("### 🚛 Vehicle Report Importer")
    st.caption(f"Web v{WEB_VERSION} · engine v{core.APP_VERSION}")

    uploads = None
    if IS_LOCAL:
        mode = st.radio("แหล่งข้อมูล", ["โฟลเดอร์ในเครื่อง", "อัปโหลดไฟล์"], horizontal=True)
    else:
        mode = "อัปโหลดไฟล์"
    if mode == "โฟลเดอร์ในเครื่อง":
        st.markdown("**① โฟลเดอร์รายงานรายคัน**")
        c1, c2 = st.columns([4, 1.4])
        if c2.button("📁", key="b_src", help="Browse…", **FULL):
            p = pick_folder(ss.src)
            if p:
                ss.src = p
        c1.text_input("โฟลเดอร์รายงาน", key="src", label_visibility="collapsed")
    else:
        st.markdown("**① อัปโหลดไฟล์รายงานรายคัน**")
        uploads = st.file_uploader("เลือกไฟล์ .xls / .xlsx หลายไฟล์ หรือ .zip ทั้งโฟลเดอร์",
                                   type=["xls", "xlsx", "zip"], accept_multiple_files=True)
        st.caption("แนะนำ: บีบอัดโฟลเดอร์ Vehicle_Reports เป็น .zip แล้วอัปโหลดไฟล์เดียว")
    dedupe = st.toggle("ตัดไฟล์ซ้ำ (เนื้อหาเหมือนกัน)", value=True)
    load = st.button("อ่านข้อมูล", type="primary", icon="🔍", **FULL)

    st.divider()
    st.markdown("**② Export Excel**")
    if IS_LOCAL:
        c1, c2 = st.columns([4, 1.4])
        if c2.button("📁", key="b_out", help="Browse…", **FULL):
            p = pick_folder(ss.out)
            if p:
                ss.out = p
        c1.text_input("โฟลเดอร์ Export", key="out", label_visibility="collapsed")
    export = st.button("Export Excel", icon="📤", disabled=ss.ds is None, **FULL)
    if IS_LOCAL and ss.last_export and os.path.exists(ss.last_export):
        st.success(f"บันทึกแล้ว\n\n`{ss.last_export}`")
        with open(ss.last_export, "rb") as fh:
            st.download_button("ดาวน์โหลดไฟล์นี้", fh.read(), file_name=os.path.basename(ss.last_export),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               icon="⬇️", **FULL)
    if not IS_LOCAL and ss.export_blob:
        name, blob = ss.export_blob
        st.success("สร้างไฟล์ Excel แล้ว")
        st.download_button(f"ดาวน์โหลด {name}", blob, file_name=name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           icon="⬇️", type="primary", **FULL)
    st.divider()
    st.caption("อ่านไฟล์ต้นทางแบบอ่านอย่างเดียว · ไม่แตะฐานข้อมูล Fleet_DB · "
               "แถว Flag No ต้องให้ DATA_OWNER ตรวจก่อนนำเข้าฐานข้อมูลหลัก")
    if not IS_LOCAL:
        st.caption("โหมด Cloud: ไฟล์ที่อัปโหลดถูกอ่านแล้วลบทันที ไม่เก็บบนเซิร์ฟเวอร์ · ข้อมูลอยู่ในหน่วยความจำเฉพาะ session นี้")
        if st.button("ล้างข้อมูลใน session", icon="🧹", **FULL):
            for k_ in ("ds", "export_blob", "last_export"):
                ss[k_] = None
            st.rerun()

# ----------------------------------------------------------------------------- actions
if load:
    bar = st.progress(0.0, text="กำลังอ่าน…")
    prog = lambda n, t, f: bar.progress(n / t, text=f"กำลังอ่าน {n}/{t} · {f}")
    if mode == "โฟลเดอร์ในเครื่อง":
        folder = ss.src.strip()
        if not os.path.isdir(folder):
            bar.empty()
            st.sidebar.error("ไม่พบโฟลเดอร์รายงาน")
        else:
            ss.ds = core.scan_folder(folder, progress=prog, dedupe=dedupe)
    else:
        if not uploads:
            bar.empty()
            st.sidebar.error("กรุณาอัปโหลดไฟล์ก่อน")
        else:
            ss.ds = scan_uploads(uploads, dedupe, prog)
    if ss.ds is not None:
        ss.last_export = None
        ss.export_blob = None
        bar.empty()
        st.rerun()

if export and ss.ds is not None:
    with st.spinner("กำลังเขียน Excel…"):
        try:
            if IS_LOCAL:
                ss.last_export = core.export_excel(ss.ds, ss.out.strip() or core.DEFAULT_OUT)
            else:
                ss.export_blob = export_bytes(ss.ds)
            st.toast("Export Excel เรียบร้อย", icon="✅")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"Export ไม่สำเร็จ: {e}")

# ----------------------------------------------------------------------------- main
st.markdown('<div class="app-hero"><h1>รายงานซ่อม &amp; เบิกพัสดุรายคัน</h1>'
            '<p>Finish Goods Takeoff and Work Order Report by Vehicle · สปีด อินเตอร์ ทรานสปอร์ต</p></div>',
            unsafe_allow_html=True)

ds = ss.ds
if ds is None:
    _hint = "เลือกโฟลเดอร์รายงานทางซ้าย" if IS_LOCAL else "อัปโหลดไฟล์รายงาน (.xls/.xlsx หรือ .zip) ทางซ้าย"
    st.markdown('<div class="empty"><div style="font-size:2.4rem">📂</div>'
                '<h3 style="margin:.3rem 0">ยังไม่มีข้อมูล</h3>'
                f'{_hint} แล้วกด <b>อ่านข้อมูล</b><br>'
                'จากนั้นตรวจแท็บ <b>ประเด็น/Flag</b> และกด <b>Export Excel</b></div>', unsafe_allow_html=True)
    st.stop()

s = ds["summary"]
dr = f"{s['date_min']:%d/%m/%Y} – {s['date_max']:%d/%m/%Y}" if s["date_min"] else "-"
ok_all = s["chk_issue_total"] and s["chk_wo_total"]
st.markdown(
    f'<span class="badge b-info">📁 {s["folder"]}</span>'
    f'<span class="badge b-info">🗓 {dr}</span>'
    f'<span class="badge b-info">อ่านเมื่อ {s["run_at"]:%d/%m/%Y %H:%M}</span>'
    f'<span class="badge {"b-ok" if ok_all else "b-no"}">ผลรวมบรรทัด = ยอดรายงาน: {"PASSED" if ok_all else "FAILED"}</span>',
    unsafe_allow_html=True)
st.write("")

k = st.columns(6)
kpi(k[0], "ไฟล์รายงาน", f"{s['n_reports']:,}", f"ข้าม {s['n_skipped']} · ซ้ำ {s['n_dup']} · ใช้ {s['n_used']}")
kpi(k[1], "ทะเบียนรถ", f"{s['n_plates']:,}", "คัน")
kpi(k[2], "กระทบยอดผ่านครบ", f"{s['n_pass']}/{s['n_reports']}", "6 จุดต่อไฟล์",
    "ok" if s["n_pass"] == s["n_reports"] else "warn")
kpi(k[3], "ยอดเบิกพัสดุ (WX)", money(s["wx"]), f"{s['n_issue_docs']:,} ใบ · {s['n_issue_lines']:,} บรรทัด")
kpi(k[4], "ยอดสั่งซ่อมภายนอก (W/O)", money(s["wo"]), f"{s['n_wo']:,} ใบ · {s['n_wo_lines']:,} บรรทัด")
kpi(k[5], "ยอดรวม (บาท)", money(s["all"]), f"Flag No {s['n_flag_no']:,} บรรทัด · ประเด็น {len(ds['flags'])}",
    "warn" if s["n_flag_no"] else "ok")
st.write("")

# ----------------------------------------------------------------------------- tabs
TAB_ORDER = [("dash", "📊 ภาพรวม")] + [(k_, t) for k_, t in core.SHEETS]
labels = []
for key, title in TAB_ORDER:
    if key == "dash":
        labels.append(title)
    else:
        n = len(ds[key])
        labels.append(f"{title} ({n:,})" if key != "flags" else f"⚠️ {title} ({n:,})")
tabs = st.tabs(labels)

with tabs[0]:
    issues = to_df("issues", ds["issues"])
    wol = to_df("wo_lines", ds["wo_lines"])
    m = pd.concat([
        issues.assign(ประเภท="เบิกพัสดุ")[["date", "total", "ประเภท"]],
        wol.assign(ประเภท="สั่งซ่อมภายนอก")[["date", "total", "ประเภท"]],
    ])
    m["เดือน"] = m["date"].dt.strftime("%Y-%m")
    veh = to_df("vehicles", ds["vehicles"]).sort_values("all", ascending=False).head(15)

    import altair as alt
    CAT = ["เบิกพัสดุ", "สั่งซ่อมภายนอก"]
    COLORS = alt.Scale(domain=CAT, range=["#1f6fb2", "#e8833a"])
    ml = (m.dropna(subset=["เดือน"]).groupby(["เดือน", "ประเภท"], as_index=False)["total"].sum())
    vl = veh.melt(id_vars=["plate", "all"], value_vars=["wx", "wo"], var_name="k", value_name="ยอด")
    vl["ประเภท"] = vl["k"].map({"wx": "เบิกพัสดุ", "wo": "สั่งซ่อมภายนอก"})
    legend = alt.Legend(orient="bottom", title=None)

    c1, c2 = st.columns([1.15, 1])
    with c1:
        st.markdown("##### ค่าใช้จ่ายรายเดือน (บาท)")
        ch = (alt.Chart(ml).mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
              .encode(x=alt.X("เดือน:N", title=None, axis=alt.Axis(labelAngle=0)),
                      y=alt.Y("sum(total):Q", title=None, axis=alt.Axis(format=",.0f")),
                      color=alt.Color("ประเภท:N", scale=COLORS, legend=legend),
                      order=alt.Order("ประเภท:N", sort="descending"),
                      tooltip=["เดือน", "ประเภท", alt.Tooltip("total:Q", title="ยอด", format=",.2f")])
              .properties(height=340))
        st.altair_chart(ch, **FULL)
        st.caption("เดือนตามวันที่เอกสาร · เดือนล่าสุดอาจยังไม่ครบเดือน")
    with c2:
        st.markdown("##### 15 คันที่ค่าใช้จ่ายสูงสุด")
        order = veh["plate"].tolist()
        ch = (alt.Chart(vl).mark_bar()
              .encode(y=alt.Y("plate:N", sort=order, title=None, axis=alt.Axis(labelOverlap=False)),
                      x=alt.X("sum(ยอด):Q", title=None, axis=alt.Axis(format=",.0f")),
                      color=alt.Color("ประเภท:N", scale=COLORS, legend=legend),
                      order=alt.Order("ประเภท:N", sort="descending"),
                      tooltip=[alt.Tooltip("plate:N", title="ทะเบียน"), "ประเภท",
                               alt.Tooltip("ยอด:Q", format=",.2f"), alt.Tooltip("all:Q", title="ยอดรวมคัน", format=",.2f")])
              .properties(height=340))
        st.altair_chart(ch, **FULL)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("##### ผู้ขาย/อู่ — ยอดสั่งซ่อมสูงสุด")
        top_v = (wol.groupby("vendor_en", dropna=False)
                 .agg(ใบสั่งซ่อม=("doc_no", "nunique"), ยอด=("total", "sum")).sort_values("ยอด", ascending=False)
                 .head(10).reset_index().rename(columns={"vendor_en": "ผู้ขาย"}))
        st.dataframe(top_v, hide_index=True, **FULL,
                     column_config={"ยอด": st.column_config.NumberColumn(format="%,.2f")})
    with c2:
        st.markdown("##### คลังที่เบิกพัสดุ")
        top_s = (issues.groupby("store").agg(บรรทัด=("total", "size"), ยอด=("total", "sum"))
                 .sort_values("ยอด", ascending=False).head(10).reset_index().rename(columns={"store": "คลัง"}))
        st.dataframe(top_s, hide_index=True, **FULL,
                     column_config={"ยอด": st.column_config.NumberColumn(format="%,.2f")})

for i, (key, title) in enumerate(TAB_ORDER[1:], start=1):
    with tabs[i]:
        df = to_df(key, ds[key])
        f1, f2, f3 = st.columns([2.2, 2.2, 1.2])
        q = f1.text_input("ค้นหา", key=f"q_{key}", placeholder="ทะเบียน / เลขเอกสาร / ชื่อรายการ / ผู้ขาย…",
                          label_visibility="collapsed")
        plates = sorted(df["plate"].dropna().unique()) if "plate" in df else []
        sel = f2.multiselect("ทะเบียน", plates, key=f"p_{key}", placeholder="กรองทะเบียน (เลือกได้หลายคัน)",
                             label_visibility="collapsed") if plates else []
        only = f3.toggle("เฉพาะ Flag No / ประเด็น", key=f"o_{key}")
        view = df
        if only:
            view = view[is_issue(key, view)]
        if sel:
            view = view[view["plate"].isin(sel)]
        if q:
            ql = q.strip().lower()
            view = view[view.astype(str).apply(lambda r: r.str.lower().str.contains(ql, regex=False)).any(axis=1)]

        money_cols = [c[0] for c in core.COLS[key] if c[2] == "money" and c[0] in ("total", "net", "wx", "wo", "all")]
        cap = f"แสดง {len(view):,} / {len(df):,} แถว"
        if money_cols:
            sv = view[~view["status"].astype(str).str.startswith("ซ้ำ")] if key == "files" else view
            cap += " · " + " · ".join(f"{dict((c[0], c[1]) for c in core.COLS[key])[c]} {sv[c].sum():,.2f}"
                                      for c in money_cols)
            if key == "files":
                cap += " (ยอดไม่รวมไฟล์ซ้ำ)"
        st.caption(cap)
        data = style_rows(key, view.reset_index(drop=True)) if len(view) <= 20000 else view
        st.dataframe(data, hide_index=True, **FULL, height=560, column_config=col_config(key))
