# -*- coding: utf-8 -*-
"""หน้า 'PM Plan รายเดือน' — ใช้ใน vehicle_report_web.py"""
import os, io, calendar, datetime as dt

import pandas as pd
import streamlit as st

import pm_plan as pm
import repair_request_db as rr
from repair_request_page import db_selector

TH_MONTH = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]

PLAN_COLS = [("asset_code", "ทะเบียน"), ("plan_status", "สถานะในแผน"), ("due_date", "วันครบกำหนด"), ("due_by", "เกณฑ์ที่ถึงก่อน"),
             ("pm_type", "PM ครั้งล่าสุดที่"), ("interval_km", "รอบ (กม.)"), ("last_pm", "วันที่ PM ล่าสุด"), ("last_pm_doc", "เอกสาร PM"),
             ("run_km", "วิ่งแล้วตั้งแต่ PM (กม.)"), ("remain_km", "เหลือ (กม.)"), ("km_day", "กม./วัน"), ("rate_src", "ที่มา กม./วัน"),
             ("due_time", "ครบ 6 เดือน"), ("due_km_date", "ครบระยะ (ประมาณ)"), ("last_center", "เข้าศูนย์ล่าสุด"), ("last_yard", "เปลี่ยนถ่ายที่ลานล่าสุด"),
             ("last_repair", "แจ้งซ่อมล่าสุด"), ("last_km", "เลขไมล์ล่าสุด"), ("last_km_date", "วันที่เลขไมล์ล่าสุด"), ("notes", "หมายเหตุ")]
DATE_COLS = {"due_date", "last_pm", "due_time", "due_km_date", "last_center", "last_yard", "last_repair", "last_km_date"}


def _fmt(df):
    out = df.copy()
    for c in DATE_COLS & set(out.columns):
        out[c] = pd.to_datetime(out[c], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
    for c in ("run_km", "remain_km", "last_km", "interval_km"):
        if c in out.columns:
            out[c] = out[c].map(lambda v: "" if pd.isna(v) else f"{int(v):,}")
    if "km_day" in out.columns:
        out["km_day"] = out["km_day"].map(lambda v: "" if pd.isna(v) else f"{v:,.0f}")
    cols = [c for c, _ in PLAN_COLS if c in out.columns]
    out = out[cols].astype(object).where(out[cols].notna(), "")
    return out.rename(columns=dict(PLAN_COLS))


def _plan_for_month(rows, y, m, with_overdue, with_nohist):
    m0, m1 = dt.date(y, m, 1), dt.date(y, m, calendar.monthrange(y, m)[1])
    out = []
    for r in rows:
        if r["due_date"] is None:
            if with_nohist:
                out.append({**r, "plan_status": "ไม่มีประวัติ PM"})
            continue
        d = dt.date.fromisoformat(r["due_date"])
        if m0 <= d <= m1:
            out.append({**r, "plan_status": "ครบกำหนดในเดือน"})
        elif d < m0 and with_overdue:
            out.append({**r, "plan_status": "เกินกำหนด (ค้าง)"})
    order = {"เกินกำหนด (ค้าง)": 0, "ครบกำหนดในเดือน": 1, "ไม่มีประวัติ PM": 2}
    out.sort(key=lambda r: (order[r["plan_status"]], r["due_date"] or "9999", r["asset_code"]))
    return out


def _excel(plan_df, all_df, crit, title):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, df in ((f"แผน PM {title}", plan_df), ("สถานะ PM ทุกคัน", all_df)):
            x = df.copy()
            for c in DATE_COLS & set(x.columns):
                x[c] = pd.to_datetime(x[c], errors="coerce")
            cols = [c for c, _ in PLAN_COLS if c in x.columns]
            x[cols].rename(columns=dict(PLAN_COLS)).to_excel(xw, sheet_name=name[:31], index=False)
        crit.to_excel(xw, sheet_name="เกณฑ์การคำนวณ", index=False)
        fill = PatternFill("solid", fgColor="DCE9F5")
        date_heads = {h for c, h in PLAN_COLS if c in DATE_COLS}
        for ws in xw.book.worksheets:
            ws.freeze_panes = "B2" if ws.title != "เกณฑ์การคำนวณ" else "A2"
            for i, cell in enumerate(ws[1], 1):
                cell.font = Font(bold=True); cell.fill = fill
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ws.column_dimensions[get_column_letter(i)].width = 14
                if cell.value in date_heads:
                    for c in ws[get_column_letter(i)][1:]:
                        c.number_format = "dd/mm/yyyy"
                if cell.value in ("วิ่งแล้วตั้งแต่ PM (กม.)", "เหลือ (กม.)", "รอบ (กม.)", "เลขไมล์ล่าสุด", "กม./วัน"):
                    for c in ws[get_column_letter(i)][1:]:
                        c.number_format = "#,##0"
                if cell.value in ("หมายเหตุ", "รายละเอียด"):
                    ws.column_dimensions[get_column_letter(i)].width = 60
            if ws.max_row > 1 and ws.title != "เกณฑ์การคำนวณ":
                ws.auto_filter.ref = ws.dimensions
    return buf.getvalue()


def render_pm(is_local, user, full, app_dir, kpi_card):
    ss = st.session_state
    db = db_selector(is_local, app_dir, full)
    st.markdown('<div class="app-hero"><h1>PM Plan รายเดือน</h1><p>หัวลาก · เข้าศูนย์บริการทุก 20,000 กม. · เปลี่ยนถ่ายที่ลานทุก 30,000 กม. · '
                'หรือครบ 6 เดือน (อย่างใดถึงก่อน) · กม./วันจากประวัติเลขไมล์และเลขไมล์ตอนเข้าบริการ</p></div>', unsafe_allow_html=True)

    o = ss.setdefault("pm_opt", dict(pm.DEFAULTS))
    nsvc = pm.fetch(db, "SELECT COUNT(*) n, SUM(pm_type IS NOT NULL) p FROM service_doc")[0]
    res = pm.compute_plan(db, **o)
    rows = res["rows"]

    t_plan, t_all, t_imp, t_hist = st.tabs(["📅 แผนรายเดือน", "🚛 สถานะรายคัน", "📥 นำเข้าประวัติบริการ", "🔎 ประวัติรายคัน"])

    # ------------------------------------------------------------------ plan
    with t_plan:
        if not nsvc["n"]:
            st.warning("ยังไม่มีประวัติบริการในฐาน — นำเข้าไฟล์ Export จากหน้า 'รายงานซ่อมรายคัน' ที่แท็บ 📥 นำเข้าประวัติบริการ ก่อน")
        today = dt.date.today()
        months = []
        y, m = today.year, today.month
        for k in range(-2, 7):
            yy, mm = divmod(m - 1 + k, 12)
            months.append((y + yy, mm + 1))
        c1, c2, c3 = st.columns([1.3, 1.2, 1.2])
        sel = c1.selectbox("เดือนของแผน", months, index=2, format_func=lambda t: f"{TH_MONTH[t[1]]} {t[0] + 543}", key="pm_month")
        with_over = c2.toggle("รวมคันที่เกินกำหนดค้าง", value=True, key="pm_over")
        with_nohist = c3.toggle("รวมคันที่ไม่มีประวัติ PM", value=True, key="pm_nohist")
        plan = _plan_for_month(rows, sel[0], sel[1], with_over, with_nohist)
        k = st.columns(4)
        n_in = sum(r["plan_status"] == "ครบกำหนดในเดือน" for r in plan)
        n_over = sum(r["plan_status"] == "เกินกำหนด (ค้าง)" for r in plan)
        n_no = sum(r["due_date"] is None for r in rows)
        kpi_card(k[0], f"ครบกำหนด {TH_MONTH[sel[1]]} {sel[0] + 543}", f"{n_in:,}", "คัน")
        kpi_card(k[1], "เกินกำหนดค้าง", f"{n_over:,}", "คัน (ก่อนเดือนนี้)", "warn" if n_over else "ok")
        kpi_card(k[2], "ไม่มีประวัติ PM", f"{n_no:,}", "คัน — ไม่พบในข้อมูลที่นำเข้า", "warn" if n_no else "ok")
        kpi_card(k[3], "กม./วัน ค่ากลาง", f"{res['fleet_rate']:,.0f}" if res["fleet_rate"] else "-", f"หัวลาก {len(rows):,} คัน")
        st.write("")
        pdf = pd.DataFrame(plan)
        if len(pdf):
            st.dataframe(_fmt(pdf), hide_index=True, height=520, **full)
        else:
            st.info("ไม่มีคันที่ครบกำหนดในเดือนที่เลือก")
        crit = pd.DataFrame({"รายการ": ["PM ศูนย์บริการ", "PM ลาน", "เกณฑ์ครบกำหนด", "รอบถัดไป", "กม./วัน", "ตัดเลขไมล์ผิดปกติ",
                                        "เลขไมล์ 6 หลัก", "คำนวณเมื่อ", "ฐานข้อมูล"],
                             "รายละเอียด": [f"W/O ที่มีน้ำมันเครื่อง/ไส้กรองน้ำมันเครื่อง → ทุก {o['center_km']:,} กม. หรือ {o['months']} เดือน",
                                            f"ใบเบิก WX ที่มีน้ำมันเครื่อง → ทุก {o['yard_km']:,} กม. หรือ {o['months']} เดือน",
                                            "ถึงอย่างใดอย่างหนึ่งก่อน (เวลา หรือ ระยะ)",
                                            "ใช้รอบตามประเภท PM ครั้งล่าสุดของคันนั้น (ศูนย์ → 20,000 · ลาน → 30,000)",
                                            f"ระยะรวม ÷ จำนวนวัน จากเลขไมล์ทุกแหล่ง (ไฟล์เลขไมล์ + W/O + WX) ใน {o['window_days']} วันล่าสุด ; "
                                            f"ใช้ได้ {o['min_km_day']}–{o['max_km_day']:,} กม./วัน ไม่เช่นนั้นใช้ค่ากลางของกองรถ",
                                            "ตัดค่าที่ขัดกับค่าอื่น (ลดลง / วิ่งเกินเพดานต่อวัน) และช่วงที่เร็วกว่า 3 เท่าของค่ากลาง ; ไม่ใช้เลขไมล์สถานะ 'ต้องตรวจ' หรือ < 1,000",
                                            "นับว่าวนรอบเฉพาะจาก ≥ 900,000 ไป < 100,000",
                                            dt.datetime.now().strftime("%d/%m/%Y %H:%M"), os.path.basename(db)]})
        title = f"{sel[0]}-{sel[1]:02d}"
        if st.button("สร้างไฟล์ Export แผน PM", icon="📊", key="pm_exp"):
            ss["pm_blob"] = (f"PM_Plan_{title}_{dt.datetime.now():%Y%m%d_%H%M}.xlsx", _excel(pdf, pd.DataFrame(rows), crit, title))
        if ss.get("pm_blob"):
            st.download_button(f"ดาวน์โหลด {ss['pm_blob'][0]}", ss["pm_blob"][1], file_name=ss["pm_blob"][0], icon="⬇️", type="primary",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="pm_dl")
        with st.expander("⚙️ เกณฑ์การคำนวณ"):
            c = st.columns(4)
            n1 = c[0].number_input("ศูนย์บริการ ทุก (กม.)", 1000, 200000, int(o["center_km"]), 1000)
            n2 = c[1].number_input("ลาน ทุก (กม.)", 1000, 200000, int(o["yard_km"]), 1000)
            n3 = c[2].number_input("หรือครบ (เดือน)", 1, 24, int(o["months"]), 1)
            n4 = c[3].number_input("เพดาน กม./วัน", 200, 3000, int(o["max_km_day"]), 100)
            if (n1, n2, n3, n4) != (o["center_km"], o["yard_km"], o["months"], o["max_km_day"]):
                ss["pm_opt"] = {**o, "center_km": n1, "yard_km": n2, "months": n3, "max_km_day": n4}
                ss.pop("pm_blob", None)
                st.rerun()
            st.dataframe(crit, hide_index=True, **full)
            st.caption("ค่าที่เปลี่ยนใช้เฉพาะการเปิดครั้งนี้ — ค่าเริ่มต้นตามที่ Owner กำหนด (20,000 / 30,000 กม. / 6 เดือน)")

    # ------------------------------------------------------------------ all vehicles
    with t_all:
        adf = pd.DataFrame(rows)
        if len(adf):
            c1, c2, c3 = st.columns([1.4, 1.4, 1.6])
            fs = c1.multiselect("สถานะ", ["เกินกำหนด", "ครบกำหนดเดือนนี้", "ยังไม่ถึง", "ไม่มีประวัติ PM"], key="pm_fs", placeholder="ทุกสถานะ")
            ft = c2.multiselect("PM ครั้งล่าสุดที่", ["ศูนย์บริการ", "ลาน", "ยังไม่มีประวัติ PM"], key="pm_ft", placeholder="ทั้งหมด")
            fq = c3.text_input("ค้นหาทะเบียน", key="pm_fq")
            x = adf
            if fs:
                x = x[x["status"].isin(fs)]
            if ft:
                x = x[x["pm_type"].isin(ft)]
            if fq:
                x = x[x["asset_code"].str.contains(fq.strip(), regex=False)]
            x = x.sort_values(["due_date", "asset_code"], na_position="first").rename(columns={"status": "plan_status"})
            st.caption(f"แสดง {len(x):,} คัน · สถานะเทียบกับวันนี้ ({dt.date.today():%d/%m/%Y})")
            st.dataframe(_fmt(x).rename(columns={"สถานะในแผน": "สถานะ"}), hide_index=True, height=560, **full)

    # ------------------------------------------------------------------ import
    with t_imp:
        st.markdown("นำเข้า **ไฟล์ Export จากหน้า 🚛 รายงานซ่อมรายคัน** (`Vehicle_Reports_*.xlsx` — ชีต ใบสั่งซ่อม_หัว / ใบสั่งซ่อม_บรรทัด / เบิกพัสดุ) "
                    "เพื่อเก็บเป็นประวัติบริการ (วันที่ · เลขไมล์ · ผู้ขาย · รายการน้ำมันเครื่อง) — กันซ้ำด้วยเลขที่เอกสาร")
        if ss.get("pm_msg"):
            st.success(ss.pop("pm_msg"))
        up = st.file_uploader("ไฟล์ Export รายงานซ่อมรายคัน (.xlsx)", type=["xlsx"], key="pm_up")
        if st.button("ตรวจไฟล์ (ยังไม่บันทึก)", icon="🔍", disabled=up is None, key="pm_chk") and up is not None:
            data = up.getvalue()
            with st.spinner("กำลังอ่าน…"):
                r = pm.read_report_export(io.BytesIO(data))
            ss["pm_an"] = dict(name=up.name, sha=rr.file_sha256(data), r=r, plan=pm.analyze_service(db, r["docs"]) if r["docs"] else None)
        an = ss.get("pm_an")
        if an and up is not None and an["name"] == up.name:
            r = an["r"]
            if r["errors"]:
                for e in r["errors"]:
                    st.error(e)
            else:
                s_ = r["stats"]
                st.caption(f"อ่านได้ {s_['docs']:,} เอกสาร (W/O {s_['wo']:,} · WX {s_['wx']:,}) · {s_['plates']:,} คัน · "
                           f"PM ศูนย์บริการ {s_['pm_center']:,} · PM ลาน {s_['pm_yard']:,}")
                p = an["plan"]
                k = st.columns(3)
                kpi_card(k[0], "ใหม่", f"{len(p['new']):,}", "เอกสาร")
                kpi_card(k[1], "อัปเดต", f"{len(p['update']):,}", "เอกสารเดิมที่ค่าเปลี่ยน")
                kpi_card(k[2], "ไม่เปลี่ยน", f"{len(p['unchanged']):,}", "เอกสาร")
                ev = pd.DataFrame([d for d in p["new"] + p["update"] if d["pm_type"]])
                if len(ev):
                    st.markdown("**ตัวอย่างเอกสารที่นับเป็น PM**")
                    st.dataframe(ev[["asset_code", "doc_kind", "doc_no", "doc_date", "km", "km_status", "pm_type", "pm_items", "vendor"]]
                                 .rename(columns={"asset_code": "ทะเบียน", "doc_kind": "ประเภท", "doc_no": "เลขที่", "doc_date": "วันที่",
                                                  "km": "เลขไมล์", "km_status": "สถานะเลขไมล์", "pm_type": "PM", "pm_items": "รายการ",
                                                  "vendor": "ผู้ขาย/คลัง"}).head(300), hide_index=True, height=300, **full)
                n_act = len(p["new"]) + len(p["update"])
                ok = st.checkbox("ตรวจผลแล้ว ยืนยันนำเข้า", key="pm_ok")
                if st.button("ยืนยันนำเข้าประวัติบริการ", type="primary", icon="💾", disabled=not ok or n_act == 0, key="pm_go"):
                    res_ = pm.commit_service(db, r["docs"], an["name"], an["sha"], user or ("local" if is_local else "user"))
                    ss.pop("pm_an", None); ss.pop("pm_blob", None)
                    ss["pm_msg"] = (f"นำเข้าแล้ว (batch {res_['batch_id']}): ใหม่ {res_['rows_inserted']:,} · อัปเดต {res_['rows_updated']:,} · "
                                    f"ไม่เปลี่ยน {res_['rows_unchanged']:,} — แผน PM คำนวณใหม่แล้ว"
                                    + (f"  \nสำรอง: `{res_['backup']}`" if res_["backup"] else ""))
                    st.rerun()
                if n_act == 0:
                    st.info("ไฟล์นี้อยู่ในฐานข้อมูลครบแล้ว")
        st.markdown("**ประวัติการนำเข้า**")
        st.dataframe(pd.DataFrame(pm.fetch(db, "SELECT batch_id, file_name ไฟล์, imported_by ผู้นำเข้า, imported_at เวลา, rows_read อ่าน, "
                                               "rows_inserted ใหม่, rows_updated อัปเดต, rows_unchanged ไม่เปลี่ยน, backup_path ไฟล์สำรอง "
                                               "FROM service_batch ORDER BY batch_id DESC")), hide_index=True, **full)

    # ------------------------------------------------------------------ history
    with t_hist:
        plates = sorted(r["asset_code"] for r in rows)
        if plates:
            p_ = st.selectbox("ทะเบียน", plates, key="pm_hp")
            row = next(r for r in rows if r["asset_code"] == p_)
            k = st.columns(4)
            kpi_card(k[0], "สถานะ", row["status"], row["due_by"] or "")
            kpi_card(k[1], "วันครบกำหนด", pd.to_datetime(row["due_date"]).strftime("%d/%m/%Y") if row["due_date"] else "-", row["pm_type"])
            kpi_card(k[2], "กม./วัน", f"{row['km_day']:,.0f}" if row["km_day"] else "-", row["rate_src"])
            kpi_card(k[3], "วิ่งแล้วตั้งแต่ PM", f"{row['run_km']:,}" if row.get("run_km") is not None else "-", row.get("run_method") or "")
            if row["notes"]:
                st.caption("หมายเหตุ: " + row["notes"])
            evs = pd.DataFrame(pm.fetch(db, "SELECT doc_date วันที่, pm_type PM, doc_kind ประเภท, doc_no เลขที่, km เลขไมล์, km_status สถานะเลขไมล์, "
                                            "pm_items รายการ, vendor 'ผู้ขาย/คลัง' FROM service_doc WHERE asset_code=? AND pm_type IS NOT NULL "
                                            "ORDER BY doc_date DESC", p_))
            st.markdown("**ประวัติ PM**")
            st.dataframe(evs, hide_index=True, **full) if len(evs) else st.info("ไม่พบ PM ในข้อมูลที่นำเข้า")
            kh = pd.DataFrame(pm.km_history(db, p_))
            st.markdown("**ประวัติเลขไมล์ (ทุกแหล่ง)**")
            if len(kh):
                used = kh[kh["used"] == "ใช้คำนวณ"]
                if len(used) > 1:
                    st.line_chart(used.set_index(pd.to_datetime(used["date"]))["km"], height=220)
                kh["date"] = pd.to_datetime(kh["date"]).dt.strftime("%d/%m/%Y")
                st.dataframe(kh.rename(columns={"date": "วันที่", "km": "เลขไมล์", "source": "แหล่ง", "doc_no": "เลขที่เอกสาร", "used": "การใช้"}),
                             hide_index=True, **full)
            else:
                st.info("ไม่มีเลขไมล์")
