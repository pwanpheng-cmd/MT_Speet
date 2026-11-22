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



def _old_plan(month):
    """แผน KPI เดิมของหัวลากเดือนนั้น — อ่านจาก data/local_settings.json (ผ่าน monthly_report) ; ไม่มีไฟล์ = 0"""
    try:
        import monthly_report as _mr
        return int((_mr.DEFAULT_KPI_PLANS.get("หัวลาก") or [0] * 12)[month - 1] or 0)
    except Exception:
        return 0

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


def _load_draft_pm_workbook(raw_bytes):
    """อ่านไฟล์ร่าง Plan_PM_หัวลาก_ใหม่_รอยืนยัน.xlsx (4 ชีต) พร้อมคำนวณค่าสรุปไม่พึ่ง cached formula ของ Excel"""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=False)
    out = dict(status=[], plan_2026=[], plan_2027=[], matrix=pd.DataFrame(), detail=pd.DataFrame(), hist=pd.DataFrame())

    # 1) ชีตสรุป
    s1_name = next((n for n in ("สรุป_Plan_PM_2026_2027", "สรุป_Plan_PM_รายเดือน") if n in wb.sheetnames), None)
    if s1_name:
        ws1 = wb[s1_name]
        for r in range(6, ws1.max_row + 1):
            a = ws1.cell(r, 1).value
            b = ws1.cell(r, 2).value
            if not a:
                continue
            sa = str(a).strip()
            if sa.startswith("2026-") and len(sa) == 7:
                ctr = int(b or 0)
                yrd = int(ws1.cell(r, 3).value or 0)
                ovd = int(ws1.cell(r, 4).value or 0)
                tot = ctr + yrd + ovd
                act_v = ws1.cell(r, 6).value
                act = int(act_v) if isinstance(act_v, (int, float)) else None
                pct = round(act * 100.0 / tot, 1) if (act is not None and tot > 0) else None
                rep_v = ws1.cell(r, 8).value
                rep = int(rep_v) if isinstance(rep_v, (int, float)) else None
                old_v = ws1.cell(r, 9).value
                old_p = int(old_v) if isinstance(old_v, (int, float)) else _old_plan(int(sa[5:7]))
                out["plan_2026"].append({
                    "เดือน": sa, "PM ศูนย์ (20k)": ctr, "PM ลาน (30k)": yrd, "ค้างยกยอด": ovd,
                    "รวม Plan PM ใหม่": tot, "ทำ PM จริง": act, "% ตามแผน PM": pct,
                    "ใบแจ้งซ่อมทั่วไป (Master)": rep, "Plan เดิมใน KPI": old_p,
                })
            elif sa.startswith("2027-") and len(sa) == 7:
                ctr = int(b or 0)
                yrd = int(ws1.cell(r, 3).value or 0)
                tot = ctr + yrd
                out["plan_2027"].append({
                    "เดือน": sa, "PM ศูนย์ (20k)": ctr, "PM ลาน (30k)": yrd,
                    "รวม Plan PM ใหม่": tot, "Plan เดิมใน KPI": _old_plan(int(sa[5:7])), "ส่วนต่าง (ใหม่-เดิม)": tot - _old_plan(int(sa[5:7])),
                })
            elif r <= 12 and isinstance(b, (int, float)) and not sa.startswith("รวม"):
                out["status"].append({"สถานะ": sa, "จำนวน (คัน)": int(b), "หมายเหตุ": str(ws1.cell(r, 3).value or "")})

    # 2) ตารางรายคัน (24 เดือน)
    s2_name = next((n for n in ("ตาราง_PM_217คัน_2026_2027", "ตาราง_PM_217คัน_รายเดือน") if n in wb.sheetnames), None)
    if s2_name:
        ws2 = wb[s2_name]
        hdrs = [str(ws2.cell(3, c).value or "").strip() for c in range(1, ws2.max_column + 1)]
        rows2 = []
        for r in range(4, ws2.max_row + 1):
            p = ws2.cell(r, 1).value
            if not p or str(p).startswith(("รวม", "ทำจริง")):
                continue
            rows2.append([str(ws2.cell(r, c).value or "") for c in range(1, len(hdrs) + 1)])
        if rows2:
            out["matrix"] = pd.DataFrame(rows2, columns=hdrs)

    # 3) รายละเอียดรายคัน
    if "รายละเอียด_217คัน" in wb.sheetnames:
        ws3 = wb["รายละเอียด_217คัน"]
        hdrs3 = [str(ws3.cell(1, c).value or "").strip() for c in range(1, ws3.max_column + 1)]
        rows3 = [[ws3.cell(r, c).value for c in range(1, len(hdrs3) + 1)] for r in range(2, ws3.max_row + 1) if ws3.cell(r, 1).value]
        if rows3:
            out["detail"] = pd.DataFrame(rows3, columns=hdrs3)

    # 4) ประวัติ PM จริง 2026
    if "ประวัติ_PM_จริง_2026" in wb.sheetnames:
        ws4 = wb["ประวัติ_PM_จริง_2026"]
        hdrs4 = [str(ws4.cell(1, c).value or "").strip() for c in range(1, ws4.max_column + 1)]
        rows4 = [[ws4.cell(r, c).value for c in range(1, len(hdrs4) + 1)] for r in range(2, ws4.max_row + 1) if ws4.cell(r, 1).value]
        if rows4:
            out["hist"] = pd.DataFrame(rows4, columns=hdrs4)

    return out


def _df_to_excel_bytes(sheets_dict):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, df in sheets_dict.items():
            (df if len(df) else pd.DataFrame({"ไม่มีข้อมูล": []})).to_excel(xw, sheet_name=name[:31], index=False)
    return buf.getvalue()


@st.cache_data(show_spinner=False, ttl=600)
def _parse_vehicle_report_cached(file_path):
    import vehicle_report_app as vra
    return vra.parse_report(file_path)


def _find_vehicle_report_file(app_dir, plate):
    plate = str(plate or "").strip()
    if not plate:
        return None
    for folder in ("Vehicle_Reports", "ทะเบียนรถ"):
        d = os.path.join(app_dir, folder)
        if not os.path.isdir(d):
            continue
        cands = [
            f for f in sorted(os.listdir(d))
            if f.lower().endswith((".xls", ".xlsx")) and f.startswith(plate)
        ]
        if cands:
            return os.path.join(d, cands[0])
    return None


def _lookup_doc_detail(app_dir, plate, doc_no, doc_kind=None):
    """ดึงรายละเอียดรายการทั้งหมดของเอกสาร W/O หรือ WX จากไฟล์รายงานรายคันในเครื่อง"""
    fpath = _find_vehicle_report_file(app_dir, plate)
    if not fpath:
        return None
    rep = _parse_vehicle_report_cached(fpath)
    raw_no = str(doc_no or "").strip()
    # ตัด prefix เช่น 'W/O ' หรือ 'WX ' ถ้ามี
    norm_no = raw_no.split()[-1] if " " in raw_no else raw_no
    kind_u = str(doc_kind or "").upper()

    # 1) ค้นใน W/O
    if kind_u in ("", "W/O", "WO"):
        hdr = next((h for h in rep.get("wo_hdr", []) if str(h.get("doc_no", "")).strip() in (raw_no, norm_no)), None)
        lines = [l for l in rep.get("wo_lines", []) if str(l.get("doc_no", "")).strip() in (raw_no, norm_no)]
        if hdr or lines:
            return dict(
                kind="W/O (ใบสั่งซ่อมภายนอก)",
                file=os.path.basename(fpath),
                doc_no=raw_no,
                date=str(hdr.get("date") or (lines[0].get("date") if lines else "")),
                vendor=hdr.get("vendor_th") or hdr.get("vendor_en") or (lines[0].get("vendor_en") if lines else "") if hdr else "",
                mileage=hdr.get("mileage_km") if hdr else None,
                net=hdr.get("net") if hdr else sum(float(l.get("total") or 0) for l in lines),
                remark=hdr.get("remark") or "" if hdr else "",
                lines=[
                    {
                        "ลำดับ": l.get("line_no"),
                        "รหัสรายการ": l.get("item_code") or "",
                        "รายละเอียดรายการ": l.get("item_name") or "",
                        "จำนวน": l.get("qty"),
                        "ราคา/หน่วย": l.get("unit_price"),
                        "จำนวนเงิน (บาท)": l.get("total"),
                    }
                    for l in lines
                ],
            )

    # 2) ค้นใน WX (ใบเบิกพัสดุ)
    if kind_u in ("", "WX"):
        lines = [l for l in rep.get("issues", []) if str(l.get("doc_no", "")).strip() in (raw_no, norm_no)]
        if lines:
            first = lines[0]
            return dict(
                kind="WX (ใบเบิกพัสดุลาน)",
                file=os.path.basename(fpath),
                doc_no=raw_no,
                date=str(first.get("date") or ""),
                vendor=f"{first.get('store') or ''} {first.get('dept') or ''}".strip(),
                mileage=first.get("mileage_km"),
                net=sum(float(l.get("total") or 0) for l in lines),
                remark=first.get("remark") or "",
                lines=[
                    {
                        "ลำดับ": l.get("line_no"),
                        "รหัสพัสดุ": l.get("inv_code") or "",
                        "รายการพัสดุ": l.get("inv_name") or "",
                        "จำนวน": l.get("qty"),
                        "ราคา/หน่วย": l.get("unit_price"),
                        "จำนวนเงิน (บาท)": l.get("total"),
                        "หมายเหตุ": l.get("remark") or "",
                    }
                    for l in lines
                ],
            )
    return None


def _render_doc_detail_panel(app_dir, db, plate, doc_no, doc_kind=None, doc_date=None, full=None):
    """แสดงการ์ดรายละเอียดใบ W/O หรือ WX แบบเต็ม + ใบแจ้งซ่อมที่เกี่ยวข้องในฐานข้อมูล"""
    full = full or {}
    with st.container(border=True):
        st.markdown(f"### 🔍 รายละเอียดเอกสาร `{doc_no}` · ทะเบียน `{plate}`")
        detail = _lookup_doc_detail(app_dir, plate, doc_no, doc_kind)
        if detail:
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("ประเภทเอกสาร", detail["kind"])
            c2.metric("วันที่", detail["date"] or (str(doc_date) if doc_date else "-"))
            c3.metric("เลขไมล์", f"{int(detail['mileage']):,} กม." if detail.get("mileage") else "-")
            c4.metric("ยอดรวมสุทธิ", f"{float(detail['net'] or 0):,.2f} บาท")
            if detail.get("vendor"):
                st.caption(f"**ศูนย์บริการ / คลัง:** {detail['vendor']}  ·  **ไฟล์ต้นทาง:** `{detail['file']}`")
            if detail.get("remark"):
                st.caption(f"**หมายเหตุ:** {detail['remark']}")
            ldf = pd.DataFrame(detail["lines"])
            if len(ldf):
                st.markdown(f"**รายการซ่อม / เบิกพัสดุทั้งหมดในเอกสาร ({len(ldf)} บรรทัด)**")
                st.dataframe(
                    ldf,
                    hide_index=True,
                    column_config={
                        "ราคา/หน่วย": st.column_config.NumberColumn(format="%,.2f"),
                        "จำนวนเงิน (บาท)": st.column_config.NumberColumn(format="%,.2f"),
                    },
                    **full,
                )
        else:
            st.info(
                f"ไม่พบไฟล์รายงานรายคัน (`Vehicle_Reports/{plate}*.xls`) บนเครื่องนี้สำหรับแสดงบรรทัดรายการ "
                "(โหมด Cloud จะไม่มีไฟล์ .xls ต้นทางในเครื่อง)"
            )

        # แสดงใบแจ้งซ่อม (repair_request) ที่เกี่ยวข้องในช่วง ±15 วัน
        if db and os.path.exists(db) and doc_date:
            try:
                d0 = dt.date.fromisoformat(str(doc_date)[:10])
                d_lo = (d0 - dt.timedelta(days=15)).isoformat()
                d_hi = (d0 + dt.timedelta(days=15)).isoformat()
                rel = rr.fetch(
                    db,
                    "SELECT doc_no เลขที่ใบแจ้งซ่อม, report_date วันที่แจ้งซ่อม, open_date วันที่เปิดเอกสาร, "
                    "status สถานะ, employee_code รหัสพนักงาน FROM repair_request "
                    "WHERE asset_code=? AND report_date BETWEEN ? AND ? ORDER BY report_date",
                    plate, d_lo, d_hi,
                )
                if rel:
                    st.markdown(f"**ใบแจ้งซ่อมที่เกี่ยวข้องของรถ `{plate}` (ช่วง ±15 วัน)**")
                    st.dataframe(pd.DataFrame(rel), hide_index=True, **full)
            except Exception:
                pass


def _render_draft_plan_tab(app_dir, db, full, kpi_card):
    """แท็บแสดงร่าง Plan PM หัวลาก 24 เดือน (ม.ค. 2026 – ธ.ค. 2027) สำหรับตรวจและ Export ก่อนนำไปแทน Plan เดิมใน KPI"""
    st.info(
        "📌 **ข้อมูลร่างแยกไว้สำหรับตรวจสอบ (รอยืนยันก่อนนำไปแทน Plan เดิมในหน้า KPI)** — "
        "คำนวณจากประวัติซ่อมจริง (W/O + WX) สำหรับหัวลาก Master (ม.ค. 2026 – ธ.ค. 2027)"
    )
    default_path = os.path.join(app_dir, "data", "Plan_PM_หัวลาก_ใหม่_รอยืนยัน.xlsx")
    raw_bytes = None
    src_label = None
    if os.path.exists(default_path):
        with open(default_path, "rb") as fh:
            raw_bytes = fh.read()
        src_label = os.path.basename(default_path)
    else:
        up = st.file_uploader("อัปโหลดไฟล์ Plan_PM_หัวลาก_ใหม่_รอยืนยัน.xlsx (โหมด Cloud)", type=["xlsx"], key="pm_draft_up")
        if up is not None:
            raw_bytes = up.getvalue()
            src_label = up.name

    if not raw_bytes:
        st.warning("ไม่พบไฟล์ร่าง `data/Plan_PM_หัวลาก_ใหม่_รอยืนยัน.xlsx` — กรุณาอัปโหลดไฟล์ด้านบนเพื่อแสดงผล")
        return

    data = _load_draft_pm_workbook(raw_bytes)
    st_map = {r["สถานะ"]: r["จำนวน (คัน)"] for r in data["status"]}
    n_master = len(data["detail"]) if len(data["detail"]) else sum(st_map.values())
    n_norm = st_map.get("ตามรอบปกติ", 0)
    n_due_ovd = st_map.get("ครบกำหนด พ.ย. 2026", 0) + st_map.get("เกินกำหนด (ค้าง)", 0)
    n_nohist = st_map.get("ไม่มีประวัติ PM (2026)", 0)

    k = st.columns(4)
    kpi_card(k[0], "หัวลาก Master", f"{n_master:,}", "คัน (ณ 22/11/2026)")
    kpi_card(k[1], "ตามรอบปกติ", f"{n_norm:,}", "คัน", "ok")
    kpi_card(k[2], "ครบกำหนด พ.ย. + ค้าง", f"{n_due_ovd:,}", "คัน (ยกยอด ธ.ค. 2026)", "warn" if n_due_ovd else "ok")
    kpi_card(k[3], "ไม่มีประวัติ PM 2026", f"{n_nohist:,}", "คัน (รอคุณโอ๋กำหนดรอบ)", "warn" if n_nohist else "ok")
    st.write("")

    c_dl1, c_dl2 = st.columns([2.2, 2.8])
    c_dl1.download_button(
        "⬇️ Export ร่าง Plan PM 2026–2027 ฉบับเต็ม (Excel 4 ชีต)",
        raw_bytes,
        file_name=f"Plan_PM_หัวลาก_2026_2027_{dt.datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key="pm_draft_dl_full",
        **full,
    )
    c_dl2.caption(f"แหล่งข้อมูล: `{src_label}` · 4 ชีตพร้อมสีและสูตร (`สรุป_Plan_PM_2026_2027`, `ตาราง_PM_217คัน_2026_2027`, `รายละเอียด_217คัน`, `ประวัติ_PM_จริง_2026`)")

    view = st.radio(
        "เลือกมุมมอง",
        ["📊 สรุปรายเดือน 2026–2027", "🗓️ ตารางรายคัน (24 เดือน)", "📋 รายละเอียดรายคัน", "🧾 ประวัติ PM จริง 2026"],
        horizontal=True,
        key="pm_draft_view",
    )

    if view == "📊 สรุปรายเดือน 2026–2027":
        df26 = pd.DataFrame(data["plan_2026"])
        df27 = pd.DataFrame(data["plan_2027"])
        st.markdown("**1) ปี 2026 (ม.ค.–ธ.ค. 2026): Plan PM ใหม่ vs ทำ PM จริง vs ใบแจ้งซ่อมทั่วไป**")
        if len(df26):
            st.dataframe(
                df26,
                hide_index=True,
                column_config={"% ตามแผน PM": st.column_config.NumberColumn(format="%.1f%%")},
                **full,
            )
            chart_df = df26[df26["เดือน"] != "2026-12"].set_index("เดือน")[["รวม Plan PM ใหม่", "ทำ PM จริง"]].fillna(0)
            st.bar_chart(chart_df, height=280, stack=False)
        st.markdown("**2) ปี 2027 (ม.ค.–ธ.ค. 2027): แผน PM หัวลากล่วงหน้า (เฉลี่ย 39.4 คัน/เดือน)**")
        if len(df27):
            st.dataframe(df27, hide_index=True, **full)

    elif view == "🗓️ ตารางรายคัน (24 เดือน)":
        mdf = data["matrix"]
        if len(mdf):
            f1, f2, f3 = st.columns([1.4, 1.4, 1.6])
            q_p = f1.text_input("ค้นหาทะเบียน", key="pm_d_q1", placeholder="เช่น 60-0000")
            pt_opts = ["ทั้งหมด"] + sorted(mdf["ประเภท PM ล่าสุด"].dropna().unique().tolist())
            f_pt = f2.selectbox("ประเภท PM ล่าสุด", pt_opts, key="pm_d_pt")
            m_cols = [c for c in mdf.columns if c.startswith(("2026-", "2027-"))]
            f_mo = f3.selectbox("กรองเฉพาะรถที่มีแผน/ทำจริงในเดือน", ["ทุกเดือน"] + m_cols, key="pm_d_mo")
            sub = mdf.copy()
            if q_p:
                sub = sub[sub["ทะเบียน"].str.contains(q_p.strip(), case=False, regex=False)]
            if f_pt != "ทั้งหมด":
                sub = sub[sub["ประเภท PM ล่าสุด"] == f_pt]
            if f_mo != "ทุกเดือน":
                sub = sub[sub[f_mo].astype(str).str.strip() != ""]
            st.caption(f"แสดง {len(sub):,} คัน")
            st.dataframe(sub, hide_index=True, height=500, **full)
            st.download_button(
                "⬇️ Export ตารางที่กรอง (Excel)",
                _df_to_excel_bytes({"ตาราง_PM_ที่กรอง": sub}),
                file_name=f"ตาราง_PM_ที่กรอง_{dt.datetime.now():%Y%m%d_%H%M}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="pm_d_dl_mat",
            )

    elif view == "📋 รายละเอียดรายคัน":
        ddf = data["detail"]
        if len(ddf):
            f1, f2 = st.columns([1.5, 1.5])
            q_d = f1.text_input("ค้นหาทะเบียน", key="pm_d_q2")
            st_col = next((c for c in ddf.columns if c.startswith("สถานะ")), None)
            sub = ddf.copy()
            if st_col:
                st_opts = ["ทุกสถานะ"] + sorted(sub[st_col].dropna().astype(str).unique().tolist())
                f_st = f2.selectbox("กรองสถานะ", st_opts, key="pm_d_st")
                if f_st != "ทุกสถานะ":
                    sub = sub[sub[st_col].astype(str) == f_st]
            if q_d:
                sub = sub[sub["ทะเบียน"].astype(str).str.contains(q_d.strip(), case=False, regex=False)]
            st.caption(f"แสดง {len(sub):,} คัน")
            st.dataframe(sub, hide_index=True, height=500, **full)
            st.download_button(
                "⬇️ Export รายละเอียดที่กรอง (Excel)",
                _df_to_excel_bytes({"รายละเอียด_PM": sub}),
                file_name=f"รายละเอียด_PM_{dt.datetime.now():%Y%m%d_%H%M}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="pm_d_dl_det",
            )

    else:
        hdf = data["hist"]
        if len(hdf):
            f1, f2 = st.columns([1.4, 1.6])
            mo_opts = ["ทุกเดือน"] + sorted(hdf["เดือน"].dropna().astype(str).unique().tolist())
            f_hm = f1.selectbox("เดือน (2026)", mo_opts, key="pm_d_hm")
            q_h = f2.text_input("ค้นหาทะเบียน / เลขที่เอกสาร / รายการ", key="pm_d_qh")
            sub = hdf.copy()
            if f_hm != "ทุกเดือน":
                sub = sub[sub["เดือน"].astype(str) == f_hm]
            if q_h:
                mask = (
                    sub["ทะเบียน"].astype(str).str.contains(q_h.strip(), case=False, regex=False)
                    | sub["เลขที่เอกสาร"].astype(str).str.contains(q_h.strip(), case=False, regex=False)
                    | sub["รายการน้ำมันเครื่อง"].astype(str).str.contains(q_h.strip(), case=False, regex=False)
                )
                sub = sub[mask]
            st.caption(f"แสดง {len(sub):,} รายการ — 💡 **คลิกเลือกแถวในตาราง** (หรือเลือกเลขที่เอกสารด้านล่าง) เพื่อเปิดดูรายละเอียดรายการซ่อม/เบิกพัสดุทั้งหมด")
            sub_reset = sub.reset_index(drop=True)
            ev_h = st.dataframe(
                sub_reset,
                hide_index=True,
                height=420,
                on_select="rerun",
                selection_mode="single-row",
                key="pm_d_hist_tbl",
                **full,
            )
            sel_rows = ev_h.selection.rows if hasattr(ev_h, "selection") and ev_h.selection else []
            doc_opts = [f"{r['เลขที่เอกสาร']} | {r['ทะเบียน']} ({r['วันที่']})" for _, r in sub_reset.iterrows()]
            picked_idx = sel_rows[0] if sel_rows else None
            c_pk, c_ex = st.columns([2.8, 1.4])
            if doc_opts:
                sel_lbl = c_pk.selectbox(
                    "หรือเลือกเลขที่เอกสารเพื่อดูรายละเอียด",
                    ["— คลิกแถวในตาราง หรือเลือกเลขที่เอกสาร —"] + doc_opts,
                    index=(picked_idx + 1) if picked_idx is not None and picked_idx < len(doc_opts) else 0,
                    key="pm_d_hist_pick",
                )
                if sel_lbl != "— คลิกแถวในตาราง หรือเลือกเลขที่เอกสาร —":
                    picked_idx = doc_opts.index(sel_lbl)
            c_ex.download_button(
                "⬇️ Export ประวัติ PM ที่กรอง (Excel)",
                _df_to_excel_bytes({"ประวัติ_PM_จริง_2026": sub_reset}),
                file_name=f"ประวัติ_PM_จริง_2026_{dt.datetime.now():%Y%m%d_%H%M}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="pm_d_dl_hist",
                **full,
            )
            if picked_idx is not None and 0 <= picked_idx < len(sub_reset):
                rec = sub_reset.iloc[picked_idx]
                _render_doc_detail_panel(
                    app_dir, db,
                    plate=rec["ทะเบียน"],
                    doc_no=rec["เลขที่เอกสาร"],
                    doc_kind=rec.get("ชนิด"),
                    doc_date=rec.get("วันที่"),
                    full=full,
                )


def render_pm(is_local, user, full, app_dir, kpi_card):
    ss = st.session_state
    db = db_selector(is_local, app_dir, full)
    st.markdown('<div class="app-hero"><h1>PM Plan รายเดือน</h1><p>หัวลาก · เข้าศูนย์บริการทุก 20,000 กม. · เปลี่ยนถ่ายที่ลานทุก 30,000 กม. · '
                'หรือครบ 6 เดือน (อย่างใดถึงก่อน) · กม./วันจากประวัติเลขไมล์และเลขไมล์ตอนเข้าบริการ</p></div>', unsafe_allow_html=True)

    o = ss.setdefault("pm_opt", dict(pm.DEFAULTS))
    nsvc = pm.fetch(db, "SELECT COUNT(*) n, SUM(pm_type IS NOT NULL) p FROM service_doc")[0]
    res = pm.compute_plan(db, **o)
    rows = res["rows"]

    t_draft, t_plan, t_all, t_imp, t_hist = st.tabs([
        "📅 ร่าง Plan PM 2026–2027 (รอยืนยัน)",
        "📅 แผนรายเดือน (จากฐานข้อมูล)",
        "🚛 สถานะรายคัน",
        "📥 นำเข้าประวัติบริการ",
        "🔎 ประวัติรายคัน",
    ])

    with t_draft:
        _render_draft_plan_tab(app_dir, db, full, kpi_card)

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
            import monthly_report as mr
            blob_ = _excel(pdf, pd.DataFrame(rows), crit, title)
            fn_ = f"PM_Plan_{title}_{dt.datetime.now():%Y%m%d_%H%M}.xlsx"
            sp_ = mr.save_local_export(app_dir, fn_, blob_, is_local=is_local)
            ss["pm_blob"] = (fn_, blob_, sp_)
        if ss.get("pm_blob"):
            _pb = ss["pm_blob"]
            if len(_pb) > 2 and _pb[2]:
                st.success(f"💾 บันทึกไฟล์แล้วที่: `{_pb[2]}`")
            st.download_button(f"ดาวน์โหลด {_pb[0]}", _pb[1], file_name=_pb[0], icon="⬇️", type="primary",
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
            st.markdown("**ประวัติ PM** — 💡 คลิกเลือกแถวเพื่อดูรายละเอียดรายการซ่อม/เบิกพัสดุทั้งหมดในเอกสาร")
            if len(evs):
                ev_p = st.dataframe(evs, hide_index=True, on_select="rerun", selection_mode="single-row", key="pm_h_ev_tbl", **full)
                sel_e = ev_p.selection.rows if hasattr(ev_p, "selection") and ev_p.selection else []
                if sel_e:
                    er = evs.iloc[sel_e[0]]
                    _render_doc_detail_panel(app_dir, db, plate=p_, doc_no=er["เลขที่"], doc_kind=er["ประเภท"], doc_date=er["วันที่"], full=full)
            else:
                st.info("ไม่พบ PM ในข้อมูลที่นำเข้า")
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
