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


def _build_draft_pm_from_db(db_path):
    """คำนวณร่าง Plan PM หัวลาก 24 เดือน (ม.ค. 2026 – ธ.ค. 2027) สดจากฐานข้อมูล SQLite โดยตรง
    (vr_wo_header / vr_wo_line / vr_issue_line -> service_doc + asset_master + repair_request + mileage_reading)"""
    if not db_path or not os.path.exists(db_path):
        return None

    res = pm.compute_plan(db_path)
    plan_by_plate = {r["asset_code"]: r for r in res["rows"]}
    fleet_rate = res["fleet_rate"] or 177.0
    today = dt.date.fromisoformat(res["today"])
    cur_ym = f"{today.year:04d}-{today.month:02d}"

    con = pm.connect(db_path)
    try:
        has_am = con.execute("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='asset_master'").fetchone()[0] > 0
        if has_am:
            masters = [
                r[0] for r in con.execute(
                    "SELECT asset_code FROM asset_master WHERE asset_type='หัวลาก' AND is_active=1 ORDER BY sort_order, asset_code"
                )
            ]
        else:
            masters = []
        if not masters:
            masters = sorted(plan_by_plate.keys())
        master_set = set(masters)

        # ใบแจ้งซ่อมทั่วไป (Master) รายเดือน ปี 2026
        rr_by_month = {
            r[0]: r[1]
            for r in con.execute(
                "SELECT report_month, COUNT(DISTINCT asset_code) FROM repair_request "
                "WHERE asset_type='หัวลาก' AND in_iso='Yes' AND report_month LIKE '2026-%' GROUP BY 1"
            )
        }

        # ประวัติ PM จริง ปี 2026 (เฉพาะหัวลากใน Master)
        pm_ev_rows = con.execute(
            "SELECT substr(doc_date,1,7) AS ym, asset_code, doc_date, doc_no, doc_kind, pm_type, km, pm_items "
            "FROM service_doc WHERE pm_type IS NOT NULL AND doc_date LIKE '2026-%' "
            "ORDER BY doc_date, asset_code, doc_no"
        ).fetchall()
    finally:
        con.close()

    ev_by_plate = {}
    hist_list = []
    act_plates_by_ym = {}
    for ym, code, dd, dno, dkind, ptype, km, pitems in pm_ev_rows:
        if code not in master_set:
            continue
        ev_by_plate.setdefault(code, []).append(dict(ym=ym, doc_date=dd, doc_no=dno, doc_kind=dkind, pm_type=ptype, km=km, pm_items=pitems))
        act_plates_by_ym.setdefault(ym, set()).add(code)
        try:
            d_fmt = dt.date.fromisoformat(str(dd)[:10]).strftime("%d/%m/%Y")
        except Exception:
            d_fmt = str(dd or "")
        hist_list.append({
            "เดือน": ym,
            "ทะเบียน": code,
            "วันที่": d_fmt,
            "เลขที่เอกสาร": dno,
            "ชนิด": dkind,
            "ประเภท PM": ptype,
            "เลขไมล์": km,
            "รายการน้ำมันเครื่อง": pitems or "",
        })

    months_24 = [f"2026-{m:02d}" for m in range(1, 13)] + [f"2027-{m:02d}" for m in range(1, 13)]
    plan_cnt = {ym: {"ศูนย์บริการ": 0, "ลาน": 0, "ovd": 0} for ym in months_24}

    status_counts = {"ครบกำหนด พ.ย. 2026": 0, "ตามรอบปกติ": 0, "เกินกำหนด (ค้าง)": 0, "ไม่มีประวัติ PM (2026)": 0}
    detail_rows = []
    matrix_rows = []

    for code in masters:
        pr = plan_by_plate.get(code, {})
        evs = ev_by_plate.get(code, [])
        rate = pr.get("km_day") or fleet_rate
        m_cells = {ym: "" for ym in months_24}

        if not evs:
            st_lbl = "ไม่มีประวัติ PM (2026)"
            status_counts[st_lbl] += 1
            detail_rows.append({
                "ทะเบียน": code,
                "วันที่ PM ล่าสุด": "",
                "ประเภท PM": "ไม่มีประวัติ PM (2026)",
                "จำนวนครั้ง PM ปี 2026": 0,
                "เลขไมล์ตอน PM": None,
                "เลขไมล์ล่าสุด": pr.get("last_km"),
                "กม./วัน": round(rate, 1) if rate else None,
                "รอบ (กม.)": None,
                "รอบ (วัน)": None,
                "ครบกำหนดถัดไป": "",
                "สถานะ ณ ปัจจุบัน": st_lbl,
                "หมายเหตุ": "ไม่พบรายการเปลี่ยนน้ำมันเครื่องในรายงาน 2026 — รอคุณแขกกำหนดรอบเริ่มต้น",
            })
            matrix_rows.append([code, "ไม่มีประวัติ PM (2026)"] + [m_cells[ym] for ym in months_24])
            continue

        last_ev = evs[-1]
        ptype = last_ev["pm_type"]
        is_center_like = ptype in ("ศูนย์บริการ", "เปลี่ยนเครื่องยนต์")
        tag = "ศูนย์" if is_center_like else "ลาน"
        interval_km = 20000 if is_center_like else 30000
        cycle_days = max(30, min(182, int(round(interval_km / rate)))) if rate else 180

        # 1) ทำ PM / เปลี่ยนเครื่องยนต์ จริงในปี 2026
        for e in evs:
            if e["pm_type"] == "เปลี่ยนเครื่องยนต์":
                m_cells[e["ym"]] = "เปลี่ยนเครื่องยนต์"
                plan_cnt[e["ym"]]["ศูนย์บริการ"] += 1
            else:
                e_tag = "ศูนย์" if e["pm_type"] == "ศูนย์บริการ" else "ลาน"
                m_cells[e["ym"]] = f"ทำจริง-{e_tag}"
                plan_cnt[e["ym"]][e["pm_type"]] += 1

        # 2) แผนย้อนหลังก่อนทำจริงครั้งแรกในปี 2026 (เฉพาะกรณีที่ครั้งแรกไม่ใช่เปลี่ยนเครื่องยนต์)
        first_ev = evs[0]
        if first_ev["pm_type"] != "เปลี่ยนเครื่องยนต์":
            f_type = first_ev["pm_type"]
            f_tag = "ศูนย์" if f_type == "ศูนย์บริการ" else "ลาน"
            f_int = 20000 if f_type == "ศูนย์บริการ" else 30000
            f_cyc = max(30, min(182, int(round(f_int / rate)))) if rate else 180
            d_back = dt.date.fromisoformat(first_ev["doc_date"]) - dt.timedelta(days=f_cyc)
            while d_back >= dt.date(2026, 1, 1):
                ym_b = f"{d_back.year:04d}-{d_back.month:02d}"
                if ym_b in m_cells and not m_cells[ym_b]:
                    m_cells[ym_b] = f"แผน-{f_tag}"
                    plan_cnt[ym_b][f_type] += 1
                d_back -= dt.timedelta(days=f_cyc)

        # 3) กำหนดครบถัดไป & สถานะ (อิงจากイベントล่าสุดตั้งแต่เปลี่ยนเครื่องยนต์เป็นต้นมา)
        plan_type_next = "ศูนย์บริการ" if is_center_like else "ลาน"
        due_iso = pr.get("due_date")
        due_d = dt.date.fromisoformat(due_iso) if due_iso else (dt.date.fromisoformat(last_ev["doc_date"]) + dt.timedelta(days=cycle_days))
        due_ym = f"{due_d.year:04d}-{due_d.month:02d}"

        if due_ym < "2026-11":
            st_lbl = "เกินกำหนด (ค้าง)"
            status_counts[st_lbl] += 1
            if due_ym in m_cells and not m_cells[due_ym]:
                m_cells[due_ym] = f"ครบกำหนด(ค้าง)-{tag}"
                plan_cnt[due_ym][plan_type_next] += 1
            if not m_cells["2026-12"]:
                m_cells["2026-12"] = f"ค้างยกยอด-{tag}"
            plan_cnt["2026-12"]["ovd"] += 1
            d_next = dt.date(2026, 12, 15) + dt.timedelta(days=cycle_days)
        elif due_ym == "2026-11":
            st_lbl = "ครบกำหนด พ.ย. 2026"
            status_counts[st_lbl] += 1
            if not m_cells["2026-11"]:
                m_cells["2026-11"] = f"แผน-{tag}"
                plan_cnt["2026-11"][plan_type_next] += 1
            if not m_cells["2026-12"]:
                m_cells["2026-12"] = f"ค้างยกยอด-{tag}"
            plan_cnt["2026-12"]["ovd"] += 1
            d_next = dt.date(2026, 12, 15) + dt.timedelta(days=cycle_days)
        else:
            st_lbl = "ตามรอบปกติ"
            status_counts[st_lbl] += 1
            d_next = due_d

        # 4) แผนล่วงหน้า 2026-12 .. 2027-12
        while d_next <= dt.date(2027, 12, 31):
            ym_f = f"{d_next.year:04d}-{d_next.month:02d}"
            if ym_f in m_cells and not m_cells[ym_f]:
                m_cells[ym_f] = f"แผน-{tag}"
                plan_cnt[ym_f][plan_type_next] += 1
            d_next += dt.timedelta(days=cycle_days)

        d_last_fmt = dt.date.fromisoformat(last_ev["doc_date"]).strftime("%d/%m/%Y")
        d_due_fmt = due_d.strftime("%d/%m/%Y")
        detail_rows.append({
            "ทะเบียน": code,
            "วันที่ PM ล่าสุด": d_last_fmt,
            "ประเภท PM": ptype,
            "จำนวนครั้ง PM ปี 2026": len(evs),
            "เลขไมล์ตอน PM": last_ev.get("km"),
            "เลขไมล์ล่าสุด": pr.get("last_km"),
            "กม./วัน": round(rate, 1) if rate else None,
            "รอบ (กม.)": interval_km,
            "รอบ (วัน)": cycle_days,
            "ครบกำหนดถัดไป": d_due_fmt,
            "สถานะ ณ ปัจจุบัน": st_lbl,
            "หมายเหตุ": f"{pr.get('rate_src') or 'ประวัติเลขไมล์'} · {pr.get('due_by') or ''}",
        })
        matrix_rows.append([code, ptype] + [m_cells[ym] for ym in months_24])

    status_list = [
        {"สถานะ": "ครบกำหนด พ.ย. 2026", "จำนวน (คัน)": status_counts["ครบกำหนด พ.ย. 2026"], "หมายเหตุ": "รวมยกยอดเข้า ธ.ค. 2026"},
        {"สถานะ": "ตามรอบปกติ", "จำนวน (คัน)": status_counts["ตามรอบปกติ"], "หมายเหตุ": "คำนวณรอบตามระยะ/6 เดือน จาก SQLite"},
        {"สถานะ": "เกินกำหนด (ค้าง)", "จำนวน (คัน)": status_counts["เกินกำหนด (ค้าง)"], "หมายเหตุ": "รวมยกยอดเข้า ธ.ค. 2026"},
        {"สถานะ": "ไม่มีประวัติ PM (2026)", "จำนวน (คัน)": status_counts["ไม่มีประวัติ PM (2026)"], "หมายเหตุ": "แยกไว้รอคุณแขกกำหนดรอบเริ่มต้น"},
    ]

    plan_2026 = []
    for m in range(1, 13):
        ym = f"2026-{m:02d}"
        ctr = plan_cnt[ym]["ศูนย์บริการ"]
        yrd = plan_cnt[ym]["ลาน"]
        ovd = plan_cnt[ym]["ovd"]
        tot = ctr + yrd + ovd
        act = len(act_plates_by_ym.get(ym, set())) if ym <= cur_ym else None
        pct = round(act * 100.0 / tot, 1) if (act is not None and tot > 0) else None
        rep = rr_by_month.get(ym)
        old_p = _old_plan(m)
        plan_2026.append({
            "เดือน": ym, "PM ศูนย์ (20k)": ctr, "PM ลาน (30k)": yrd, "ค้างยกยอด": ovd,
            "รวม Plan PM ใหม่": tot, "ทำ PM จริง": act, "% ตามแผน PM": pct,
            "ใบแจ้งซ่อมทั่วไป (Master)": rep, "Plan เดิมใน KPI": old_p,
        })

    plan_2027 = []
    for m in range(1, 13):
        ym = f"2027-{m:02d}"
        ctr = plan_cnt[ym]["ศูนย์บริการ"]
        yrd = plan_cnt[ym]["ลาน"]
        tot = ctr + yrd
        old_p = _old_plan(m)
        plan_2027.append({
            "เดือน": ym, "PM ศูนย์ (20k)": ctr, "PM ลาน (30k)": yrd,
            "รวม Plan PM ใหม่": tot, "Plan เดิมใน KPI": old_p, "ส่วนต่าง (ใหม่-เดิม)": tot - old_p,
        })

    mat_hdrs = ["ทะเบียน", "ประเภท PM ล่าสุด"] + months_24
    return dict(
        status=status_list,
        plan_2026=plan_2026,
        plan_2027=plan_2027,
        matrix=pd.DataFrame(matrix_rows, columns=mat_hdrs),
        detail=pd.DataFrame(detail_rows),
        hist=pd.DataFrame(hist_list),
        fleet_rate=fleet_rate,
        today_str=today.strftime("%d/%m/%Y"),
    )


def _classify_pm4_line(item_name, item_code="", doc_kind=""):
    """จำแนกรายการซ่อม/เบิกพัสดุจริง 1 บรรทัด เข้า 4 คอลัมน์ PM ประจำเดือน:
      c1 = เช็คระยะเปลี่ยนถ่ายน้ำมันเครื่อง (30,000 km / 6 เดือน)
      c2 = เช็คระยะเปลี่ยนถ่ายน้ำมันเฟืองท้าย/เกียร์ (60,000 km / 6 เดือน)
      c3 = เช็คช่วงล่าง (ระบบเบรก / ระบบคลัทช์ / จาระบีลูกปืนล้อ) (120,000 km)
      c4 = ตรวจเช็คลมยาง / ดอกยางคงเหลือ
    """
    s = str(item_name or "").strip()
    su = s.upper()
    c1 = c2 = c3 = c4 = False

    # Col 1: เปลี่ยนถ่ายน้ำมันเครื่อง / เปลี่ยนเครื่องยนต์
    if doc_kind == "W/O":
        if pm.is_center_item(s, item_code) or pm.is_engine_replace_item(s, item_code):
            c1 = True
    else:
        if pm.is_yard_item(s):
            c1 = True

    # Col 2: เปลี่ยนถ่ายน้ำมันเฟืองท้าย / น้ำมันเกียร์ / กรองเกียร์ / กรองเฟืองท้าย
    if any(k in s for k in ("น้ำมันเกียร์", "น้ำมันเฟืองท้าย", "กรองเกียร์", "กรองเฟืองท้าย")) or any(
        k in su for k in ("GEAR OIL", "DIFF OIL", "AXLE OIL", "TRANSMISSION OIL")
    ):
        c2 = True

    # Col 3: ช่วงล่าง (ระบบเบรก / ระบบคลัทช์ / จาระบีลูกปืนล้อ)
    if any(
        k in s
        for k in (
            "ผ้าเบรก", "ผ้าเบรค", "จานเบรก", "จานเบรค", "น้ำมันเบรก", "น้ำมันเบรค", "หม้อลมเบรก", "หม้อลมเบรค",
            "คลัทช์", "คลัช", "ครัช",
            "จาระบี", "จารบี", "ลูกปืนล้อ", "ช่วงล่าง", "ตั้งเบรก", "ตั้งเบรค", "ตรวจเบรก", "ตรวจเบรค",
        )
    ) or any(k in su for k in ("BRAKE", "CLUTCH", "GREASE", "WHEEL BEARING")):
        c3 = True

    # Col 4: ตรวจเช็คลมยาง / ดอกยาง / สลับยาง / ตั้งศูนย์ / ถ่วงล้อ / ปะยาง / เปลี่ยนยาง
    if any(
        k in s
        for k in ("ลมยาง", "ดอกยาง", "สลับยาง", "ปะยาง", "เปลี่ยนยาง", "ตั้งศูนย์", "ถ่วงล้อ", "ยางนอก")
    ) or any(k in su for k in ("TIRE", "TYRE")):
        c4 = True

    return c1, c2, c3, c4


def _build_monthly_pm_template_rows(db_path, ym):
    """สร้างข้อมูลตาราง 'แผนซ่อมบำรุงรักษา - หัวลาก' ตาม Template ประจำเดือน ym (YYYY-MM) จากรายการซ่อมจริงใน SQLite"""
    import calendar
    y, m = int(ym[:4]), int(ym[5:7])
    last_day = calendar.monthrange(y, m)[1]
    date_range_txt = f"ตั้งแต่วันที่ 01/{m:02d}/{y:04d} - {last_day:02d}/{m:02d}/{y:04d}"

    COL1 = "เช็คระยะเปลี่ยน\nถ่ายน้ำมันเครื่อง\n(ระบบไฟ/ระบบลม)\nเช็คทั่วไป 30,000\nkm หรือ 6 เดือน"
    COL2 = "เช็คระยะเปลี่ยนถ่าย\nน้ำมันเฟืองท้าย/เกียร์\n(ระบบเกียร์) 60,000\nkm หรือ 6 เดือน"
    COL3 = "เช็คช่วงล่าง\n(ระบบเบรก / ระบบ\nคลัทช์/\nจาระบีลูกปืนล้อ)\n120,000 km"
    COL4 = "ตรวจเช็คลมยาง /\nดอกยางคงเหลือ"

    con = pm.connect(db_path)
    try:
        has_am = con.execute("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='asset_master'").fetchone()[0] > 0
        if has_am:
            masters = [
                r[0] for r in con.execute(
                    "SELECT asset_code FROM asset_master WHERE asset_type='หัวลาก' AND is_active=1 ORDER BY sort_order, asset_code"
                )
            ]
        else:
            masters = []
        master_set = set(masters)

        # ดึงยี่ห้อ/รุ่นรถจาก vr_wo_header / vr_issue_line
        brand_map = {}
        try:
            for plate, br in con.execute(
                "SELECT plate, MAX(brand) FROM ("
                "  SELECT plate, brand FROM vr_wo_header WHERE IFNULL(brand,'') != '' "
                "  UNION ALL "
                "  SELECT plate, brand FROM vr_issue_line WHERE IFNULL(brand,'') != '' "
                ") GROUP BY plate"
            ):
                brand_map[plate] = br or ""
        except Exception:
            pass

        # ตรวจรายการซ่อมจริงในเดือน ym จากทั้ง vr_wo_line (ศูนย์) และ vr_issue_line (ลาน)
        flags = {p: [False, False, False, False] for p in masters}
        refs_by_plate = {p: [] for p in masters}
        seen_ref = set()

        def _add_ref(plate, label):
            key = (plate, label)
            if key not in seen_ref:
                seen_ref.add(key)
                refs_by_plate[plate].append(label)

        try:
            for plate, dno, icode, iname in con.execute(
                "SELECT plate, doc_no, item_code, item_name FROM vr_wo_line "
                "WHERE substr(doc_date,1,7) = ? ORDER BY doc_date, doc_no, line_no",
                (ym,),
            ):
                if plate not in master_set:
                    continue
                c1, c2, c3, c4 = _classify_pm4_line(iname, icode, "W/O")
                if c1 or c2 or c3 or c4:
                    if c1: flags[plate][0] = True
                    if c2: flags[plate][1] = True
                    if c3: flags[plate][2] = True
                    if c4: flags[plate][3] = True
                    _add_ref(plate, f"ศูนย์ W/O {dno}")
        except Exception:
            pass

        try:
            for plate, dno, iname in con.execute(
                "SELECT plate, doc_no, inv_name FROM vr_issue_line "
                "WHERE substr(doc_date,1,7) = ? ORDER BY doc_date, doc_no, line_no",
                (ym,),
            ):
                if plate not in master_set:
                    continue
                c1, c2, c3, c4 = _classify_pm4_line(iname, "", "WX")
                if c1 or c2 or c3 or c4:
                    if c1: flags[plate][0] = True
                    if c2: flags[plate][1] = True
                    if c3: flags[plate][2] = True
                    if c4: flags[plate][3] = True
                    _add_ref(plate, f"ลาน WX {dno}")
        except Exception:
            pass
    finally:
        con.close()

    rows = []
    n_checked = 0
    for plate in masters:
        f1, f2, f3, f4 = flags[plate]
        if f1 or f2 or f3 or f4:
            n_checked += 1
        rows.append({
            "ทะเบียน": plate,
            "รุ่นรถ": brand_map.get(plate, ""),
            COL1: "Yes" if f1 else "",
            COL2: "Yes" if f2 else "",
            COL3: "Yes" if f3 else "",
            COL4: "Yes" if f4 else "",
            "อ้างอิงเอกสาร": " , ".join(refs_by_plate[plate]),
        })
    return rows, date_range_txt, n_checked


def _write_monthly_pm_template_sheet(wb, rows, date_range_txt, sheet_name="สรุป_PM_ประจำเดือน", logo_path=None):
    """เขียนชีต 'สรุป_PM_ประจำเดือน' ตาม Template แบบฟอร์ม แผนซ่อมบำรุงรักษา - หัวลาก"""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    ws = wb.create_sheet(sheet_name[:31])
    f_co_th = Font(name="Tahoma", size=12, bold=True, color="1F4E78")
    f_co_en = Font(name="Tahoma", size=10, bold=True, color="1F4E78")
    f_co_sub = Font(name="Tahoma", size=9, color="374151")
    f_doc = Font(name="Tahoma", size=12, bold=True, color="000000")
    f_date = Font(name="Tahoma", size=10, bold=True, color="1F4E78")
    f_hdr = Font(name="Tahoma", size=9.5, bold=True, color="FFFFFF")
    f_cell = Font(name="Tahoma", size=9.5, color="1F2937")
    f_yes = Font(name="Tahoma", size=9.5, bold=True, color="0F5132")
    fill_hdr = PatternFill("solid", fgColor="1F4E78")
    fill_yes = PatternFill("solid", fgColor="D1E7DD")
    thin = Side(style="thin", color="BFBFBF")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)

    # โลโก้บริษัท (ถ้ามี)
    if logo_path and os.path.exists(logo_path):
        try:
            from openpyxl.drawing.image import Image as XLImage
            img = XLImage(logo_path)
            img.width = 115
            img.height = 56
            ws.add_image(img, "A1")
        except Exception:
            pass

    # หัวบริษัท (Row 1–4 คอลัมน์ C–G)
    ws.merge_cells("C1:G1"); ws["C1"] = "บริษัท สปีด อินเตอร์ ทรานสปอร์ต จำกัด"; ws["C1"].font = f_co_th
    ws.merge_cells("C2:G2"); ws["C2"] = "SPEED INTER TRANSPORT CO.,LTD"; ws["C2"].font = f_co_en
    ws.merge_cells("C3:G3"); ws["C3"] = "36/34-36 KhlongSong Ton Nun, Lat Krabang, Bangkok 10520."; ws["C3"].font = f_co_sub
    ws.merge_cells("C4:G4"); ws["C4"] = "เลขประจำตัวผู้เสียภาษี : 0105543094654"; ws["C4"].font = f_co_sub
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 17
    ws.row_dimensions[3].height = 15
    ws.row_dimensions[4].height = 15
    ws.row_dimensions[5].height = 8

    # ชื่อเอกสาร & ช่วงวันที่ (Row 6–7)
    ws.merge_cells("A6:G6"); ws["A6"] = "แผนซ่อมบำรุงรักษา - หัวลาก"; ws["A6"].font = f_doc
    ws.merge_cells("A7:G7"); ws["A7"] = date_range_txt; ws["A7"].font = f_date
    ws.row_dimensions[6].height = 22
    ws.row_dimensions[7].height = 18

    # หัวตาราง (Row 8)
    headers = [
        "ทะเบียน",
        "รุ่นรถ",
        "เช็คระยะเปลี่ยน\nถ่ายน้ำมันเครื่อง\n(ระบบไฟ/ระบบลม)\nเช็คทั่วไป 30,000\nkm หรือ 6 เดือน",
        "เช็คระยะเปลี่ยนถ่าย\nน้ำมันเฟืองท้าย/เกียร์\n(ระบบเกียร์) 60,000\nkm หรือ 6 เดือน",
        "เช็คช่วงล่าง\n(ระบบเบรก / ระบบ\nคลัทช์/\nจาระบีลูกปืนล้อ)\n120,000 km",
        "ตรวจเช็คลมยาง /\nดอกยางคงเหลือ",
        "อ้างอิงเอกสาร",
    ]
    widths = [13, 13, 21, 21, 21, 18, 30]
    ws.row_dimensions[8].height = 66
    for j, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(8, j, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.border = b_thin
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = w

    # ข้อมูลรายคัน (Row 9 ..)
    for idx, r in enumerate(rows, start=9):
        ws.row_dimensions[idx].height = 19
        vals = [r[h] for h in headers]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(idx, j, v)
            c.border = b_thin
            if j in (3, 4, 5, 6) and v in ("Yes", "Check Yes"):
                c.font = f_yes
                c.fill = fill_yes
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif j in (1, 2, 3, 4, 5, 6):
                c.font = f_cell
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.font = f_cell
                c.alignment = Alignment(horizontal="left", vertical="center")

    ws.freeze_panes = "C9"
    if rows:
        ws.auto_filter.ref = f"A8:G{8 + len(rows)}"
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:8"
    ws.views.sheetView[0].showGridLines = True


def _df_to_excel_bytes(sheets_dict, report_title="ร่าง Plan PM หัวลาก 2026–2027", subtitle=None, monthly_template=None):
    """Export DataFrames เป็น Excel พร้อมหัวรายงานสวยงาม, Freeze Panes และ Excel Table (ListObject)"""
    import re
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    _ILLEGAL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

    def _clean(v):
        if isinstance(v, str):
            return _ILLEGAL_RE.sub("", v)
        return v

    subtitle_txt = subtitle or f"ข้อมูลสดจากฐานข้อมูล SQLite · พิมพ์เมื่อ {dt.datetime.now():%d/%m/%Y %H:%M} · รอคุณแขก (เจ้าของระบบ) ยืนยัน"

    f_title = Font(name="Tahoma", size=13, bold=True, color="FFFFFF")
    f_sub = Font(name="Tahoma", size=9.5, italic=True, color="1F4E78")
    f_hdr = Font(name="Tahoma", size=10, bold=True, color="FFFFFF")
    f_cell = Font(name="Tahoma", size=10, color="1F2937")
    fill_title = PatternFill("solid", fgColor="1F4E78")
    fill_sub = PatternFill("solid", fgColor="E8F1FA")
    fill_hdr = PatternFill("solid", fgColor="2F5597")
    thin = Side(style="thin", color="D0D7DE")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)

    int_cols = {
        "PM ศูนย์ (20k)", "PM ลาน (30k)", "ค้างยกยอด", "รวม Plan PM ใหม่",
        "ทำ PM จริง", "ใบแจ้งซ่อมทั่วไป (Master)", "Plan เดิมใน KPI", "ส่วนต่าง (ใหม่-เดิม)",
        "จำนวนครั้ง PM ปี 2026", "เลขไมล์ตอน PM", "เลขไมล์ล่าสุด", "เลขไมล์", "รอบ (กม.)", "รอบ (วัน)",
    }
    dec_cols = {"กม./วัน", "% ตามแผน PM"}
    wide_cols = {"รายการน้ำมันเครื่อง", "หมายเหตุ"}
    med_cols = {"ประเภท PM", "ประเภท PM ล่าสุด", "สถานะ ณ ปัจจุบัน", "เลขที่เอกสาร"}

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for idx_s, (name, raw_df) in enumerate(sheets_dict.items(), start=1):
            sh_name = name[:31]
            df = raw_df.copy() if (raw_df is not None and len(raw_df)) else pd.DataFrame({"ไม่มีข้อมูล": []})
            # ทำความสะอาดอักขระควบคุมก่อนเขียน
            for c in df.columns:
                if df[c].dtype == object:
                    df[c] = df[c].map(_clean)
            # เขียนตารางเริ่มที่แถว 4 (แถว 1 = หัวรายงาน, แถว 2 = คำอธิบาย, แถว 3 = เว้นว่าง)
            df.to_excel(xw, sheet_name=sh_name, index=False, startrow=3)
            ws = xw.sheets[sh_name]
            ncols = max(1, len(df.columns))
            nrows = len(df)
            last_col = get_column_letter(ncols)

            # หัวรายงาน (Row 1-2)
            ws.merge_cells(f"A1:{last_col}1")
            ws["A1"] = f"{report_title} — {name}"
            ws["A1"].font = f_title
            ws["A1"].fill = fill_title
            ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.row_dimensions[1].height = 28

            ws.merge_cells(f"A2:{last_col}2")
            ws["A2"] = subtitle_txt
            ws["A2"].font = f_sub
            ws["A2"].fill = fill_sub
            ws["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.row_dimensions[2].height = 20
            ws.row_dimensions[3].height = 8

            # จัดรูปแบบหัวตาราง (Row 4)
            ws.row_dimensions[4].height = 24
            for j, col_name in enumerate(df.columns, start=1):
                c_hdr = ws.cell(4, j)
                c_hdr.font = f_hdr
                c_hdr.fill = fill_hdr
                c_hdr.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                c_hdr.border = b_thin

            # จัดรูปแบบข้อมูล (Row 5 .. 4+nrows)
            for r_i in range(5, 5 + nrows):
                ws.row_dimensions[r_i].height = 19
                for j, col_name in enumerate(df.columns, start=1):
                    cell = ws.cell(r_i, j)
                    cell.font = f_cell
                    cell.border = b_thin
                    if col_name in int_cols:
                        cell.number_format = "#,##0"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    elif col_name in dec_cols:
                        cell.number_format = "#,##0.0"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    elif col_name in ("ทะเบียน", "เดือน", "วันที่", "วันที่ PM ล่าสุด", "ครบกำหนดถัดไป", "ชนิด"):
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

            # ปรับความกว้างคอลัมน์อัตโนมัติ
            for j, col_name in enumerate(df.columns, start=1):
                col_ltr = get_column_letter(j)
                if col_name in wide_cols:
                    ws.column_dimensions[col_ltr].width = 44
                elif col_name in med_cols:
                    ws.column_dimensions[col_ltr].width = 22
                elif str(col_name).startswith(("2026-", "2027-")):
                    ws.column_dimensions[col_ltr].width = 14
                else:
                    max_len = max([len(str(col_name))] + [len(str(v)) for v in df[col_name].head(200).tolist()])
                    ws.column_dimensions[col_ltr].width = max(12, min(32, max_len + 4))

            # Freeze Panes (ตรึงหัวรายงาน + หัวตาราง; ชีต Matrix ตรึงคอลัมน์ทะเบียน + ประเภท PM ด้วย)
            if "ทะเบียน" in df.columns and any(str(c).startswith(("2026-", "2027-")) for c in df.columns):
                ws.freeze_panes = "C5"
            else:
                ws.freeze_panes = "A5"

            # สร้าง Excel Table (ListObject) เมื่อมีแถวข้อมูลอย่างน้อย 1 แถว
            if nrows >= 1:
                tbl_ref = f"A4:{last_col}{4 + nrows}"
                tbl = Table(displayName=f"Tbl_PM_{idx_s}", ref=tbl_ref)
                tbl.tableStyleInfo = TableStyleInfo(
                    name="TableStyleMedium9",
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=False,
                )
                ws.add_table(tbl)

            ws.views.sheetView[0].showGridLines = True

        # เพิ่มชีต 'สรุป_PM_ประจำเดือน' ตาม Template (ถ้ามีการส่ง monthly_template มา)
        if monthly_template:
            _write_monthly_pm_template_sheet(
                xw.book,
                monthly_template["rows"],
                monthly_template["date_range_txt"],
                sheet_name=monthly_template.get("sheet_name", "สรุป_PM_ประจำเดือน"),
                logo_path=monthly_template.get("logo_path"),
            )
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


def _lookup_doc_detail(app_dir, plate, doc_no, doc_kind=None, db_path=None):
    """ดึงรายละเอียดรายการทั้งหมดของเอกสาร W/O หรือ WX จากฐานข้อมูล SQLite (vr_wo_header/line, vr_issue_line) ก่อน แล้วค่อย fallback ไฟล์ .xls"""
    raw_no = str(doc_no or "").strip()
    norm_no = raw_no.split()[-1] if " " in raw_no else raw_no
    kind_u = str(doc_kind or "").upper()

    if db_path and os.path.exists(db_path):
        try:
            con = rr.connect(db_path)
            try:
                if kind_u in ("", "W/O", "WO"):
                    hdr = con.execute(
                        "SELECT * FROM vr_wo_header WHERE plate=? AND doc_no IN (?,?)",
                        (plate, raw_no, norm_no),
                    ).fetchone()
                    lines = con.execute(
                        "SELECT * FROM vr_wo_line WHERE plate=? AND doc_no IN (?,?) ORDER BY line_no",
                        (plate, raw_no, norm_no),
                    ).fetchall()
                    if hdr or lines:
                        return dict(
                            kind="W/O (ใบสั่งซ่อมภายนอก)",
                            file=(hdr["source_file"] if hdr else (lines[0]["source_file"] if lines else "SQLite DB")),
                            doc_no=raw_no,
                            date=(hdr["doc_date"] if hdr else (lines[0]["doc_date"] if lines else "")),
                            vendor=((hdr["vendor_th"] or hdr["vendor_en"]) if hdr else (lines[0]["vendor_en"] if lines else "")),
                            mileage=(hdr["mileage_km"] if hdr else None),
                            net=(hdr["net"] if hdr else sum(float(l["total"] or 0) for l in lines)),
                            remark=(hdr["remark"] or "" if hdr else ""),
                            lines=[
                                {
                                    "ลำดับ": l["line_no"],
                                    "รหัสรายการ": l["item_code"] or "",
                                    "รายละเอียดรายการ": l["item_name"] or "",
                                    "จำนวน": l["qty"],
                                    "ราคา/หน่วย": l["unit_price"],
                                    "จำนวนเงิน (บาท)": l["total"],
                                }
                                for l in lines
                            ],
                        )
                if kind_u in ("", "WX"):
                    lines = con.execute(
                        "SELECT * FROM vr_issue_line WHERE plate=? AND doc_no IN (?,?) ORDER BY line_no",
                        (plate, raw_no, norm_no),
                    ).fetchall()
                    if lines:
                        first = lines[0]
                        return dict(
                            kind="WX (ใบเบิกพัสดุลาน)",
                            file=first["source_file"] or "SQLite DB",
                            doc_no=raw_no,
                            date=first["doc_date"] or "",
                            vendor=f"{first['store'] or ''} {first['dept'] or ''}".strip(),
                            mileage=first["mileage_km"],
                            net=sum(float(l["total"] or 0) for l in lines),
                            remark=first["remark"] or "",
                            lines=[
                                {
                                    "ลำดับ": l["line_no"],
                                    "รหัสพัสดุ": l["inv_code"] or "",
                                    "รายการพัสดุ": l["inv_name"] or "",
                                    "จำนวน": l["qty"],
                                    "ราคา/หน่วย": l["unit_price"],
                                    "จำนวนเงิน (บาท)": l["total"],
                                    "หมายเหตุ": l["remark"] or "",
                                }
                                for l in lines
                            ],
                        )
            finally:
                con.close()
        except Exception:
            pass

    fpath = _find_vehicle_report_file(app_dir, plate)
    if not fpath:
        return None
    rep = _parse_vehicle_report_cached(fpath)

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
        detail = _lookup_doc_detail(app_dir, plate, doc_no, doc_kind, db_path=db)
        if detail:
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("ประเภทเอกสาร", detail["kind"])
            _d_raw = detail["date"] or (str(doc_date) if doc_date else "-")
            if len(str(_d_raw)) >= 10 and str(_d_raw)[4:5] == "-" and str(_d_raw)[7:8] == "-":
                _d_raw = f"{str(_d_raw)[8:10]}/{str(_d_raw)[5:7]}/{str(_d_raw)[0:4]}"
            c2.metric("วันที่", _d_raw)
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
                ds_ = str(doc_date).strip()[:10]
                d0 = dt.datetime.strptime(ds_, "%d/%m/%Y").date() if "/" in ds_ else dt.date.fromisoformat(ds_)
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
                    rel_df = pd.DataFrame(rel)
                    for col_r in ("วันที่แจ้งซ่อม", "วันที่เปิดเอกสาร"):
                        if col_r in rel_df.columns:
                            rel_df[col_r] = pd.to_datetime(rel_df[col_r], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
                    st.markdown(f"**ใบแจ้งซ่อมที่เกี่ยวข้องของรถ `{plate}` (ช่วง ±15 วัน)**")
                    st.dataframe(rel_df, hide_index=True, **full)
            except Exception:
                pass


def _build_fp_mtd_01_02_excel(db_path):
    """สร้างไฟล์ Excel แบบฟอร์ม FP-MTD-01-02 REV.01 รวมทุกประเภทอุปกรณ์ในไฟล์เดียว (แยกชีต หัวลาก / หางพ่วง / ตู้)"""
    import sqlite3
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    # ตรวจสอบคอลัมน์ brand / service_name ที่มีจริง
    am_cols = {c[1] for c in cur.execute("PRAGMA table_info(asset_master)")}
    rr_cols = {c[1] for c in cur.execute("PRAGMA table_info(repair_request)")}
    brand_expr = "brand" if "brand" in am_cols else "''"
    has_svc = "service_name" in rr_cols
    has_km = "current_mileage" in rr_cols

    # ดึงประวัติล่าสุดต่อ (asset_code, category) จาก repair_request.service_name
    latest_cat = {}  # (asset_code, col_key) -> max report_date
    latest_km = {}   # asset_code -> (max_date, km)
    if has_svc:
        q = (
            "SELECT asset_code, report_date, service_name"
            + (", current_mileage" if has_km else ", NULL")
            + " FROM repair_request WHERE asset_code IS NOT NULL AND report_date IS NOT NULL ORDER BY report_date"
        )
        for ac, rdt, sname, km in cur.execute(q):
            sn = str(sname or "")
            if km and float(km) > 0:
                prev = latest_km.get(ac)
                if not prev or str(rdt) >= prev[0]:
                    latest_km[ac] = (str(rdt), float(km))
            keys = []
            if any(k in sn for k in ("เปลี่ยนถ่ายน้ำมันเครื่อง", "เปลี่ยนถ่ายน้ำมันเฟืองท้าย", "เช็คระยะเปลี่ยนถ่าย", "30,000", "60,000", "45,000")):
                keys.append("pm_oil")
            if any(k in sn for k in ("ช่วงล่าง", "ระบบเบรค", "ระบบคลัตซ์", "จาระบีลูกปืนล้อ", "120,000")):
                keys.append("pm_chassis")
            if any(k in sn for k in ("ลมยาง", "ดอกยาง", "มิลยาง", "งานยาง")):
                keys.append("pm_tire")
            if any(k in sn for k in ("หางพ่วง", "ระบบไฟ / ระบบลม")):
                keys.append("pm_trailer")
            if any(k in sn for k in ("ทวิตล็อค", "Twist Lock")):
                keys.append("pm_twist")
            if "20,000" in sn:
                keys.append("km_20k")
            if "30,000" in sn:
                keys.append("km_30k")
            if "45,000" in sn:
                keys.append("km_45k")
            for k in keys:
                if str(rdt) > latest_cat.get((ac, k), ""):
                    latest_cat[(ac, k)] = str(rdt)

    def _fmt_d(iso_d):
        if not iso_d or len(iso_d) < 10:
            return ""
        return f"{iso_d[8:10]}/{iso_d[5:7]}/{iso_d[0:4]}"

    thin = Side(style="thin", color="BFBFBF")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", fgColor="1F4E78")
    sub_fill = PatternFill("solid", fgColor="D9E1F2")
    alt_fill = PatternFill("solid", fgColor="F7FAFC")

    sheet_specs = [
        ("หัวลาก", "หัวลาก", "ทะเบียน"),
        ("หางพ่วง", "หางพ่วง", "ทะเบียน"),
        ("ตู้", "ตู้เย็น", "เบอร์ตู้"),
    ]

    for sh_name, db_atype, plate_lbl in sheet_specs:
        ws = wb.create_sheet(title=sh_name)
        ws.views.sheetView[0].showGridLines = True

        # หัวบริษัท (แถว 1-4)
        ws.merge_cells("A1:L1")
        ws["A1"] = "บริษัท สปีด อินเตอร์ ทรานสปอร์ต จำกัด (สำนักงานใหญ่)"
        ws["A1"].font = Font(name="Tahoma", size=12, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("A2:L2")
        ws["A2"] = "36/34-36 ถนนมอเตอร์เวย์ แขวงคลองสองต้นนุ่น เขตลาดกระบัง กรุงเทพมหานคร 10520"
        ws["A2"].font = Font(name="Tahoma", size=9)
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("A3:L3")
        ws["A3"] = "Head Office : 02-171-7918-20  Fax : 02-171-7922   |   เลขที่ผู้เสียภาษี 0105543094654"
        ws["A3"].font = Font(name="Tahoma", size=9)
        ws["A3"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("A4:L4")
        ws["A4"] = f"แผนและบันทึกการซ่อมบำรุงรักษา — ประเภท: {sh_name}"
        ws["A4"].font = Font(name="Tahoma", size=11, bold=True, color="1F4E78")
        ws["A4"].alignment = Alignment(horizontal="left", vertical="center")

        headers = [
            "ลำดับ",
            "รุ่น",
            plate_lbl,
            "เช็คระยะ/เปลี่ยนถ่าย",
            "เช็คช่วงล่าง/ระบบคลัทซ์/ระบบเบรค/เช็คจารบีลูกปืนล้อ",
            "ตรวจเช็คลมยางเช็คมิลยาง",
            "เช็คสภาพทั่วไปหางพ่วง ระบบไฟ ระบบลม",
            "เช็คทวิตล็อคตู้",
            "ทุก 20,000 กม. (ศูนย์บริการ)",
            "ทุก 30,000 กม.(ลานบริษัท)",
            "ทุก 45,000 กม.(ศูนย์บริการ)",
            "หมายเหตุ",
        ]
        sub_headers = [
            "", "", "",
            "ทุก 6 เดือน", "ทุก 6 เดือน", "ทุก 4 เดือน", "ทุก 4 เดือน", "ทุก 4 เดือน",
            "", "", "", "",
        ]
        for ci, h in enumerate(headers, 1):
            cell = ws.cell(row=6, column=ci, value=h)
            cell.font = Font(name="Tahoma", size=9, bold=True, color="FFFFFF")
            cell.fill = hdr_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = bdr
        ws.row_dimensions[6].height = 36

        for ci, sh_txt in enumerate(sub_headers, 1):
            cell = ws.cell(row=7, column=ci, value=sh_txt)
            cell.font = Font(name="Tahoma", size=8, bold=True, color="1F4E78")
            cell.fill = sub_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = bdr
        ws.row_dimensions[7].height = 20

        assets = cur.execute(
            f"SELECT asset_code, {brand_expr} FROM asset_master WHERE asset_type=? ORDER BY asset_code",
            (db_atype,),
        ).fetchall()

        for idx, (ac, br) in enumerate(assets, 1):
            r_idx = 7 + idx
            km_info = latest_km.get(ac)
            note_txt = f"ไมล์ล่าสุด {int(km_info[1]):,} กม. ({_fmt_d(km_info[0])})" if km_info else ""
            row_vals = [
                idx,
                (br or "") if sh_name == "หัวลาก" else "",
                ac,
                _fmt_d(latest_cat.get((ac, "pm_oil"), "")),
                _fmt_d(latest_cat.get((ac, "pm_chassis"), "")),
                _fmt_d(latest_cat.get((ac, "pm_tire"), "")),
                _fmt_d(latest_cat.get((ac, "pm_trailer"), "")),
                _fmt_d(latest_cat.get((ac, "pm_twist"), "")),
                _fmt_d(latest_cat.get((ac, "km_20k"), "")),
                _fmt_d(latest_cat.get((ac, "km_30k"), "")),
                _fmt_d(latest_cat.get((ac, "km_45k"), "")),
                note_txt,
            ]
            for ci, val in enumerate(row_vals, 1):
                c = ws.cell(row=r_idx, column=ci, value=val)
                c.font = Font(name="Tahoma", size=9)
                c.border = bdr
                c.alignment = Alignment(horizontal="center" if ci != 12 else "left", vertical="center")
                if idx % 2 == 0:
                    c.fill = alt_fill

        footer_row = 7 + len(assets) + 2
        ws.cell(row=footer_row, column=12, value="FP-MTD-01-02 REV.01").font = Font(name="Tahoma", size=8, italic=True, color="555555")
        ws.cell(row=footer_row, column=12).alignment = Alignment(horizontal="right")

        col_widths = [7, 14, 15, 16, 20, 16, 18, 15, 16, 16, 16, 28]
        for ci, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.freeze_panes = "D8"

    conn.close()
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def _render_draft_plan_tab(app_dir, db, full, kpi_card):
    """แท็บแสดงร่าง Plan PM หัวลาก 24 เดือน (ม.ค. 2026 – ธ.ค. 2027) ดึงและคำนวณสดจากฐานข้อมูล SQLite โดยตรง"""
    data = _build_draft_pm_from_db(db)
    if not data:
        st.warning("ยังไม่มีฐานข้อมูล SQLite — กรุณาเลือกไฟล์ `.sqlite` ทางซ้ายก่อน")
        return

    st.info(
        "📌 **ดึงและคำนวณสดจากฐานข้อมูล SQLite (`Repair_Request_v2.sqlite`) โดยตรง** — "
        "อ้างอิงประวัติซ่อมจริง (`vr_wo_header`, `vr_wo_line`, `vr_issue_line` → `service_doc`), `asset_master` และ `mileage_reading` (ม.ค. 2026 – ธ.ค. 2027)"
    )

    st_map = {r["สถานะ"]: r["จำนวน (คัน)"] for r in data["status"]}
    n_master = len(data["detail"]) if len(data["detail"]) else sum(st_map.values())
    n_norm = st_map.get("ตามรอบปกติ", 0)
    n_due_ovd = st_map.get("ครบกำหนด พ.ย. 2026", 0) + st_map.get("เกินกำหนด (ค้าง)", 0)
    n_nohist = st_map.get("ไม่มีประวัติ PM (2026)", 0)

    k = st.columns(4)
    kpi_card(k[0], "หัวลาก Master", f"{n_master:,}", f"คัน (ณ {data['today_str']})")
    kpi_card(k[1], "ตามรอบปกติ", f"{n_norm:,}", "คัน", "ok")
    kpi_card(k[2], "ครบกำหนด พ.ย. + ค้าง", f"{n_due_ovd:,}", "คัน (ยกยอด ธ.ค. 2026)", "warn" if n_due_ovd else "ok")
    kpi_card(k[3], "ไม่มีประวัติ PM 2026", f"{n_nohist:,}", "คัน (รอคุณแขกกำหนดรอบ)", "warn" if n_nohist else "ok")
    st.write("")

    # เลือกเดือนสำหรับชีต 'สรุป_PM_ประจำเดือน' ตาม Template
    avail_ym = sorted(data["hist"]["เดือน"].dropna().astype(str).unique().tolist(), reverse=True) if len(data["hist"]) else ["2026-10"]
    default_ym_idx = avail_ym.index("2026-10") if "2026-10" in avail_ym else 0
    c_mo_tpl, c_dl1, c_dl2 = st.columns([1.3, 2.2, 2.5])
    sel_tpl_ym = c_mo_tpl.selectbox(
        "เดือนสรุป PM (Template)",
        avail_ym,
        index=default_ym_idx,
        key="pm_draft_tpl_ym",
    )
    tpl_rows, tpl_dr_txt, tpl_n_chk = _build_monthly_pm_template_rows(db, sel_tpl_ym)
    _logo_p = None
    try:
        import monthly_report as _mr_logo
        _logo_p = _mr_logo.find_logo_path(app_dir)
    except Exception:
        pass

    full_xl_bytes = _df_to_excel_bytes(
        {
            "สรุป_Plan_PM_2026": pd.DataFrame(data["plan_2026"]),
            "สรุป_Plan_PM_2027": pd.DataFrame(data["plan_2027"]),
            "ตาราง_PM_24เดือน": data["matrix"],
            "รายละเอียดรายคัน": data["detail"],
            "ประวัติ_PM_จริง_2026": data["hist"],
        },
        monthly_template={
            "rows": tpl_rows,
            "date_range_txt": tpl_dr_txt,
            "sheet_name": f"สรุป_PM_{sel_tpl_ym}",
            "logo_path": _logo_p,
        },
    )
    c_dl1.write("")
    c_dl1.download_button(
        f"⬇️ Export ร่าง Plan PM ฉบับเต็ม (+ ชีตสรุป PM {sel_tpl_ym})",
        full_xl_bytes,
        file_name=f"Plan_PM_หัวลาก_2026_2027_{sel_tpl_ym}_{dt.datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key="pm_draft_dl_full",
        **full,
    )
    c_dl2.write("")
    fp_xl_bytes = _build_fp_mtd_01_02_excel(db)
    c_dl2.download_button(
        "📋 Export แบบฟอร์ม FP-MTD-01-02 REV.01 (รวมทุกอุปกรณ์)",
        fp_xl_bytes,
        file_name=f"FP-MTD-01-02_REV01_แผนและบันทึกการซ่อมบำรุงรักษา_{dt.datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="pm_draft_dl_fp_mtd",
        **full,
    )
    st.caption(
        f"แหล่งข้อมูลสด: `{os.path.basename(db)}` · ประวัติ PM จริงปี 2026 รวม **{len(data['hist']):,}** ครั้ง · "
        f"เดือน `{sel_tpl_ym}` เข้า PM **{tpl_n_chk:,}** คัน"
    )

    view = st.radio(
        "เลือกมุมมอง",
        [
            "📊 สรุปรายเดือน 2026–2027",
            "🗓️ ตารางรายคัน (24 เดือน)",
            "📋 รายละเอียดรายคัน",
            "🧾 ประวัติ PM จริง 2026",
            f"📑 สรุป PM ประจำเดือน ({sel_tpl_ym})",
        ],
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

    elif view.startswith("📑 สรุป PM ประจำเดือน"):
        st.markdown(f"**แผนซ่อมบำรุงรักษา - หัวลาก · {tpl_dr_txt}** (เข้า PM จริง **{tpl_n_chk:,}** คัน จากทั้งหมด **{len(tpl_rows):,}** คัน)")
        tpl_df = pd.DataFrame(tpl_rows)
        c_f1, c_f2 = st.columns([1.4, 1.6])
        only_chk = c_f1.checkbox("แสดงเฉพาะรถที่มีการเข้า PM ในเดือนนี้", value=False, key="pm_tpl_only_chk")
        q_tpl = c_f2.text_input("ค้นหาทะเบียน / รุ่นรถ / อ้างอิงเอกสาร", key="pm_tpl_q")
        sub_tpl = tpl_df.copy()
        if only_chk and len(sub_tpl):
            sub_tpl = sub_tpl[sub_tpl["อ้างอิงเอกสาร"].astype(str).str.strip() != ""]
        if q_tpl and len(sub_tpl):
            mask_t = (
                sub_tpl["ทะเบียน"].astype(str).str.contains(q_tpl.strip(), case=False, regex=False)
                | sub_tpl["รุ่นรถ"].astype(str).str.contains(q_tpl.strip(), case=False, regex=False)
                | sub_tpl["อ้างอิงเอกสาร"].astype(str).str.contains(q_tpl.strip(), case=False, regex=False)
            )
            sub_tpl = sub_tpl[mask_t]
        st.caption(f"แสดง {len(sub_tpl):,} คัน")
        st.dataframe(sub_tpl, hide_index=True, height=520, **full)
        # ปุ่ม Export เฉพาะชีตแบบฟอร์มเดือนนี้
        import openpyxl as _oxl_tpl
        _wb_single = _oxl_tpl.Workbook()
        _wb_single.remove(_wb_single.active)
        _write_monthly_pm_template_sheet(
            _wb_single,
            tpl_rows,
            tpl_dr_txt,
            sheet_name=f"สรุป_PM_{sel_tpl_ym}",
            logo_path=_logo_p,
        )
        _buf_s = io.BytesIO()
        _wb_single.save(_buf_s)
        st.download_button(
            f"⬇️ Export เฉพาะแบบฟอร์ม สรุป PM ประจำเดือน ({sel_tpl_ym})",
            _buf_s.getvalue(),
            file_name=f"สรุป_PM_หัวลาก_{sel_tpl_ym}_{dt.datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="pm_d_dl_tpl_only",
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
        "📅 ร่าง Plan PM 2026–2027 (รอคุณแขกยืนยัน)",
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
            st.caption("ค่าที่เปลี่ยนใช้เฉพาะการเปิดครั้งนี้ — ค่าเริ่มต้นตามที่คุณแขก (เจ้าของระบบ) กำหนด (20,000 / 30,000 กม. / 6 เดือน)")

    # ------------------------------------------------------------------ all vehicles
    with t_all:
        adf = pd.DataFrame(rows)
        if len(adf):
            c1, c2, c3 = st.columns([1.4, 1.4, 1.6])
            fs = c1.multiselect("สถานะ", ["เกินกำหนด", "ครบกำหนดเดือนนี้", "ยังไม่ถึง", "ไม่มีประวัติ PM"], key="pm_fs", placeholder="ทุกสถานะ")
            ft = c2.multiselect("PM ครั้งล่าสุดที่", ["ศูนย์บริการ", "เปลี่ยนเครื่องยนต์", "ลาน", "ยังไม่มีประวัติ PM"], key="pm_ft", placeholder="ทั้งหมด")
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
                    if "doc_date" in ev.columns:
                        ev["doc_date"] = pd.to_datetime(ev["doc_date"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
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
        sb_df = pd.DataFrame(pm.fetch(db, "SELECT batch_id, file_name ไฟล์, imported_by ผู้นำเข้า, imported_at เวลา, rows_read อ่าน, "
                                          "rows_inserted ใหม่, rows_updated อัปเดต, rows_unchanged ไม่เปลี่ยน, backup_path ไฟล์สำรอง "
                                          "FROM service_batch ORDER BY batch_id DESC"))
        if len(sb_df) and "เวลา" in sb_df.columns:
            sb_df["เวลา"] = pd.to_datetime(sb_df["เวลา"], errors="coerce").dt.strftime("%d/%m/%Y %H:%M").fillna(sb_df["เวลา"])
        st.dataframe(sb_df, hide_index=True, **full)

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
            if len(evs) and "วันที่" in evs.columns:
                evs["วันที่"] = pd.to_datetime(evs["วันที่"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
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
