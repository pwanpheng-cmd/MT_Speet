# -*- coding: utf-8 -*-
"""
Monthly KPI & Maintenance Print Report Engine — Speed Fleet Maintenance (v1.8.0)
- จัดการทะเบียนรถหลัก (asset_master: หัวลาก / หางพ่วง / ตู้เย็น) ใน SQLite
- คำนวณสรุป KPI ประจำเดือน และรายปี (แยกใน Master vs นอก Master รอ DATA_OWNER ตรวจสอบ)
- สร้างไฟล์ Excel (.xlsx) พร้อมพิมพ์ตาม Template บริษัท (1_หัวลาก / 2_หางพ่วง / 3_ซ่อมตู้ + สรุป KPI)
- สร้างหน้า HTML Print View พร้อมสั่งพิมพ์ / Save as PDF จากเบราว์เซอร์ (หัวกระดาษซ้ำทุกหน้า)
"""
import base64
import calendar
import datetime as dt
import html
import io
import math
import os
import re

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import repair_request_db as rr

TYPES = ["หัวลาก", "หางพ่วง", "ตู้เย็น"]

COMPANY_INFO = dict(
    name_th="บริษัท สปีด อินเตอร์ ทรานสปอร์ต จำกัด",
    name_en="SPEED INTER TRANSPORT CO.,LTD",
    address="36/34-36 KhlongSong Ton Nun, Lat Krabang, Bangkok 10520.",
    tax_id="เลขประจำตัวผู้เสียภาษี : 0105543094654",
)

# โครงสร้างคอลัมน์และตำแหน่งลงเลขที่ใบแจ้งซ่อมตาม Template บริษัท (แก้คำผิดใน 2_หางพ่วง และ 3_ซ่อมตู้ ตามอนุมัติ)
FORM_SPECS = {
    "หัวลาก": dict(
        sheet_name="1_หัวลาก",
        title_label="หัวลาก",
        doc_col_idx=6,  # คอลัมน์ F: อื่นๆและแจ้งซ่อมทั่วไป
        columns=[
            ("ทะเบียน", 14.5),
            ("เช็คระยะเปลี่ยน\nถ่ายน้ำมันเครื่อง\n(ระบบไฟ/ระบบลม)\nเช็คทั่วไป 30,000\nkm หรือ 6 เดือน", 16.0),
            ("เช็คระยะเปลี่ยนถ่าย\nน้ำมันเฟืองท้าย/เกียร์\n(ระบบเกียร์) 60,000\nkm หรือ 6 เดือน", 19.5),
            ("เช็คช่วงล่าง\n(ระบบเบรก / ระบบ\nคลัทช์/\nจาระบีลูกปืนล้อ)\n120,000 km", 20.5),
            ("ตรวจเช็คลมยาง /\nดอกยางคงเหลือ", 16.0),
            ("อื่นๆและแจ้งซ่อมทั่วไป", 32.0),
        ],
        html_widths=["12%", "15%", "16%", "17%", "14%", "26%"],
        helper_col_idx=8,   # H: นับจำนวนที่เข้าซ่อม
        kpi_label_col=9,    # I
        kpi_val_col=10,     # J
    ),
    "หางพ่วง": dict(
        sheet_name="2_หางพ่วง",
        title_label="หางพ่วง",
        doc_col_idx=5,  # คอลัมน์ E: อื่นๆและแจ้งซ่อมทั่วไป
        columns=[
            ("ทะเบียน", 15.0),
            ("เช็คสภาพทั่วไปหางพ่วง\n(ระบบไฟ / ระบบลม) /\nตรวจเช็คทวิตล็อคตู้ (Twist Lock)", 22.0),
            ("เช็คช่วงล่าง\n( ระบบเบรค / ระบบคลัทซ์ /\nจาระบีลูกปืนล้อ )\n120,000 km", 22.0),
            ("ตรวจเช็คลมยาง /\nดอกยางคงเหลือ", 18.0),
            ("อื่นๆและแจ้งซ่อมทั่วไป", 34.0),
        ],
        html_widths=["14%", "21%", "21%", "16%", "28%"],
        helper_col_idx=7,   # G: นับจำนวนที่เข้าซ่อม
        kpi_label_col=8,    # H
        kpi_val_col=9,      # I
    ),
    "ตู้เย็น": dict(
        sheet_name="3_ซ่อมตู้",
        title_label="ตู้คอนเทนเนอร์เย็น",
        doc_col_idx=2,  # คอลัมน์ B: เช็คตู้คอนเทนเนอร์ระบบการทำอุณหภูมิความเย็น
        columns=[
            ("ทะเบียน", 24.0),
            ("เช็คตู้คอนเทนเนอร์ระบบการทำอุณหภูมิความเย็น\n( หน้าเครื่อง / พัดลมคอยล์ร้อนและเย็น / คอยล์ร้อน\nและเย็น / คอมแอร์ / ท่อน้ำยาแอร์ / สายไฟ ฯ )", 42.0),
            ("อื่นๆและแจ้งซ่อมทั่วไป", 32.0),
        ],
        html_widths=["22%", "48%", "30%"],
        helper_col_idx=5,   # E: นับจำนวนที่เข้าซ่อม
        kpi_label_col=6,    # F
        kpi_val_col=7,      # G
    ),
}

ASSET_MASTER_SQL = """
CREATE TABLE IF NOT EXISTS asset_master (
  asset_type  TEXT NOT NULL,
  asset_code  TEXT NOT NULL,
  is_active   INTEGER NOT NULL DEFAULT 1,
  sort_order  INTEGER NOT NULL DEFAULT 0,
  note        TEXT,
  updated_at  TEXT NOT NULL DEFAULT (datetime('now','localtime')),
  updated_by  TEXT,
  PRIMARY KEY (asset_type, asset_code)
);
CREATE INDEX IF NOT EXISTS idx_am_type ON asset_master(asset_type, is_active, sort_order, asset_code);
"""

THAI_MONTHS = [
    "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม",
]


def connect_master(db_path):
    con = rr.connect(db_path)
    con.executescript(ASSET_MASTER_SQL)
    return con


def month_bounds(report_month):
    """'2026-01' -> (date(2026,1,1), date(2026,1,31))"""
    m = re.fullmatch(r"(\d{4})-(\d{2})", str(report_month).strip())
    if not m:
        raise ValueError(f"รูปแบบเดือนไม่ถูกต้อง: {report_month} (ต้องเป็น YYYY-MM)")
    y, mo = int(m.group(1)), int(m.group(2))
    last_day = calendar.monthrange(y, mo)[1]
    return dt.date(y, mo, 1), dt.date(y, mo, last_day)


def format_month_th(report_month):
    d1, _ = month_bounds(report_month)
    return f"{THAI_MONTHS[d1.month]} {d1.year + 543} ({report_month})"


def available_months(db_path):
    if not os.path.exists(db_path):
        return []
    con = connect_master(db_path)
    try:
        rows = con.execute(
            "SELECT DISTINCT report_month FROM repair_request "
            "WHERE report_month IS NOT NULL AND report_month != '' ORDER BY report_month DESC"
        ).fetchall()
        return [r[0] for r in rows]
    finally:
        con.close()


# ============================================================================= Asset Master CRUD & Import
MASTER_SHEET_MAP = {
    "หัวลาก": ["F1_หัวลาก", "1_หัวลาก"],
    "หางพ่วง": ["F2_หางพ่วง", "2_ห่างพ่วง", "2_หางพ่วง"],
    "ตู้เย็น": ["F3_ตู้เย็น", "3_ซ่อมตู้"],
}


def read_master_from_excel(source):
    """อ่านรายชื่อทะเบียนรถหลักจากไฟล์ Template บริษัท (F1/F2/F3 หรือ 1_หัวลาก/2_ห่างพ่วง/3_ซ่อมตู้)
    คืน dict: {records: {asset_type: [(code, sort_order), ...]}, sheets: [...], errors: [...]}"""
    wb = openpyxl.load_workbook(source, data_only=True, read_only=True)
    out = dict(records={}, sheets=[], errors=[])
    try:
        names = wb.sheetnames
        for at, candidates in MASTER_SHEET_MAP.items():
            picked = next((c for c in candidates if c in names), None)
            if not picked:
                continue
            ws = wb[picked]
            seen = set()
            items = []
            order = 1
            for r_idx, row in enumerate(ws.iter_rows(min_row=10, max_col=1, values_only=True), start=10):
                val = row[0] if row else None
                if val is None:
                    continue
                code = str(val).strip()
                if not code or code.startswith("=") or code in ("ทะเบียน", "รวม"):
                    continue
                if code not in seen:
                    seen.add(code)
                    items.append((code, order))
                    order += 1
            out["records"][at] = items
            out["sheets"].append(dict(asset_type=at, sheet=picked, count=len(items)))
    finally:
        wb.close()
    if not out["records"]:
        out["errors"].append("ไม่พบชีตรายชื่อรถหลัก (F1_หัวลาก / F2_หางพ่วง / F3_ตู้เย็น หรือ 1_หัวลาก / 2_ห่างพ่วง / 3_ซ่อมตู้)")
    return out


def import_asset_master_from_excel(db_path, source, user=None, replace_mode=False):
    """นำเข้า/อัปเดตรายชื่อทะเบียนรถหลักเข้า asset_master (สำรอง DB ก่อนเสมอ)"""
    parsed = read_master_from_excel(source)
    if parsed["errors"]:
        return dict(ok=False, errors=parsed["errors"], counts={})
    bpath = rr.backup_db(db_path) if os.path.exists(db_path) else None
    con = connect_master(db_path)
    counts = {}
    try:
        with con:
            for at, items in parsed["records"].items():
                if replace_mode:
                    con.execute("UPDATE asset_master SET is_active=0, updated_at=datetime('now','localtime'), updated_by=? WHERE asset_type=?",
                                (user, at))
                inserted = 0
                updated = 0
                for code, order in items:
                    ex = con.execute("SELECT 1 FROM asset_master WHERE asset_type=? AND asset_code=?", (at, code)).fetchone()
                    if ex:
                        con.execute(
                            "UPDATE asset_master SET is_active=1, sort_order=?, updated_at=datetime('now','localtime'), updated_by=? "
                            "WHERE asset_type=? AND asset_code=?",
                            (order, user, at, code),
                        )
                        updated += 1
                    else:
                        con.execute(
                            "INSERT INTO asset_master(asset_type, asset_code, is_active, sort_order, note, updated_by) "
                            "VALUES (?,?,1,?,?,?)",
                            (at, code, order, "นำเข้าจากไฟล์ทะเบียนรถหลัก", user),
                        )
                        inserted += 1
                counts[at] = dict(total=len(items), inserted=inserted, updated=updated)
            rr.sync_m1_iso(con)
        return dict(ok=True, errors=[], counts=counts, sheets=parsed["sheets"], backup=bpath)
    finally:
        con.close()


def upsert_asset_master(db_path, asset_type, asset_code, is_active=1, note=None, user=None):
    code = str(asset_code or "").strip()
    if asset_type not in TYPES or not code:
        return False, "กรุณาระบุประเภทและทะเบียน/เบอร์ตู้ให้ครบ"
    con = connect_master(db_path)
    try:
        with con:
            ex = con.execute("SELECT sort_order FROM asset_master WHERE asset_type=? AND asset_code=?", (asset_type, code)).fetchone()
            if ex:
                con.execute(
                    "UPDATE asset_master SET is_active=?, note=COALESCE(?, note), updated_at=datetime('now','localtime'), updated_by=? "
                    "WHERE asset_type=? AND asset_code=?",
                    (1 if is_active else 0, note, user, asset_type, code),
                )
            else:
                mx = con.execute("SELECT IFNULL(MAX(sort_order), 0) FROM asset_master WHERE asset_type=?", (asset_type,)).fetchone()[0]
                con.execute(
                    "INSERT INTO asset_master(asset_type, asset_code, is_active, sort_order, note, updated_by) VALUES (?,?,?,?,?,?)",
                    (asset_type, code, 1 if is_active else 0, mx + 1, note, user),
                )
            if asset_type == "หัวลาก":
                rr.sync_m1_iso(con)
        return True, f"บันทึก {asset_type}: {code} เรียบร้อย"
    finally:
        con.close()


def get_master_summary(db_path):
    if not os.path.exists(db_path):
        return {t: dict(active=0, inactive=0) for t in TYPES}
    con = connect_master(db_path)
    try:
        res = {t: dict(active=0, inactive=0) for t in TYPES}
        for r in con.execute("SELECT asset_type, is_active, COUNT(*) n FROM asset_master GROUP BY asset_type, is_active"):
            if r["asset_type"] in res:
                res[r["asset_type"]]["active" if r["is_active"] else "inactive"] = r["n"]
        return res
    finally:
        con.close()


def list_asset_master(db_path, asset_type=None, include_inactive=True):
    if not os.path.exists(db_path):
        return []
    con = connect_master(db_path)
    try:
        sql = "SELECT asset_type, asset_code, is_active, sort_order, note, updated_at, updated_by FROM asset_master WHERE 1=1"
        args = []
        if asset_type:
            sql += " AND asset_type=?"
            args.append(asset_type)
        if not include_inactive:
            sql += " AND is_active=1"
        sql += " ORDER BY asset_type, is_active DESC, sort_order, asset_code"
        return [dict(r) for r in con.execute(sql, args).fetchall()]
    finally:
        con.close()


def bulk_update_asset_master(db_path, asset_type, codes_list, replace_all=False, note=None, user=None):
    """อัปเดตทะเบียนรถหลักแบบหลายรายการพร้อมกัน (สำรอง DB ก่อนเสมอ)"""
    if asset_type not in TYPES:
        return dict(ok=False, msg="ประเภทไม่ถูกต้อง", total=0)
    clean = []
    seen = set()
    for raw in codes_list:
        c = str(raw or "").strip()
        if c and c not in seen and c not in ("ทะเบียน", "หัวลาก", "หางพ่วง", "ตู้เย็น"):
            seen.add(c)
            clean.append(c)
    if not clean:
        return dict(ok=False, msg="ไม่พบทะเบียน/เบอร์ตู้ที่ใช้งานได้", total=0)
    bpath = rr.backup_db(db_path) if os.path.exists(db_path) else None
    con = connect_master(db_path)
    inserted = updated = 0
    try:
        with con:
            if replace_all:
                con.execute("DELETE FROM asset_master WHERE asset_type=?", (asset_type,))
                for idx, code in enumerate(clean, start=1):
                    con.execute(
                        "INSERT INTO asset_master(asset_type, asset_code, is_active, sort_order, note, updated_by) "
                        "VALUES (?,?,1,?,?,?)",
                        (asset_type, code, idx, note or "แทนที่รายชื่อทั้งหมดจากหน้าจัดการอุปกรณ์", user),
                    )
                    inserted += 1
            else:
                mx = con.execute("SELECT IFNULL(MAX(sort_order), 0) FROM asset_master WHERE asset_type=?", (asset_type,)).fetchone()[0]
                for code in clean:
                    ex = con.execute("SELECT 1 FROM asset_master WHERE asset_type=? AND asset_code=?", (asset_type, code)).fetchone()
                    if ex:
                        con.execute(
                            "UPDATE asset_master SET is_active=1, note=COALESCE(?, note), "
                            "updated_at=datetime('now','localtime'), updated_by=? WHERE asset_type=? AND asset_code=?",
                            (note, user, asset_type, code),
                        )
                        updated += 1
                    else:
                        mx += 1
                        con.execute(
                            "INSERT INTO asset_master(asset_type, asset_code, is_active, sort_order, note, updated_by) "
                            "VALUES (?,?,1,?,?,?)",
                            (asset_type, code, mx, note or "เพิ่มจากหน้าจัดการอุปกรณ์", user),
                        )
                        inserted += 1
            if asset_type == "หัวลาก":
                rr.sync_m1_iso(con)
        return dict(ok=True, msg="บันทึกเรียบร้อย", total=len(clean), inserted=inserted, updated=updated, backup=bpath)
    finally:
        con.close()


def export_asset_master_excel(db_path):
    """Export รายการอุปกรณ์ทั้งหมดเป็น Excel (ชีตรวม + แยกชีต หัวลาก/หางพ่วง/ตู้เย็น)"""
    rows = list_asset_master(db_path, include_inactive=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    f_hdr = Font(name="Arial", size=10, bold=True)
    f_cell = Font(name="Arial", size=10)
    fill_hdr = PatternFill("solid", fgColor="DCE9F5")
    b_thin = _thin_border()

    sheets_spec = [("รวมทั้งหมด", rows)] + [(t, [r for r in rows if r["asset_type"] == t]) for t in TYPES]
    headers = ["ลำดับ", "ประเภท", "ทะเบียน/เบอร์ตู้", "สถานะ", "หมายเหตุ", "อัปเดตล่าสุด", "โดย"]
    widths = [8, 13, 20, 12, 34, 20, 16]
    for sname, srows in sheets_spec:
        ws = wb.create_sheet(sname)
        for j, (h, w) in enumerate(zip(headers, widths), 1):
            c = ws.cell(1, j, h)
            c.font = f_hdr; c.fill = fill_hdr; c.border = b_thin
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(j)].width = w
        for idx, r in enumerate(srows, start=1):
            vals = [idx, r["asset_type"], r["asset_code"], "ใช้งาน" if r["is_active"] else "ปิดใช้งาน",
                    r["note"] or "", r["updated_at"] or "", r["updated_by"] or ""]
            for j, v in enumerate(vals, 1):
                c = ws.cell(idx + 1, j, v)
                c.font = f_cell; c.border = b_thin
                c.alignment = Alignment(horizontal="center" if j in (1, 2, 3, 4) else "left", vertical="center")
        ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ============================================================================= Monthly Report & KPI Computation
def compute_monthly_report(db_path, report_month):
    """คำนวณข้อมูลรายงานซ่อมประจำเดือน + KPI สำหรับเดือนที่เลือก
    - แบบ ข (ตามอนุมัติ): แบบฟอร์มแสดงเฉพาะทะเบียนที่อยู่ใน asset_master (active)
    - ทะเบียนที่มีใบซ่อมในเดือนนั้นแต่ไม่อยู่ใน asset_master แยกไว้ใน unmatched ให้ DATA_OWNER ตัดสินใจ
    - รถ 1 คันมีหลายใบแจ้งซ่อมในเดือนเดียวกัน แสดงทุกใบคั่นด้วย ', ' (เรียงตามวันที่แจ้งซ่อม, เลขที่เอกสาร)"""
    d1, d2 = month_bounds(report_month)
    con = connect_master(db_path)
    try:
        by_type = {}
        for at in TYPES:
            masters = con.execute(
                "SELECT asset_code, is_active, sort_order FROM asset_master WHERE asset_type=? ORDER BY sort_order, asset_code",
                (at,),
            ).fetchall()
            active_list = [r["asset_code"] for r in masters if r["is_active"]]
            active_set = set(active_list)
            inactive_set = {r["asset_code"] for r in masters if not r["is_active"]}

            rep_rows = con.execute(
                "SELECT doc_no, report_date, asset_code, asset_raw, brand, status "
                "FROM repair_request WHERE asset_type=? AND report_month=? AND in_iso='Yes' "
                "ORDER BY report_date, doc_no",
                (at, report_month),
            ).fetchall()

            docs_by_asset = {}
            meta_by_asset = {}
            for r in rep_rows:
                code = (r["asset_code"] or "").strip()
                if not code:
                    continue
                docs_by_asset.setdefault(code, [])
                if r["doc_no"] not in docs_by_asset[code]:
                    docs_by_asset[code].append(r["doc_no"])
                m = meta_by_asset.setdefault(code, dict(dates=[], brands=set(), statuses=set(), raw=r["asset_raw"]))
                if r["report_date"]:
                    m["dates"].append(r["report_date"])
                if r["brand"]:
                    m["brands"].add(r["brand"])
                if r["status"]:
                    m["statuses"].add(r["status"])

            form_rows = []
            repaired_in_master = 0
            docs_in_master = 0
            for code in active_list:
                docs = docs_by_asset.get(code, [])
                has_rep = 1 if docs else 0
                repaired_in_master += has_rep
                docs_in_master += len(docs)
                form_rows.append(dict(
                    asset_code=code,
                    doc_nos=docs,
                    doc_text=", ".join(docs),
                    has_repair=has_rep,
                    repair_count=len(docs),
                ))

            unmatched = []
            for code in sorted(k for k in docs_by_asset if k not in active_set):
                docs = docs_by_asset[code]
                m = meta_by_asset[code]
                unmatched.append(dict(
                    asset_type=at,
                    asset_code=code,
                    asset_raw=m["raw"],
                    repair_count=len(docs),
                    doc_nos=docs,
                    doc_text=", ".join(docs),
                    first_date=min(m["dates"]) if m["dates"] else "",
                    last_date=max(m["dates"]) if m["dates"] else "",
                    brands=", ".join(sorted(m["brands"])),
                    statuses=", ".join(sorted(m["statuses"])),
                    in_master_inactive=(code in inactive_set),
                ))

            total_master = len(active_list)
            pct = (repaired_in_master / total_master) if total_master > 0 else 0.0
            by_type[at] = dict(
                spec=FORM_SPECS[at],
                rows=form_rows,
                unmatched=unmatched,
                kpi=dict(
                    total_master=total_master,
                    repaired_in_master=repaired_in_master,
                    pct_repaired=pct,
                    docs_in_master=docs_in_master,
                    unmatched_assets=len(unmatched),
                    unmatched_docs=sum(x["repair_count"] for x in unmatched),
                    total_db_assets=len(docs_by_asset),
                    total_db_docs=len(rep_rows),
                ),
            )

        return dict(
            report_month=report_month,
            month_th=format_month_th(report_month),
            start_date=d1,
            end_date=d2,
            date_range_text=f"ตั้งแต่วันที่ {d1:%d/%m/%Y} - {d2:%d/%m/%Y}",
            by_type=by_type,
        )
    finally:
        con.close()


def compute_yearly_kpi(db_path, year=None):
    """ตาราง KPI รายเดือนทั้งปี (หรือทุกเดือนที่มีใน DB): แยกใน Master vs นอก Master"""
    con = connect_master(db_path)
    try:
        masters = {t: set() for t in TYPES}
        for r in con.execute("SELECT asset_type, asset_code FROM asset_master WHERE is_active=1"):
            if r["asset_type"] in masters:
                masters[r["asset_type"]].add(r["asset_code"])

        sql = "SELECT report_month, asset_type, asset_code, doc_no FROM repair_request WHERE report_month IS NOT NULL AND in_iso='Yes'"
        args = []
        if year:
            sql += " AND report_month LIKE ?"
            args.append(f"{year}-%")
        rows = con.execute(sql, args).fetchall()

        bucket = {}
        for r in rows:
            ym, at, code = r["report_month"], r["asset_type"], (r["asset_code"] or "").strip()
            if at not in masters or not code:
                continue
            b = bucket.setdefault((ym, at), dict(in_m=set(), out_m=set(), docs_in=0, docs_out=0))
            if code in masters[at]:
                b["in_m"].add(code)
                b["docs_in"] += 1
            else:
                b["out_m"].add(code)
                b["docs_out"] += 1

        months = sorted({k[0] for k in bucket})
        out = []
        for ym in months:
            rec = dict(report_month=ym)
            for at in TYPES:
                tot = len(masters[at])
                b = bucket.get((ym, at), dict(in_m=set(), out_m=set(), docs_in=0, docs_out=0))
                in_n = len(b["in_m"])
                out_n = len(b["out_m"])
                rec[f"{at}_total"] = tot
                rec[f"{at}_in_master"] = in_n
                rec[f"{at}_pct"] = (in_n / tot) if tot > 0 else 0.0
                rec[f"{at}_unmatched"] = out_n
                rec[f"{at}_db_assets"] = in_n + out_n
                rec[f"{at}_db_docs"] = b["docs_in"] + b["docs_out"]
            out.append(rec)
        return out
    finally:
        con.close()


DEFAULT_EXPORT_DIR = r"D:\MT_Speet\Export_File"


def get_export_dir(app_dir=None):
    """คืนโฟลเดอร์เก็บไฟล์ Export (ค่าเริ่มต้น <app_dir>/Export_File เช่น D:\\MT_Speet\\Export_File)"""
    if app_dir:
        return os.path.join(app_dir, "Export_File")
    return DEFAULT_EXPORT_DIR


def save_local_export(app_dir, file_name, data_bytes, is_local=True):
    """บันทึกไฟล์ Export ลงโฟลเดอร์ Export_File อัตโนมัติเมื่อรันบนเครื่อง (โหมด Local) — คืน path หรือ None"""
    if not is_local:
        return None
    try:
        out_dir = get_export_dir(app_dir)
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, os.path.basename(file_name))
        with open(out_path, "wb") as fh:
            fh.write(data_bytes)
        return out_path
    except Exception:
        return None


# ============================================================================= Logo & ISO KPI Template helpers
def find_logo_path(app_dir, explicit_path=None):
    if explicit_path and os.path.exists(explicit_path):
        return explicit_path
    for cand in (
        os.path.join(app_dir, "data", "company_logo.png"),
        os.path.join(app_dir, "static", "company_logo.png"),
    ):
        if os.path.exists(cand):
            return cand
    return None


def find_kpi_template_path(app_dir, explicit_path=None):
    if explicit_path and os.path.exists(explicit_path):
        return explicit_path
    for cand in (
        os.path.join(app_dir, "data", "kpi_template_iso.xlsx"),
        os.path.join(app_dir, "static", "kpi_template_iso.xlsx"),
    ):
        if os.path.exists(cand):
            return cand
    return None


# ---- ข้อมูลเฉพาะบริษัท (แผน KPI รายเดือน, %Achieve ปีก่อน, หมายเหตุ Last PM รายคัน)
# เก็บในไฟล์ data/local_settings.json บนเครื่อง (โฟลเดอร์ data ไม่ขึ้น GitHub) — โค้ดใน repo ไม่มีข้อมูลจริง
# รูปแบบไฟล์: {"kpi_plans": {"หัวลาก": [12 ค่า], ...}, "ave_prev": {"หัวลาก": 85.0, ...}, "last_pm_notes": {"ทะเบียน": "หมายเหตุ"}}
LOCAL_SETTINGS_NAME = "local_settings.json"


def _local_settings_candidates():
    here = os.path.dirname(os.path.abspath(__file__))
    env = os.environ.get("MT_SPEET_SETTINGS")
    return [p for p in (env, os.path.join(here, "data", LOCAL_SETTINGS_NAME),
                        os.path.join(os.path.dirname(here), "data", LOCAL_SETTINGS_NAME)) if p]


def load_local_settings():
    import json
    for p in _local_settings_candidates():
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as fh:
                    return json.load(fh), p
            except Exception:
                return {}, p
    return {}, None


_LOCAL, LOCAL_SETTINGS_PATH = load_local_settings()
DEFAULT_KPI_PLANS = {t: list((_LOCAL.get("kpi_plans") or {}).get(t) or [0] * 12)[:12] for t in TYPES}
DEFAULT_AVE_PREV = dict(_LOCAL.get("ave_prev") or {})
LAST_PM_DEFAULT_NOTES = dict(_LOCAL.get("last_pm_notes") or {})

_EN_MONTH_UPPER = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]

ISO_SHEET_BY_TYPE = {
    "หัวลาก": "MTD-หัวลาก",
    "หางพ่วง": "MTD-หางพ่วง",
    "ตู้เย็น": "MTD-ตู้เย็น",
}


def _strip_pivot_bloat_bytes(raw_bytes):
    """ล้าง pivotCache / pivotTables / externalLinks / calcChain ออกจากไฟล์ Excel Template (แก้ปัญหา Pivot Cache บวม ~146 MB)"""
    import zipfile
    import xml.etree.ElementTree as ET
    out_buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(raw_bytes), "r") as zin, zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            fn = item.filename
            if any(k in fn for k in ("pivotCache", "pivotTables", "externalLinks", "calcChain")):
                continue
            data = zin.read(fn)
            if fn == "xl/workbook.xml":
                ET.register_namespace("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
                ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                root = ET.fromstring(data)
                for tag in ("pivotCaches", "externalReferences"):
                    for el in list(root):
                        if el.tag.endswith("}" + tag) or el.tag == tag:
                            root.remove(el)
                # ล้าง Named Ranges ขยะ/เสีย (#REF!, [0]!, [1]!, Macro เก่า) ที่ทำให้ Excel แจ้ง repair log
                for el in list(root):
                    if el.tag.endswith("}definedNames") or el.tag == "definedNames":
                        for dn in list(el):
                            nm = dn.attrib.get("name", "")
                            txt = (dn.text or "").strip()
                            is_print = nm in ("_xlnm.Print_Area", "_xlnm.Print_Titles")
                            is_bad = ("#REF!" in txt) or ("[" in txt and "]" in txt) or (not is_print)
                            if is_bad:
                                el.remove(dn)
                        if len(el) == 0:
                            root.remove(el)
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            elif fn.endswith(".rels"):
                ET.register_namespace("", "http://schemas.openxmlformats.org/package/2006/relationships")
                root = ET.fromstring(data)
                for el in list(root):
                    tgt = el.attrib.get("Target", "")
                    typ = el.attrib.get("Type", "")
                    if any(k in tgt or k in typ for k in ("pivotCache", "pivotTable", "externalLink")):
                        root.remove(el)
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            elif fn == "[Content_Types].xml":
                ET.register_namespace("", "http://schemas.openxmlformats.org/package/2006/content-types")
                root = ET.fromstring(data)
                for el in list(root):
                    pn = el.attrib.get("PartName", "")
                    if any(k in pn for k in ("pivotCache", "pivotTables", "externalLinks", "calcChain")):
                        root.remove(el)
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            zout.writestr(item, data)
    return out_buf.getvalue()


def _build_iso_kpi_workbook(
    db_path,
    year,
    template_source,
    use_master_only=True,
    plans_by_type=None,
    target_pct=80,
    meta=None,
    include_plate_matrix=True,
    max_month=None,
):
    """สร้าง openpyxl Workbook สำหรับแบบฟอร์ม ISO FP-MR-00-01 Rev.00
    (MTD-หัวลาก / MTD-หางพ่วง / MTD-ตู้เย็น + ทะเบียนทั้งหมด)
    - max_month: ถ้าระบุ (1..12) จะเติมข้อมูลเฉพาะเดือน Jan..max_month และเว้นเดือนถัดไปว่าง
    คืน (wb, summary_dict)"""
    year_str = str(year).strip()
    year_int = int(year_str)
    max_m = int(max_month) if max_month is not None else 12
    if isinstance(template_source, (bytes, bytearray)):
        raw_tpl = bytes(template_source)
    else:
        with open(template_source, "rb") as fh:
            raw_tpl = fh.read()

    clean_tpl = _strip_pivot_bloat_bytes(raw_tpl)
    wb = openpyxl.load_workbook(io.BytesIO(clean_tpl))

    # ลบชีตที่ไม่ใช่ 3 ชีตหลัก ISO (ถ้าผู้ใช้อัปโหลดไฟล์ต้นฉบับที่มีหลายชีต)
    keep_sheets = list(ISO_SHEET_BY_TYPE.values())
    for sname in list(wb.sheetnames):
        if sname not in keep_sheets:
            del wb[sname]
    # ล้าง workbook-level & sheet-level defined_names ขยะ (openpyxl จะสร้าง Print_Area/Print_Titles ใหม่ให้ถูกต้องตอน save)
    wb.defined_names.clear()
    for ws_tmp in wb.worksheets:
        ws_tmp.defined_names.clear()
        if not ws_tmp.print_area:
            ws_tmp.print_area = "A1:AP37"

    yearly = compute_yearly_kpi(db_path, year=year_str)
    by_month = {int(r["report_month"].split("-")[1]): r for r in yearly if int(r["report_month"].split("-")[1]) <= max_m}
    plans = plans_by_type or DEFAULT_KPI_PLANS
    meta = meta or {}
    summary = {}

    for at, sname in ISO_SHEET_BY_TYPE.items():
        if sname not in wb.sheetnames:
            continue
        ws = wb[sname]
        # ล้างคอลัมน์ช่วย VLOOKUP ด้านขวา (AQ..BC) ถ้ายังติดมา
        for rr_i in range(1, 20):
            for cc_i in range(43, 56):
                ws.cell(rr_i, cc_i).value = None

        ws["D5"] = year_int
        ws["AD16"] = year_int
        yy_cur = f"{year_int % 100:02d}"
        yy_prev = f"{(year_int - 1) % 100:02d}"
        ws["J8"] = f"Ave. '{yy_prev}"
        ws["K8"] = f"Target '{yy_cur}"
        ws["L8"] = f"Actual '{yy_cur}"
        ws["K9"] = f">{target_pct}%"
        if year_int == 2026 and (ws["J9"].value in (None, "", " -", "-")):
            ws["J9"] = DEFAULT_AVE_PREV.get(at)
            ws["J9"].number_format = "0.00"
        if meta.get("issued_by"):
            ws["J4"] = meta["issued_by"]
        if meta.get("reviewed_by"):
            ws["L4"] = meta["reviewed_by"]
        if meta.get("approved_mgr"):
            ws["P4"] = meta["approved_mgr"]
        if meta.get("approved_qmr"):
            ws["V4"] = meta["approved_qmr"]

        at_plans = plans.get(at, DEFAULT_KPI_PLANS[at])
        filled_months = 0
        tot_actual = 0
        tot_plan = 0

        for m in range(1, 13):
            col_idx = 29 + m  # AD = 30 .. AO = 41
            col_let = get_column_letter(col_idx)
            rec = by_month.get(m) if m <= max_m else None
            has_data = rec is not None and rec[f"{at}_db_docs"] > 0
            if has_data:
                actual = rec[f"{at}_in_master"] if use_master_only else rec[f"{at}_db_assets"]
                plan_v = int(at_plans[m - 1]) if m - 1 < len(at_plans) and at_plans[m - 1] else 0
                ws.cell(11, col_idx, int(actual))
                ws.cell(12, col_idx, plan_v if plan_v > 0 else None)
                ws.cell(14, col_idx, target_pct)
                ws.cell(15, col_idx, f'=IFERROR({col_let}11*100/{col_let}12,"")')
                for act_r in (19, 21, 23, 25):
                    ws.cell(act_r, col_idx, "•")
                filled_months += 1
                tot_actual += int(actual)
                tot_plan += plan_v
            else:
                ws.cell(11, col_idx, None)
                ws.cell(12, col_idx, None)
                ws.cell(14, col_idx, None)
                ws.cell(15, col_idx, None)
                for act_r in (19, 21, 23, 25):
                    ws.cell(act_r, col_idx, None)

        # คอลัมน์รวม/เฉลี่ย (AP = 42) และ Actual ปีปัจจุบัน (L9)
        ws["AP11"] = "=SUM(AD11:AO11)"
        ws["AP12"] = "=SUM(AD12:AO12)"
        ws["AP14"] = target_pct
        ws["AP15"] = '=IFERROR(AP11*100/AP12,0)'
        ws["L9"] = "=AP15"
        ws["L9"].number_format = "0.00"
        wb.calculation.fullCalcOnLoad = True
        summary[at] = dict(
            months=filled_months,
            actual=tot_actual,
            plan=tot_plan,
            achieve_pct=(tot_actual * 100.0 / tot_plan) if tot_plan > 0 else 0.0,
        )

    # ชีต 4: ทะเบียนทั้งหมด (ตาราง Yes / No PM รายเดือนของหัวลากใน Master)
    if include_plate_matrix:
        ws_m = wb.create_sheet("ทะเบียนทั้งหมด")
        con = connect_master(db_path)
        try:
            plates = [
                r[0]
                for r in con.execute(
                    "SELECT asset_code FROM asset_master WHERE asset_type='หัวลาก' AND is_active=1 ORDER BY sort_order, asset_code"
                ).fetchall()
            ]
            rep_pairs = {
                (r[0], r[1])
                for r in con.execute(
                    "SELECT DISTINCT asset_code, report_month FROM repair_request "
                    "WHERE asset_type='หัวลาก' AND in_iso='Yes' AND report_month LIKE ? AND report_month <= ?",
                    (f"{year_str}-%", f"{year_str}-{max_m:02d}"),
                ).fetchall()
            }
        finally:
            con.close()

        active_months = sorted(by_month.keys())
        yy = year_str[-2:]
        f_title = Font(name="Arial", size=11, bold=True)
        f_hdr = Font(name="Arial", size=10, bold=True)
        f_cell = Font(name="Arial", size=10)
        fill_hdr = PatternFill("solid", fgColor="DCE9F5")
        fill_yes = PatternFill("solid", fgColor="E2EFDA")
        fill_no = PatternFill("solid", fgColor="FCE4D6")
        fill_tot = PatternFill("solid", fgColor="F2F2F2")
        b_thin = _thin_border()

        ws_m["A2"] = f"เพิ่มทะเบียนทั้งหมด (หัวลาก Master ปี {year_str})"
        ws_m["A2"].font = f_title
        ws_m.cell(3, 1, "ทะเบียน").font = f_hdr
        ws_m.cell(3, 1).fill = fill_hdr
        ws_m.cell(3, 1).border = b_thin
        ws_m.cell(3, 1).alignment = Alignment(horizontal="center", vertical="center")
        ws_m.column_dimensions["A"].width = 14

        for j, m in enumerate(active_months, start=2):
            c = ws_m.cell(3, j, f"{m:02d}-{yy}")
            c.font = f_hdr; c.fill = fill_hdr; c.border = b_thin
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws_m.column_dimensions[get_column_letter(j)].width = 10

        for i, p in enumerate(plates, start=4):
            cp = ws_m.cell(i, 1, p)
            cp.font = f_cell; cp.border = b_thin
            cp.alignment = Alignment(horizontal="center", vertical="center")
            for j, m in enumerate(active_months, start=2):
                ym = f"{year_str}-{m:02d}"
                is_yes = (p, ym) in rep_pairs
                c = ws_m.cell(i, j, "Yes" if is_yes else "No PM")
                c.font = f_cell; c.border = b_thin
                c.fill = fill_yes if is_yes else fill_no
                c.alignment = Alignment(horizontal="center", vertical="center")

        tot_r = len(plates) + 4
        ct0 = ws_m.cell(tot_r, 1, f"รวม ({len(plates)})")
        ct0.font = f_hdr; ct0.fill = fill_tot; ct0.border = b_thin
        ct0.alignment = Alignment(horizontal="center", vertical="center")
        for j in range(2, len(active_months) + 2):
            col_let = get_column_letter(j)
            ct = ws_m.cell(tot_r, j, f'=COUNTIF({col_let}4:{col_let}{tot_r-1},"Yes")' if plates else 0)
            ct.font = f_hdr; ct.fill = fill_tot; ct.border = b_thin
            ct.alignment = Alignment(horizontal="center", vertical="center")
        ws_m.freeze_panes = "B4"

    return wb, summary


def export_kpi_iso_template(
    db_path,
    year,
    template_source,
    use_master_only=True,
    plans_by_type=None,
    target_pct=80,
    meta=None,
    include_plate_matrix=True,
    max_month=None,
):
    """เติมข้อมูล KPI รายเดือนลงในแบบฟอร์ม ISO FP-MR-00-01 Rev.00 (MTD-หัวลาก / MTD-หางพ่วง / MTD-ตู้เย็น + ทะเบียนทั้งหมด)
    คืน (excel_bytes, summary_dict)"""
    wb, summary = _build_iso_kpi_workbook(
        db_path=db_path,
        year=year,
        template_source=template_source,
        use_master_only=use_master_only,
        plans_by_type=plans_by_type,
        target_pct=target_pct,
        meta=meta,
        include_plate_matrix=include_plate_matrix,
        max_month=max_month,
    )
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), summary


def _logo_data_uri(logo_path):
    if not logo_path or not os.path.exists(logo_path):
        return None
    with open(logo_path, "rb") as fh:
        b64 = base64.b64encode(fh.read()).decode("ascii")
    return f"data:image/png;base64,{b64}"


# ============================================================================= Excel Ready-to-Print Builder
def _thin_border():
    s = Side(style="thin", color="808080")
    return Border(left=s, right=s, top=s, bottom=s)


def _header_border():
    s = Side(style="thin", color="333333")
    return Border(left=s, right=s, top=s, bottom=s)


def _add_form_sheet(wb, at, rep, data, logo_path=None):
    spec = data["spec"]
    ws = wb.create_sheet(spec["sheet_name"])
    ncols = len(spec["columns"])
    last_col_letter = get_column_letter(ncols)

    # Fonts & Styles
    f_comp_th = Font(name="Arial", size=14, bold=True)
    f_comp_en = Font(name="Arial", size=11, bold=True)
    f_comp_sub = Font(name="Arial", size=10)
    f_title = Font(name="Arial", size=13, bold=True)
    f_sub = Font(name="Arial", size=11)
    f_hdr = Font(name="Arial", size=10, bold=True)
    f_cell = Font(name="Arial", size=10)
    f_kpi_lbl = Font(name="Arial", size=10, bold=True)
    f_kpi_val = Font(name="Arial", size=10)
    fill_hdr = PatternFill("solid", fgColor="F2F2F2")
    fill_kpi = PatternFill("solid", fgColor="EEF5FC")
    fill_doc = PatternFill("solid", fgColor="FFFFFF")
    b_thin = _thin_border()
    b_hdr = _header_border()

    # Logo area A1:A4
    ws.merge_cells("A1:A4")
    if logo_path and os.path.exists(logo_path):
        try:
            from openpyxl.drawing.image import Image as XLImage
            img = XLImage(logo_path)
            img.width = 92
            img.height = 74
            ws.add_image(img, "A1")
        except Exception:
            pass

    # Company header B1:last_col
    ws["B1"] = COMPANY_INFO["name_th"]; ws["B1"].font = f_comp_th
    ws["B2"] = COMPANY_INFO["name_en"]; ws["B2"].font = f_comp_en
    ws["B3"] = COMPANY_INFO["address"]; ws["B3"].font = f_comp_sub
    ws["B4"] = COMPANY_INFO["tax_id"]; ws["B4"].font = f_comp_sub
    for r in range(1, 5):
        ws.row_dimensions[r].height = 19.5

    # Title & Date range (Rows 6-7)
    ws.merge_cells(f"A6:{last_col_letter}6")
    ws["A6"] = f"แผนซ่อมบำรุงรักษา - {spec['title_label']}"
    ws["A6"].font = f_title
    ws["A6"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[6].height = 24

    ws.merge_cells(f"A7:{last_col_letter}7")
    ws["A7"] = rep["date_range_text"]
    ws["A7"].font = f_sub
    ws["A7"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[7].height = 20

    ws.merge_cells(f"A8:{last_col_letter}8")
    ws.row_dimensions[8].height = 8

    # Table Headers (Row 9)
    ws.row_dimensions[9].height = 56
    for idx, (col_title, col_w) in enumerate(spec["columns"], start=1):
        c = ws.cell(9, idx, col_title)
        c.font = f_hdr
        c.fill = fill_hdr
        c.border = b_hdr
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(idx)].width = col_w

    # Helper column header (outside Print Area)
    hcol = spec["helper_col_idx"]
    hc = ws.cell(9, hcol, "นับจำนวนที่เข้าซ่อม")
    hc.font = f_hdr
    hc.fill = fill_kpi
    hc.border = b_thin
    hc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.column_dimensions[get_column_letter(hcol - 1)].width = 3.0
    ws.column_dimensions[get_column_letter(hcol)].width = 14.0

    # Data rows (Row 10+)
    doc_col = spec["doc_col_idx"]
    doc_col_letter = get_column_letter(doc_col)
    doc_col_width = spec["columns"][doc_col - 1][1]
    start_row = 10
    rows = data["rows"]
    for i, item in enumerate(rows):
        r = start_row + i
        for col_i in range(1, ncols + 1):
            cell = ws.cell(r, col_i)
            cell.font = f_cell
            cell.border = b_thin
            cell.fill = fill_doc
            if col_i == 1:
                cell.value = item["asset_code"]
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_i == doc_col:
                cell.value = item["doc_text"] if item["doc_text"] else None
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

        # Helper formula =IF(F10<>"",1,0)
        h_cell = ws.cell(r, hcol, f'=IF({doc_col_letter}{r}<>"",1,0)')
        h_cell.font = f_cell
        h_cell.alignment = Alignment(horizontal="center", vertical="center")

        # Auto row height for wrapped multi-doc cells
        if item["doc_text"]:
            chars_per_line = max(int(doc_col_width * 0.88), 16)
            lines = max(1, math.ceil(len(item["doc_text"]) / chars_per_line))
            ws.row_dimensions[r].height = max(19, lines * 15 + 4)
        else:
            ws.row_dimensions[r].height = 19

    last_row = max(start_row, start_row + len(rows) - 1)

    # Right-side KPI Box (outside Print Area) — exact formulas (fixes old COUNTA +2 bug!)
    kl = spec["kpi_label_col"]
    kv = spec["kpi_val_col"]
    h_letter = get_column_letter(hcol)
    kv_letter = get_column_letter(kv)
    ws.column_dimensions[get_column_letter(kl)].width = 18.0
    ws.column_dimensions[kv_letter].width = 16.0

    kpi_rows = [
        (3, "ประเภท", spec["title_label"], None),
        (4, "เดือน", rep["start_date"], "dd/mm/yyyy"),
        (5, "End of Month", rep["end_date"], "dd/mm/yyyy"),
        (6, "Total (ใน Master)", f"=COUNTA(A{start_row}:A{last_row})" if rows else 0, "#,##0"),
        (7, "มีงานซ่อม (ใน Master)", f"=SUM({h_letter}{start_row}:{h_letter}{last_row})" if rows else 0, "#,##0"),
        (8, "% การเข้าซ่อม", f"=IF({kv_letter}6>0,{kv_letter}7/{kv_letter}6,0)", "0.0%"),
    ]
    for rr_i, lbl, val, num_fmt in kpi_rows:
        c1 = ws.cell(rr_i, kl, lbl)
        c1.font = f_kpi_lbl; c1.fill = fill_kpi; c1.border = b_thin
        c2 = ws.cell(rr_i, kv, val)
        c2.font = f_kpi_val; c2.border = b_thin
        c2.alignment = Alignment(horizontal="right", vertical="center")
        if num_fmt:
            c2.number_format = num_fmt

    # Note if unmatched assets exist
    unm = data["unmatched"]
    if unm:
        c_u1 = ws.cell(10, kl, "นอก Master (รอตรวจ)")
        c_u1.font = Font(name="Arial", size=10, bold=True, color="9C0006")
        c_u1.fill = PatternFill("solid", fgColor="FFC7CE")
        c_u1.border = b_thin
        c_u2 = ws.cell(10, kv, len(unm))
        c_u2.font = Font(name="Arial", size=10, bold=True, color="9C0006")
        c_u2.fill = PatternFill("solid", fgColor="FFC7CE")
        c_u2.border = b_thin
        c_u2.alignment = Alignment(horizontal="right", vertical="center")

    # Page Setup — A4 Portrait, Fit to 1 page wide, Repeat Header Rows 1:9
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = 0.35
    ws.page_margins.right = 0.35
    ws.page_margins.top = 0.45
    ws.page_margins.bottom = 0.45
    ws.page_margins.header = 0.2
    ws.page_margins.footer = 0.2
    ws.print_title_rows = "1:9"
    ws.print_area = f"A1:{last_col_letter}{last_row}"
    ws.oddFooter.right.text = "หน้า &P / &N"
    ws.oddFooter.left.text = f"{spec['sheet_name']} · {rep['report_month']}"
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = "A10"


def _add_kpi_sheet(wb, rep, yearly_rows, logo_path=None, index=0):
    ws = wb.create_sheet("สรุป_KPI_ประจำเดือน", index) if index is not None else wb.create_sheet("สรุป_KPI_ประจำเดือน")
    f_comp = Font(name="Arial", size=14, bold=True)
    f_sub = Font(name="Arial", size=10)
    f_title = Font(name="Arial", size=13, bold=True)
    f_sec = Font(name="Arial", size=11, bold=True, color="1F4E78")
    f_hdr = Font(name="Arial", size=10, bold=True)
    f_cell = Font(name="Arial", size=10)
    f_bold = Font(name="Arial", size=10, bold=True)
    fill_hdr = PatternFill("solid", fgColor="DCE9F5")
    fill_tot = PatternFill("solid", fgColor="F2F2F2")
    fill_warn = PatternFill("solid", fgColor="FFF2CC")
    b_thin = _thin_border()

    ws.merge_cells("A1:A3")
    if logo_path and os.path.exists(logo_path):
        try:
            from openpyxl.drawing.image import Image as XLImage
            img = XLImage(logo_path)
            img.width = 76
            img.height = 62
            ws.add_image(img, "A1")
        except Exception:
            pass
    ws["B1"] = COMPANY_INFO["name_th"]; ws["B1"].font = f_comp
    ws["B2"] = COMPANY_INFO["name_en"]; ws["B2"].font = f_sub
    ws["B3"] = COMPANY_INFO["address"]; ws["B3"].font = f_sub

    ws.merge_cells("A5:H5")
    ws["A5"] = f"สรุป KPI การเข้าซ่อมประจำเดือน {rep['month_th']}"
    ws["A5"].font = f_title
    ws["A5"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 24

    ws.merge_cells("A6:H6")
    ws["A6"] = f"{rep['date_range_text']}   |   พิมพ์เมื่อ {dt.datetime.now():%d/%m/%Y %H:%M}"
    ws["A6"].font = f_sub
    ws["A6"].alignment = Alignment(horizontal="center", vertical="center")

    # Section 1: Selected Month KPI
    ws["A8"] = "1. สรุป KPI ประจำเดือนที่เลือก (อิงทะเบียนรถหลัก Master และแยกทะเบียนนอก Master ให้ DATA_OWNER ตรวจสอบ)"
    ws["A8"].font = f_sec
    hdr1 = ["ประเภท", "Total (ใน Master)", "เข้าซ่อม (ใน Master)", "% การเข้าซ่อม", "ใบแจ้งซ่อม (ใน Master)",
            "นอก Master (คัน/ตู้)", "รวมในฐาน DB (คัน/ตู้)", "รวมใบแจ้งซ่อม (ใบ)"]
    ws.row_dimensions[9].height = 30
    for j, h in enumerate(hdr1, 1):
        c = ws.cell(9, j, h)
        c.font = f_hdr; c.fill = fill_hdr; c.border = b_thin
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    r = 10
    sum_m = sum_rep = sum_dm = sum_unm = sum_dba = sum_dbd = 0
    for at in TYPES:
        k = rep["by_type"][at]["kpi"]
        vals = [at, k["total_master"], k["repaired_in_master"], k["pct_repaired"], k["docs_in_master"],
                k["unmatched_assets"], k["total_db_assets"], k["total_db_docs"]]
        sum_m += k["total_master"]; sum_rep += k["repaired_in_master"]; sum_dm += k["docs_in_master"]
        sum_unm += k["unmatched_assets"]; sum_dba += k["total_db_assets"]; sum_dbd += k["total_db_docs"]
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v)
            c.font = f_cell; c.border = b_thin
            if j == 1:
                c.alignment = Alignment(horizontal="left", vertical="center")
            elif j == 4:
                c.number_format = "0.0%"; c.alignment = Alignment(horizontal="right", vertical="center")
            else:
                c.number_format = "#,##0"; c.alignment = Alignment(horizontal="right", vertical="center")
                if j == 6 and v > 0:
                    c.fill = fill_warn
        ws.row_dimensions[r].height = 21
        r += 1

    # Total row
    tot_vals = ["รวมทั้งหมด", sum_m, sum_rep, (sum_rep / sum_m if sum_m else 0.0), sum_dm, sum_unm, sum_dba, sum_dbd]
    for j, v in enumerate(tot_vals, 1):
        c = ws.cell(r, j, v)
        c.font = f_bold; c.fill = fill_tot; c.border = b_thin
        if j == 1:
            c.alignment = Alignment(horizontal="left", vertical="center")
        elif j == 4:
            c.number_format = "0.0%"; c.alignment = Alignment(horizontal="right", vertical="center")
        else:
            c.number_format = "#,##0"; c.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[r].height = 22

    # Section 2: Yearly monthly breakdown
    r += 2
    ws.cell(r, 1, "2. ตารางเปรียบเทียบ KPI รายเดือน (ใน Master / % เข้าซ่อม / นอก Master)").font = f_sec
    r += 1
    hdr2 = ["เดือน",
            "หัวลาก (เข้าซ่อม)", "หัวลาก (%)", "หัวลาก (นอก Master)",
            "หางพ่วง (เข้าซ่อม)", "หางพ่วง (%)", "หางพ่วง (นอก Master)",
            "ตู้เย็น (เข้าซ่อม)", "ตู้เย็น (%)", "ตู้เย็น (นอก Master)"]
    ws.row_dimensions[r].height = 30
    for j, h in enumerate(hdr2, 1):
        c = ws.cell(r, j, h)
        c.font = f_hdr; c.fill = fill_hdr; c.border = b_thin
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    r += 1
    for yr in yearly_rows:
        row_vals = [
            yr["report_month"],
            yr["หัวลาก_in_master"], yr["หัวลาก_pct"], yr["หัวลาก_unmatched"],
            yr["หางพ่วง_in_master"], yr["หางพ่วง_pct"], yr["หางพ่วง_unmatched"],
            yr["ตู้เย็น_in_master"], yr["ตู้เย็น_pct"], yr["ตู้เย็น_unmatched"],
        ]
        is_sel = (yr["report_month"] == rep["report_month"])
        for j, v in enumerate(row_vals, 1):
            c = ws.cell(r, j, v)
            c.font = f_bold if is_sel else f_cell
            c.border = b_thin
            if is_sel:
                c.fill = PatternFill("solid", fgColor="E8F4FD")
            if j == 1:
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif j in (3, 6, 9):
                c.number_format = "0.0%"; c.alignment = Alignment(horizontal="right", vertical="center")
            else:
                c.number_format = "#,##0"; c.alignment = Alignment(horizontal="right", vertical="center")
        ws.row_dimensions[r].height = 20
        r += 1

    widths = [15, 16, 15, 16, 16, 15, 16, 16, 15, 16]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_area = f"A1:J{max(r - 1, 13)}"
    ws.oddFooter.right.text = "หน้า &P / &N"


def _add_unmatched_sheet(wb, rep):
    all_unm = []
    for at in TYPES:
        all_unm.extend(rep["by_type"][at]["unmatched"])
    if not all_unm:
        return
    ws = wb.create_sheet("นอก_Master_รอตรวจ")
    f_hdr = Font(name="Arial", size=10, bold=True)
    f_cell = Font(name="Arial", size=10)
    fill_hdr = PatternFill("solid", fgColor="FCE4D6")
    b_thin = _thin_border()
    ws["A1"] = f"รายการที่มีใบแจ้งซ่อมเดือน {rep['month_th']} แต่ไม่อยู่ในทะเบียนรถหลัก (Master) — รอ DATA_OWNER ตรวจสอบ"
    ws["A1"].font = Font(name="Arial", size=12, bold=True, color="9C0006")
    headers = ["ประเภท", "ทะเบียน/เบอร์ตู้", "ทะเบียนต้นฉบับ", "จำนวนใบซ่อม", "เลขที่ใบแจ้งซ่อม", "วันที่แรก", "วันที่ล่าสุด", "ยี่ห้อ", "สถานะ", "หมายเหตุ"]
    widths = [12, 18, 32, 13, 38, 13, 13, 14, 16, 26]
    for j, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(3, j, h)
        c.font = f_hdr; c.fill = fill_hdr; c.border = b_thin
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(j)].width = w
    for i, x in enumerate(all_unm, start=4):
        note = "อยู่ใน Master (สถานะปิดใช้งาน)" if x["in_master_inactive"] else "ไม่อยู่ใน Master (อาจไม่ใช่รถในกลุ่ม/รถใหม่)"
        vals = [x["asset_type"], x["asset_code"], x["asset_raw"], x["repair_count"], x["doc_text"],
                x["first_date"], x["last_date"], x["brands"], x["statuses"], note]
        for j, v in enumerate(vals, 1):
            c = ws.cell(i, j, v)
            c.font = f_cell; c.border = b_thin
            c.alignment = Alignment(horizontal="center" if j in (1, 2, 4, 6, 7) else "left", vertical="center", wrap_text=(j == 5))
    ws.freeze_panes = "A4"
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def build_monthly_excel(db_path, report_month, logo_path=None, include_kpi=True, include_types=None):
    """สร้างไฟล์ Excel (.xlsx) พร้อมพิมพ์ (คืน bytes)"""
    rep = compute_monthly_report(db_path, report_month)
    year = report_month[:4]
    yearly = compute_yearly_kpi(db_path, year=year) if include_kpi else []
    selected_types = include_types if include_types is not None else TYPES

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    if include_kpi:
        _add_kpi_sheet(wb, rep, yearly, logo_path=logo_path)
    for at in selected_types:
        if at in rep["by_type"]:
            _add_form_sheet(wb, at, rep, rep["by_type"][at], logo_path=logo_path)
    _add_unmatched_sheet(wb, rep)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), rep


def _add_last_pm_sheet(wb, db_path, report_month, pm_interval_days=180):
    """เพิ่มชีต 'Last PM' ตาม Template บริษัท:
    - หัวลาก Master เรียงตามลำดับ: (No.), (M/C CODE), วันที่เข้าศูนย์, วันที่เข้าศูนย์ล่าสุด,
      แพลนเข้าศูนย์ (+180 วัน), สถานะ (on plan / overdue), แผน pm (mmm-yy), note, ISO (YES)
    - ตารางสรุปด้านล่าง 18 เดือน (JUL-YY .. DEC-YY+1): PM PLAN (=COUNTIF) และ เข้าบริการแล้ว (จำนวนคันจริงใน Master)"""
    year_int = int(report_month.split("-")[0])
    today = dt.date.today()

    con = connect_master(db_path)
    try:
        masters = con.execute(
            "SELECT asset_code, note FROM asset_master "
            "WHERE asset_type='หัวลาก' AND is_active=1 ORDER BY sort_order, asset_code"
        ).fetchall()
        last_dates = {
            r[0]: r[1]
            for r in con.execute(
                "SELECT asset_code, MAX(report_date) FROM repair_request "
                "WHERE asset_type='หัวลาก' AND in_iso='Yes' AND report_month <= ? "
                "GROUP BY asset_code",
                (report_month,),
            ).fetchall()
        }
        master_set = {m[0] for m in masters}
        monthly_serviced = {}
        for r in con.execute(
            "SELECT report_month, asset_code FROM repair_request "
            "WHERE asset_type='หัวลาก' AND in_iso='Yes' AND report_month <= ? "
            "GROUP BY report_month, asset_code",
            (report_month,),
        ).fetchall():
            if r[1] in master_set:
                monthly_serviced[r[0]] = monthly_serviced.get(r[0], 0) + 1
    finally:
        con.close()

    ws = wb.create_sheet("Last PM")
    f_hdr = Font(name="Arial", size=10, bold=True)
    f_cell = Font(name="Arial", size=10)
    f_bold = Font(name="Arial", size=10, bold=True)
    fill_hdr = PatternFill("solid", fgColor="DCE9F5")
    fill_ok = PatternFill("solid", fgColor="E2EFDA")
    fill_over = PatternFill("solid", fgColor="FFC7CE")
    fill_sum = PatternFill("solid", fgColor="F2F2F2")
    b_thin = _thin_border()

    headers = [
        "(No.)", "(M/C CODE)", "วันที่เข้าศูนย์", "วันที่เข้าศูนย์ล่าสุด",
        "แพลนเข้าศูนย์", "สถานะ", "แผน pm", "note", "ISO",
    ]
    widths = [8, 14, 15, 17, 15, 12, 11, 16, 8]
    ws.row_dimensions[1].height = 24
    for j, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(1, j, h)
        c.font = f_hdr; c.fill = fill_hdr; c.border = b_thin
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = w

    for idx, (code, db_note) in enumerate(masters, start=1):
        r = idx + 1
        d_str = last_dates.get(code)
        d_val = dt.date.fromisoformat(d_str) if d_str else None
        plan_date = (d_val + dt.timedelta(days=pm_interval_days)) if d_val else None
        is_overdue = (plan_date is not None and plan_date < today)

        note_txt = None
        if db_note and "คืนรถ" in db_note:
            note_txt = db_note
        elif code in LAST_PM_DEFAULT_NOTES:
            note_txt = LAST_PM_DEFAULT_NOTES[code]

        row_cells = [
            (1, idx, None, "center"),
            (2, code, None, "center"),
            (3, d_val, "yyyy-mm-dd", "center"),
            (4, f'=IF(C{r}="","",C{r})' if d_val else None, "yyyy-mm-dd", "center"),
            (5, f'=IF(D{r}="","",D{r}+{pm_interval_days})' if d_val else None, "yyyy-mm-dd", "center"),
            (6, f'=IF(E{r}="","",IF(E{r}<TODAY(),"overdue","on plan"))' if d_val else None, None, "center"),
            (7, f'=IF(E{r}="","",TEXT(E{r},"mmm-yy"))' if d_val else None, None, "center"),
            (8, note_txt, None, "left"),
            (9, "YES", None, "center"),
        ]
        for col_i, val, nfmt, align_h in row_cells:
            c = ws.cell(r, col_i, val)
            c.font = f_cell; c.border = b_thin
            c.alignment = Alignment(horizontal=align_h, vertical="center")
            if nfmt:
                c.number_format = nfmt
            if col_i == 6 and d_val:
                c.fill = fill_over if is_overdue else fill_ok
        ws.row_dimensions[r].height = 19

    # ตารางสรุป 18 เดือน (JUL-<YY> .. DEC-<YY+1>) ใต้รายการรถ
    last_r = max(2, len(masters) + 1)
    r_hdr = last_r + 1
    r_plan = r_hdr + 1
    r_act = r_hdr + 2

    span_months = [(year_int, m) for m in range(7, 13)] + [(year_int + 1, m) for m in range(1, 13)]
    ws.cell(r_plan, 1, "PM PLAN").font = f_bold
    ws.cell(r_act, 1, "เข้าบริการแล้ว").font = f_bold
    for col_offset, (yy, mm) in enumerate(span_months):
        col_i = 5 + col_offset  # E = 5 .. V = 22
        col_let = get_column_letter(col_i)
        lbl = f"{_EN_MONTH_UPPER[mm - 1]}-{str(yy)[-2:]}"
        ym_key = f"{yy:04d}-{mm:02d}"

        ch = ws.cell(r_hdr, col_i, lbl)
        ch.font = f_hdr; ch.fill = fill_hdr; ch.border = b_thin
        ch.alignment = Alignment(horizontal="center", vertical="center")

        cp = ws.cell(r_plan, col_i, f'=COUNTIF($G$2:$G${last_r},{col_let}{r_hdr})')
        cp.font = f_bold; cp.fill = fill_sum; cp.border = b_thin
        cp.number_format = "#,##0"
        cp.alignment = Alignment(horizontal="center", vertical="center")

        ca = ws.cell(r_act, col_i, int(monthly_serviced.get(ym_key, 0)))
        ca.font = f_cell; ca.border = b_thin
        ca.number_format = "#,##0"
        ca.alignment = Alignment(horizontal="center", vertical="center")

        if col_i > len(widths):
            ws.column_dimensions[col_let].width = 10

    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:I{last_r}"
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    return len(masters)


def _add_index_sheet(wb, report_month, rep, iso_summary, m_counts, last_pm_rows, logo_path=None):
    """เพิ่มชีต 'Index' ไว้หน้าสุดของไฟล์ KPI อธิบายข้อมูลและวัตถุประสงค์ของทุกชีตพร้อม Hyperlink คลิกข้ามไปแต่ละชีต"""
    ws = wb.create_sheet("Index", 0)
    f_comp = Font(name="Arial", size=14, bold=True)
    f_sub = Font(name="Arial", size=10)
    f_title = Font(name="Arial", size=13, bold=True, color="1F4E78")
    f_sec = Font(name="Arial", size=11, bold=True, color="1F4E78")
    f_hdr = Font(name="Arial", size=10, bold=True)
    f_cell = Font(name="Arial", size=10)
    f_link = Font(name="Arial", size=10, bold=True, color="0563C1", underline="single")
    f_note = Font(name="Arial", size=9.5, color="333333")
    fill_hdr = PatternFill("solid", fgColor="DCE9F5")
    fill_alt = PatternFill("solid", fgColor="F7FAFC")
    fill_note = PatternFill("solid", fgColor="FFF9E6")
    b_thin = _thin_border()

    ws.merge_cells("A1:A3")
    if logo_path and os.path.exists(logo_path):
        try:
            from openpyxl.drawing.image import Image as XLImage
            img = XLImage(logo_path)
            img.width = 76
            img.height = 62
            ws.add_image(img, "A1")
        except Exception:
            pass
    ws["B1"] = COMPANY_INFO["name_th"]; ws["B1"].font = f_comp
    ws["B2"] = COMPANY_INFO["name_en"]; ws["B2"].font = f_sub
    ws["B3"] = COMPANY_INFO["address"]; ws["B3"].font = f_sub

    ws.merge_cells("A5:E5")
    ws["A5"] = f"สารบัญรายงาน KPI และการบำรุงรักษาประจำเดือน {rep['month_th']} ({report_month})"
    ws["A5"].font = f_title
    ws["A5"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 24

    ws.merge_cells("A6:E6")
    ws["A6"] = (
        f"{rep['date_range_text']}   |   เงื่อนไข: เฉพาะทะเบียนใน Master / ISO = Yes (ตัด ISO = No ออกทุกชีต)   |   "
        f"สร้างเมื่อ {dt.datetime.now():%d/%m/%Y %H:%M}"
    )
    ws["A6"].font = f_sub
    ws["A6"].alignment = Alignment(horizontal="center", vertical="center")

    headers = ["ลำดับ", "ชื่อชีต (คลิกเพื่อเปิด)", "กลุ่มรายงาน", "คำอธิบายข้อมูลและวัตถุประสงค์", "สรุปข้อมูลในชีต"]
    widths = [8, 24, 20, 62, 34]
    ws.row_dimensions[8].height = 26
    for j, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(8, j, h)
        c.font = f_hdr; c.fill = fill_hdr; c.border = b_thin
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = w

    def _iso_stat(at):
        s = iso_summary.get(at)
        if not s:
            return "-"
        return f"{s['months']} เดือน · จริง {s['actual']:,} / แผน {s['plan']:,} ({s['achieve_pct']:.1f}%)"

    def _m_stat(at):
        k = rep["by_type"][at]["kpi"]
        return f"เข้าซ่อม {k['repaired_in_master']:,} / {k['total_master']:,} ({k['pct_repaired']*100:.1f}%)"

    k_all_m = sum(rep["by_type"][at]["kpi"]["total_master"] for at in TYPES)
    k_all_r = sum(rep["by_type"][at]["kpi"]["repaired_in_master"] for at in TYPES)

    sheet_rows = [
        ("MTD-หัวลาก", "แบบฟอร์ม KPI (ISO)",
         "รายงาน Objective Action Plan & Evaluation (FP-MR-00-01 Rev.00) สำหรับหัวลาก พร้อมกราฟและสูตรเปรียบเทียบแผน-ผลจริง (Jan ถึงเดือนที่เลือก)",
         _iso_stat("หัวลาก")),
        ("MTD-หางพ่วง", "แบบฟอร์ม KPI (ISO)",
         "รายงาน Objective Action Plan & Evaluation (FP-MR-00-01 Rev.00) สำหรับหางพ่วง พร้อมกราฟและสูตรเปรียบเทียบแผน-ผลจริง (Jan ถึงเดือนที่เลือก)",
         _iso_stat("หางพ่วง")),
        ("MTD-ตู้เย็น", "แบบฟอร์ม KPI (ISO)",
         "รายงาน Objective Action Plan & Evaluation (FP-MR-00-01 Rev.00) สำหรับตู้เย็น พร้อมกราฟและสูตรเปรียบเทียบแผน-ผลจริง (Jan ถึงเดือนที่เลือก)",
         _iso_stat("ตู้เย็น")),
        ("ทะเบียนทั้งหมด", "ติดตามรายคัน",
         "ตารางติดตามการเข้าซ่อม/PM รายเดือน (Yes / No PM) ของหัวลากใน Master ทุกคัน ตั้งแต่ต้นปีถึงเดือนที่เลือก พร้อมแถวรวม COUNTIF",
         f"หัวลาก Master {rep['by_type']['หัวลาก']['kpi']['total_master']:,} คัน"),
        ("สรุป_KPI_ประจำเดือน", "สรุปผู้บริหาร",
         "สรุป KPI การเข้าซ่อมประจำเดือนที่เลือก แยกตามประเภท (หัวลาก/หางพ่วง/ตู้เย็น) และตารางเปรียบเทียบ KPI รายเดือนตั้งแต่ต้นปีถึงเดือนที่เลือก",
         f"รวมเข้าซ่อม {k_all_r:,} / {k_all_m:,} รายการ"),
        ("1_หัวลาก", "แบบฟอร์มประจำเดือน",
         "แบบฟอร์มรายงานการบำรุงรักษาประจำเดือนสำหรับหัวลาก แสดงเลขที่ใบแจ้งซ่อมรายคัน (จัดหน้า A4 แนวตั้งพร้อมพิมพ์)",
         _m_stat("หัวลาก")),
        ("2_หางพ่วง", "แบบฟอร์มประจำเดือน",
         "แบบฟอร์มรายงานการบำรุงรักษาประจำเดือนสำหรับหางพ่วง แสดงเลขที่ใบแจ้งซ่อมรายคัน (จัดหน้า A4 แนวตั้งพร้อมพิมพ์)",
         _m_stat("หางพ่วง")),
        ("3_ซ่อมตู้", "แบบฟอร์มประจำเดือน",
         "แบบฟอร์มรายงานการบำรุงรักษาประจำเดือนสำหรับตู้เย็น แสดงเลขที่ใบแจ้งซ่อมรายตู้ (จัดหน้า A4 แนวตั้งพร้อมพิมพ์)",
         _m_stat("ตู้เย็น")),
        ("M1_ซ่อมรถ", "ข้อมูลดิบ (Raw Data)",
         "รายละเอียดใบแจ้งซ่อมหัวลากสะสมตั้งแต่ต้นปีถึงเดือนที่เลือก (1 แถว = 1 เลขที่เอกสาร, เฉพาะทะเบียนใน Master / ISO = Yes)",
         f"{m_counts.get('M1_ซ่อมรถ', 0):,} รายการ (ISO = Yes)"),
        ("M2_ซ่อมหาง", "ข้อมูลดิบ (Raw Data)",
         "รายละเอียดใบแจ้งซ่อมหางพ่วงสะสมตั้งแต่ต้นปีถึงเดือนที่เลือก (1 แถว = 1 เลขที่เอกสาร, เฉพาะ ISO = Yes)",
         f"{m_counts.get('M2_ซ่อมหาง', 0):,} รายการ"),
        ("M3_ซ่อมตู้", "ข้อมูลดิบ (Raw Data)",
         "รายละเอียดใบแจ้งซ่อมตู้เย็นสะสมตั้งแต่ต้นปีถึงเดือนที่เลือก (1 แถว = 1 เลขที่เอกสาร, เฉพาะ ISO = Yes)",
         f"{m_counts.get('M3_ซ่อมตู้', 0):,} รายการ"),
        ("Last PM", "แผน PM รอบถัดไป",
         "ติดตามวันที่เข้าศูนย์ล่าสุดของหัวลาก Master, คำนวณแพลนเข้าศูนย์รอบถัดไป (+180 วัน), สถานะ on plan/overdue และสรุปแผน PM 18 เดือน",
         f"หัวลาก {last_pm_rows:,} คัน (+สรุป 18 เดือน)"),
    ]

    r = 9
    for idx, (sname, grp, desc, stat) in enumerate(sheet_rows, start=1):
        row_fill = fill_alt if idx % 2 == 0 else None
        vals = [idx, sname, grp, desc, stat]
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v)
            c.border = b_thin
            if row_fill:
                c.fill = row_fill
            if j == 1:
                c.font = f_cell
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif j == 2:
                c.font = f_link
                c.hyperlink = f"#'{sname}'!A1"
                c.alignment = Alignment(horizontal="left", vertical="center")
            elif j == 3:
                c.font = f_cell
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.font = f_cell
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 28
        r += 1

    # หมายเหตุ / กติกาข้อมูล
    r += 1
    ws.cell(r, 1, "หมายเหตุและกติกาข้อมูล").font = f_sec
    r += 1
    notes = [
        f"1. ฐานทะเบียนอุปกรณ์หลัก (Active Master): หัวลาก {rep['by_type']['หัวลาก']['kpi']['total_master']:,} คัน · "
        f"หางพ่วง {rep['by_type']['หางพ่วง']['kpi']['total_master']:,} หาง · "
        f"ตู้เย็น {rep['by_type']['ตู้เย็น']['kpi']['total_master']:,} ตู้ (รวม {k_all_m:,} รายการ)",
        "2. ข้อมูลทุกชีตคัดกรองเฉพาะรายการที่ ISO = Yes เท่านั้น (ตัดรายการ ISO = No ออกทั้งหมดอัตโนมัติ)",
        "3. ชีต MTD-*, ทะเบียนทั้งหมด, M1..M3 และ Last PM แสดงข้อมูลสะสมตั้งแต่ต้นปีถึงเดือนที่เลือก และคงสูตร Excel ไว้ครบถ้วน",
        "4. สามารถคลิกที่ชื่อชีตในคอลัมน์ B เพื่อกระโดดไปยังชีตที่ต้องการได้ทันที",
    ]
    for nt in notes:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        cn = ws.cell(r, 1, nt)
        cn.font = f_note; cn.fill = fill_note
        cn.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[r].height = 20
        r += 1

    ws.freeze_panes = "A9"
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.print_area = f"A1:E{r - 1}"
    wb.active = 0


def build_kpi_monthly_package(
    db_path,
    report_month,
    template_source,
    logo_path=None,
    use_master_only=True,
    plans_by_type=None,
    target_pct=80,
    meta=None,
):
    """สร้างไฟล์ Excel ชุดรายงาน KPI ประจำเดือนครบ 13 ชีตตาม Template บริษัท:
      1.    Index (สารบัญและคำอธิบายข้อมูลทุกชีต พร้อม Hyperlink)
      2-4.  MTD-หัวลาก / MTD-หางพ่วง / MTD-ตู้เย็น (ISO FP-MR-00-01 Rev.00 ถึงเดือนที่เลือก)
      5.    ทะเบียนทั้งหมด (หัวลาก Master x Yes/No PM ถึงเดือนที่เลือก)
      6.    สรุป_KPI_ประจำเดือน (สรุปเดือนที่เลือก + ตารางเปรียบเทียบรายเดือนถึงเดือนที่เลือก)
      7-9.  1_หัวลาก / 2_หางพ่วง / 3_ซ่อมตู้ (แบบฟอร์มรายงานซ่อมประจำเดือนที่เลือก)
      10-12. M1_ซ่อมรถ / M2_ซ่อมหาง / M3_ซ่อมตู้ (รายละเอียดใบแจ้งซ่อม YTD ถึงเดือนที่เลือก — เฉพาะ ISO = Yes เท่านั้น)
      13.   Last PM (วันที่เข้าศูนย์ล่าสุด + แพลนเข้าศูนย์ +180 วัน + สถานะ on plan/overdue + ตารางสรุป PM PLAN 18 เดือน)
    คืน (excel_bytes, info_dict)"""
    year_str, m_str = report_month.split("-")
    sel_m = int(m_str)

    # MTD-หัวลาก / MTD-หางพ่วง / MTD-ตู้เย็น + ทะเบียนทั้งหมด
    wb, iso_summary = _build_iso_kpi_workbook(
        db_path=db_path,
        year=year_str,
        template_source=template_source,
        use_master_only=use_master_only,
        plans_by_type=plans_by_type,
        target_pct=target_pct,
        meta=meta,
        include_plate_matrix=True,
        max_month=sel_m,
    )

    # สรุป_KPI_ประจำเดือน + 1_หัวลาก / 2_หางพ่วง / 3_ซ่อมตู้
    rep = compute_monthly_report(db_path, report_month)
    yearly = [y for y in compute_yearly_kpi(db_path, year=year_str) if y["report_month"] <= report_month]
    _add_kpi_sheet(wb, rep, yearly, logo_path=logo_path, index=None)
    for at in TYPES:
        if at in rep["by_type"]:
            _add_form_sheet(wb, at, rep, rep["by_type"][at], logo_path=logo_path)

    # M1_ซ่อมรถ / M2_ซ่อมหาง / M3_ซ่อมตู้ (เฉพาะ ISO = Yes, YTD ถึงเดือนที่เลือก)
    con = rr.connect(db_path)
    try:
        m_counts = rr.append_m_sheets(wb, con, year=year_str, max_month=sel_m)
    finally:
        con.close()

    # Last PM (หัวลาก Master + แพลนเข้าศูนย์ +180 วัน + ตารางสรุป 18 เดือน)
    last_pm_rows = _add_last_pm_sheet(wb, db_path, report_month)

    # Index (สารบัญไว้หน้าสุด พร้อมคำอธิบายและ Hyperlink ทุกชีต)
    _add_index_sheet(wb, report_month, rep, iso_summary, m_counts, last_pm_rows, logo_path=logo_path)

    buf = io.BytesIO()
    wb.save(buf)
    info = dict(
        report_month=report_month,
        month_th=rep["month_th"],
        sheets=list(wb.sheetnames),
        iso_summary=iso_summary,
        m_counts=m_counts,
        last_pm_rows=last_pm_rows,
        rep=rep,
    )
    return buf.getvalue(), info


# ============================================================================= HTML Print View Builder (Ctrl+P / Save PDF)
PRINT_CSS = """
* { box-sizing: border-box; }
body { font-family: "Sarabun", "Tahoma", "Angsana New", Arial, sans-serif; color: #111; margin: 0; padding: 14px; background: #f4f6f8; font-size: 12px; }
.toolbar { position: sticky; top: 0; z-index: 99; background: #1e3a5f; color: #fff; padding: 10px 16px; border-radius: 8px; margin-bottom: 14px; display: flex; align-items: center; justify-content: space-between; gap: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); }
.toolbar button { background: #fff; color: #1e3a5f; border: none; padding: 8px 16px; border-radius: 6px; font-weight: 700; cursor: pointer; font-size: 13px; }
.toolbar button:hover { background: #e8f1fa; }
.sheet { background: #fff; padding: 14mm 10mm; margin: 0 auto 18px auto; max-width: 210mm; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
.sheet.landscape { max-width: 297mm; }
.page-break { page-break-before: always; break-before: page; }
.co-hdr { display: flex; align-items: center; gap: 12px; margin-bottom: 6px; }
.co-logo { width: 68px; height: auto; flex-shrink: 0; }
.co-text .th { font-size: 15px; font-weight: 700; }
.co-text .en { font-size: 11px; font-weight: 700; color: #333; }
.co-text .addr { font-size: 10.5px; color: #444; }
.doc-title { text-align: center; font-size: 15px; font-weight: 700; margin: 6px 0 2px 0; }
.doc-date { text-align: center; font-size: 12px; margin-bottom: 8px; color: #222; }
.kpi-strip { display: flex; gap: 10px; justify-content: flex-end; margin-bottom: 6px; font-size: 11px; flex-wrap: wrap; }
.kpi-pill { background: #eef5fc; border: 1px solid #c5dcf2; padding: 3px 8px; border-radius: 4px; }
.kpi-pill.warn { background: #fff3cd; border-color: #ffe08a; color: #7a4f01; }
table.rpt { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 11px; }
table.rpt thead { display: table-header-group; }
table.rpt tr { page-break-inside: avoid; break-inside: avoid; }
table.rpt th, table.rpt td { border: 1px solid #555; padding: 4px 5px; vertical-align: middle; word-wrap: break-word; overflow-wrap: break-word; }
table.rpt th { background: #f2f2f2; font-weight: 700; text-align: center; line-height: 1.25; }
table.rpt td.c { text-align: center; }
table.rpt td.r { text-align: right; }
table.rpt td.doc { text-align: left; font-size: 10.5px; }
table.rpt tr.tot td { background: #f2f2f2; font-weight: 700; }
table.rpt tr.hl td { background: #e8f4fd; font-weight: 700; }
.repeat-hdr-cell { border: none !important; background: #fff !important; padding: 0 0 6px 0 !important; text-align: left !important; }
.sec-title { font-size: 12.5px; font-weight: 700; color: #1f4e78; margin: 12px 0 6px 0; }
@media print {
  @page { size: A4 portrait; margin: 8mm; }
  body { background: #fff; padding: 0; }
  .no-print { display: none !important; }
  .sheet { box-shadow: none; margin: 0; padding: 0; max-width: 100%; }
}
"""


def _co_header_html(logo_uri, title, subtitle):
    logo_tag = f'<img class="co-logo" src="{logo_uri}" alt="logo"/>' if logo_uri else ""
    return (
        f'<div class="co-hdr">{logo_tag}'
        f'<div class="co-text">'
        f'<div class="th">{html.escape(COMPANY_INFO["name_th"])}</div>'
        f'<div class="en">{html.escape(COMPANY_INFO["name_en"])}</div>'
        f'<div class="addr">{html.escape(COMPANY_INFO["address"])} · {html.escape(COMPANY_INFO["tax_id"])}</div>'
        f'</div></div>'
        f'<div class="doc-title">{html.escape(title)}</div>'
        f'<div class="doc-date">{html.escape(subtitle)}</div>'
    )


def _kpi_section_html(rep, yearly_rows, logo_uri):
    hdr = _co_header_html(
        logo_uri,
        f"สรุป KPI การเข้าซ่อมประจำเดือน {rep['month_th']}",
        f"{rep['date_range_text']} · พิมพ์เมื่อ {dt.datetime.now():%d/%m/%Y %H:%M}",
    )
    rows_html = []
    sm = sr = sdm = sunm = sdba = sdbd = 0
    for at in TYPES:
        k = rep["by_type"][at]["kpi"]
        sm += k["total_master"]; sr += k["repaired_in_master"]; sdm += k["docs_in_master"]
        sunm += k["unmatched_assets"]; sdba += k["total_db_assets"]; sdbd += k["total_db_docs"]
        rows_html.append(
            f"<tr><td>{html.escape(at)}</td>"
            f"<td class='r'>{k['total_master']:,}</td>"
            f"<td class='r'>{k['repaired_in_master']:,}</td>"
            f"<td class='r'>{k['pct_repaired']*100:.1f}%</td>"
            f"<td class='r'>{k['docs_in_master']:,}</td>"
            f"<td class='r'>{k['unmatched_assets']:,}</td>"
            f"<td class='r'>{k['total_db_assets']:,}</td>"
            f"<td class='r'>{k['total_db_docs']:,}</td></tr>"
        )
    tot_pct = (sr / sm * 100) if sm else 0.0
    rows_html.append(
        f"<tr class='tot'><td>รวมทั้งหมด</td>"
        f"<td class='r'>{sm:,}</td><td class='r'>{sr:,}</td><td class='r'>{tot_pct:.1f}%</td>"
        f"<td class='r'>{sdm:,}</td><td class='r'>{sunm:,}</td><td class='r'>{sdba:,}</td><td class='r'>{sdbd:,}</td></tr>"
    )

    yr_html = []
    for y in yearly_rows:
        cls = " class='hl'" if y["report_month"] == rep["report_month"] else ""
        yr_html.append(
            f"<tr{cls}><td class='c'>{html.escape(y['report_month'])}</td>"
            f"<td class='r'>{y['หัวลาก_in_master']:,}</td><td class='r'>{y['หัวลาก_pct']*100:.1f}%</td><td class='r'>{y['หัวลาก_unmatched']:,}</td>"
            f"<td class='r'>{y['หางพ่วง_in_master']:,}</td><td class='r'>{y['หางพ่วง_pct']*100:.1f}%</td><td class='r'>{y['หางพ่วง_unmatched']:,}</td>"
            f"<td class='r'>{y['ตู้เย็น_in_master']:,}</td><td class='r'>{y['ตู้เย็น_pct']*100:.1f}%</td><td class='r'>{y['ตู้เย็น_unmatched']:,}</td></tr>"
        )

    return (
        f'<div class="sheet">{hdr}'
        f'<div class="sec-title">1. สรุป KPI ประจำเดือนที่เลือก</div>'
        f'<table class="rpt"><thead><tr>'
        f'<th>ประเภท</th><th>Total (Master)</th><th>เข้าซ่อม (Master)</th><th>% เข้าซ่อม</th>'
        f'<th>ใบซ่อม (Master)</th><th>นอก Master (รอตรวจ)</th><th>รวม DB (คัน/ตู้)</th><th>รวมใบซ่อม</th>'
        f'</tr></thead><tbody>{"".join(rows_html)}</tbody></table>'
        f'<div class="sec-title">2. ตารางเปรียบเทียบ KPI รายเดือน (ปี {rep["report_month"][:4]})</div>'
        f'<table class="rpt"><thead><tr>'
        f'<th>เดือน</th><th>หัวลาก (คัน)</th><th>%</th><th>นอก Master</th>'
        f'<th>หางพ่วง (คัน)</th><th>%</th><th>นอก Master</th>'
        f'<th>ตู้เย็น (ตู้)</th><th>%</th><th>นอก Master</th>'
        f'</tr></thead><tbody>{"".join(yr_html)}</tbody></table>'
        f'</div>'
    )


def _form_section_html(at, rep, data, logo_uri, page_break=False):
    spec = data["spec"]
    k = data["kpi"]
    ncols = len(spec["columns"])
    doc_col = spec["doc_col_idx"]
    hdr_block = _co_header_html(logo_uri, f"แผนซ่อมบำรุงรักษา - {spec['title_label']}", rep["date_range_text"])
    pills = (
        f'<div class="kpi-strip no-print">'
        f'<span class="kpi-pill">Total (Master): <b>{k["total_master"]:,}</b></span>'
        f'<span class="kpi-pill">มีงานซ่อม: <b>{k["repaired_in_master"]:,}</b> ({k["pct_repaired"]*100:.1f}%)</span>'
        + (f'<span class="kpi-pill warn">นอก Master รอคุณแขกตรวจ: <b>{k["unmatched_assets"]:,}</b></span>' if k["unmatched_assets"] else "")
        + '</div>'
    )
    colgroup = "".join(f'<col style="width:{w}"/>' for w in spec["html_widths"])
    col_ths = "".join(f'<th>{html.escape(c[0]).replace(chr(10), "<br/>")}</th>' for c in spec["columns"])

    body_rows = []
    for item in data["rows"]:
        tds = []
        for idx in range(1, ncols + 1):
            if idx == 1:
                tds.append(f'<td class="c">{html.escape(item["asset_code"])}</td>')
            elif idx == doc_col:
                tds.append(f'<td class="doc">{html.escape(item["doc_text"])}</td>')
            else:
                tds.append('<td></td>')
        body_rows.append(f'<tr>{"".join(tds)}</tr>')

    pb_cls = " page-break" if page_break else ""
    return (
        f'<div class="sheet{pb_cls}">{pills}'
        f'<table class="rpt"><colgroup>{colgroup}</colgroup>'
        f'<thead>'
        f'<tr><th class="repeat-hdr-cell" colspan="{ncols}">{hdr_block}</th></tr>'
        f'<tr>{col_ths}</tr>'
        f'</thead>'
        f'<tbody>{"".join(body_rows)}</tbody>'
        f'</table></div>'
    )


def build_print_html(db_path, report_month, section="all", logo_path=None):
    """สร้างหน้า HTML พร้อมสั่งพิมพ์ / Save PDF (section: 'all' | 'kpi' | 'forms' | 'หัวลาก' | 'หางพ่วง' | 'ตู้เย็น')"""
    rep = compute_monthly_report(db_path, report_month)
    yearly = compute_yearly_kpi(db_path, year=report_month[:4])
    logo_uri = _logo_data_uri(logo_path)

    parts = []
    if section in ("all", "kpi"):
        parts.append(_kpi_section_html(rep, yearly, logo_uri))
    if section in ("all", "forms"):
        for i, at in enumerate(TYPES):
            parts.append(_form_section_html(at, rep, rep["by_type"][at], logo_uri, page_break=(len(parts) > 0)))
    elif section in TYPES:
        parts.append(_form_section_html(section, rep, rep["by_type"][section], logo_uri, page_break=False))

    title_map = {
        "all": f"รายงาน KPI และแบบฟอร์มซ่อมประจำเดือน {report_month}",
        "kpi": f"สรุป KPI ประจำเดือน {report_month}",
        "forms": f"แบบฟอร์มซ่อมประจำเดือน 1_หัวลาก 2_หางพ่วง 3_ซ่อมตู้ ({report_month})",
    }
    doc_title = title_map[section] if section in title_map else f"{FORM_SPECS[section]['sheet_name']} ({report_month})"
    toolbar = (
        f'<div class="toolbar no-print">'
        f'<div><b>🖨️ {html.escape(doc_title)}</b> — กดปุ่มพิมพ์ หรือ Ctrl+P (เลือก Save as PDF ได้ทันที · หัวกระดาษซ้ำทุกหน้า)</div>'
        f'<button onclick="window.print()">🖨️ พิมพ์ / บันทึก PDF</button>'
        f'</div>'
    )
    return (
        f'<!DOCTYPE html><html lang="th"><head><meta charset="utf-8"/>'
        f'<title>{html.escape(doc_title)}</title>'
        f'<style>{PRINT_CSS}</style></head><body>'
        f'{toolbar}{"".join(parts)}</body></html>'
    )
