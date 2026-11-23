# -*- coding: utf-8 -*-
"""
Vehicle Report DB engine v1.0.0 — เก็บข้อมูลรายงานซ่อม & เบิกพัสดุรายคัน (Crystal Reports) ลง SQLite
- ใช้ไฟล์ SQLite เดียวกับระบบหลัก (Repair_Request_v2.sqlite) แต่แยกตารางต่างหาก ไม่กระทบตารางใบแจ้งซ่อม/Master/เลขไมล์/PM
- มี UNIQUE INDEX ป้องกันข้อมูลซ้ำระดับฐานข้อมูลทุกตาราง:
    vr_issue_line  -> UNIQUE (plate, doc_no, line_no)
    vr_wo_header   -> UNIQUE (plate, doc_no)
    vr_wo_line     -> UNIQUE (plate, doc_no, line_no)
- นำเข้าไฟล์เดิมซ้ำได้ไม่เกิดแถวซ้ำ (เทียบด้วย row_hash) + สำรองฐานข้อมูลอัตโนมัติก่อนบันทึกทุกครั้ง
- รองรับการซิงก์เข้าตาราง service_doc (PM Plan รายเดือน) พร้อมกันในคลิกเดียว
"""
import io
import os
import json
import hashlib
import datetime as dt

import repair_request_db as rr

VR_DB_VERSION = "1.0.0"

VR_SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS vr_import_batch (
    batch_id          INTEGER PRIMARY KEY AUTOINCREMENT,
    source_desc       TEXT NOT NULL,
    imported_by       TEXT,
    imported_at       TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    files_used        INTEGER NOT NULL DEFAULT 0,
    issues_inserted   INTEGER NOT NULL DEFAULT 0,
    issues_updated    INTEGER NOT NULL DEFAULT 0,
    issues_unchanged  INTEGER NOT NULL DEFAULT 0,
    wo_hdr_inserted   INTEGER NOT NULL DEFAULT 0,
    wo_hdr_updated    INTEGER NOT NULL DEFAULT 0,
    wo_hdr_unchanged  INTEGER NOT NULL DEFAULT 0,
    wo_lines_inserted INTEGER NOT NULL DEFAULT 0,
    wo_lines_updated  INTEGER NOT NULL DEFAULT 0,
    wo_lines_unchanged INTEGER NOT NULL DEFAULT 0,
    pm_synced         INTEGER NOT NULL DEFAULT 0,
    backup_path       TEXT,
    note              TEXT
);

CREATE TABLE IF NOT EXISTS vr_issue_line (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    plate           TEXT NOT NULL,
    brand           TEXT,
    vtype           TEXT,
    doc_no          TEXT NOT NULL,
    line_no         INTEGER NOT NULL,
    doc_date        TEXT,
    report_month    TEXT,
    inv_code        TEXT,
    inv_name        TEXT,
    qty             REAL,
    unit_price      REAL,
    total           REAL,
    dept            TEXT,
    store           TEXT,
    mileage_raw     TEXT,
    mileage_km      INTEGER,
    mileage_status  TEXT,
    mileage_src     TEXT,
    remark          TEXT,
    confirm_flag    TEXT NOT NULL DEFAULT 'Yes' CHECK (confirm_flag IN ('Yes','No')),
    flag_reason     TEXT,
    source_file     TEXT,
    src_row         INTEGER,
    row_hash        TEXT NOT NULL,
    first_batch_id  INTEGER NOT NULL REFERENCES vr_import_batch(batch_id),
    last_batch_id   INTEGER NOT NULL REFERENCES vr_import_batch(batch_id),
    created_at      TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    updated_at      TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE UNIQUE INDEX IF NOT EXISTS ux_vr_issue_key   ON vr_issue_line(plate, doc_no, line_no);
CREATE INDEX        IF NOT EXISTS ix_vr_issue_month ON vr_issue_line(report_month, plate);
CREATE INDEX        IF NOT EXISTS ix_vr_issue_plate ON vr_issue_line(plate, doc_date);

CREATE TABLE IF NOT EXISTS vr_wo_header (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    plate           TEXT NOT NULL,
    brand           TEXT,
    vtype           TEXT,
    doc_no          TEXT NOT NULL,
    doc_date        TEXT,
    report_month    TEXT,
    vendor_en       TEXT,
    vendor_th       TEXT,
    remark          TEXT,
    amount          REAL,
    discount        REAL,
    net             REAL,
    line_sum        REAL,
    n_lines         INTEGER,
    has_vat_line    INTEGER NOT NULL DEFAULT 0,
    mileage_km      INTEGER,
    mileage_status  TEXT,
    mileage_src     TEXT,
    confirm_flag    TEXT NOT NULL DEFAULT 'Yes' CHECK (confirm_flag IN ('Yes','No')),
    flag_reason     TEXT,
    source_file     TEXT,
    src_row         INTEGER,
    row_hash        TEXT NOT NULL,
    first_batch_id  INTEGER NOT NULL REFERENCES vr_import_batch(batch_id),
    last_batch_id   INTEGER NOT NULL REFERENCES vr_import_batch(batch_id),
    created_at      TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    updated_at      TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE UNIQUE INDEX IF NOT EXISTS ux_vr_wo_hdr_key   ON vr_wo_header(plate, doc_no);
CREATE INDEX        IF NOT EXISTS ix_vr_wo_hdr_month ON vr_wo_header(report_month, plate);
CREATE INDEX        IF NOT EXISTS ix_vr_wo_hdr_plate ON vr_wo_header(plate, doc_date);

CREATE TABLE IF NOT EXISTS vr_wo_line (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    plate           TEXT NOT NULL,
    brand           TEXT,
    vtype           TEXT,
    doc_no          TEXT NOT NULL,
    line_no         INTEGER NOT NULL,
    doc_date        TEXT,
    report_month    TEXT,
    item_code       TEXT,
    item_name       TEXT,
    qty             REAL,
    unit_price      REAL,
    total           REAL,
    is_vat          INTEGER NOT NULL DEFAULT 0,
    vendor_en       TEXT,
    mileage_km      INTEGER,
    mileage_status  TEXT,
    confirm_flag    TEXT NOT NULL DEFAULT 'Yes' CHECK (confirm_flag IN ('Yes','No')),
    flag_reason     TEXT,
    source_file     TEXT,
    src_row         INTEGER,
    row_hash        TEXT NOT NULL,
    first_batch_id  INTEGER NOT NULL REFERENCES vr_import_batch(batch_id),
    last_batch_id   INTEGER NOT NULL REFERENCES vr_import_batch(batch_id),
    created_at      TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    updated_at      TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE UNIQUE INDEX IF NOT EXISTS ux_vr_wo_line_key   ON vr_wo_line(plate, doc_no, line_no);
CREATE INDEX        IF NOT EXISTS ix_vr_wo_line_month ON vr_wo_line(report_month, plate);
CREATE INDEX        IF NOT EXISTS ix_vr_wo_line_plate ON vr_wo_line(plate, doc_date);
"""


def connect(db_path):
    con = rr.connect(db_path)
    con.executescript(VR_SCHEMA_SQL)
    return con


def _iso_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    s = str(v).strip()
    return s[:10] if len(s) >= 10 else None


def _month(d_iso):
    return d_iso[:7] if d_iso and len(d_iso) >= 7 else None


def _f(v):
    if v is None or v == "":
        return None
    try:
        x = float(v)
        return None if x != x else round(x, 4)
    except (TypeError, ValueError):
        return None


def _i(v):
    if v is None or v == "":
        return None
    try:
        x = float(v)
        return None if x != x else int(round(x))
    except (TypeError, ValueError):
        return None


def _s(v):
    if v is None:
        return None
    t = str(v).strip()
    return t if t != "" else None


def _hash_dict(d, ignore=("row_hash", "source_file", "src_row")):
    payload = {k: d[k] for k in sorted(d) if k not in ignore}
    return hashlib.sha256(json.dumps(payload, ensure_ascii=False, default=str).encode("utf-8")).hexdigest()[:24]


def _norm_issue(r):
    plate = _s(r.get("plate"))
    doc_no = _s(r.get("doc_no"))
    line_no = _i(r.get("line_no"))
    if not plate or not doc_no or line_no is None:
        return None
    d_iso = _iso_date(r.get("date"))
    rec = dict(
        plate=plate,
        brand=_s(r.get("brand")),
        vtype=_s(r.get("vtype")),
        doc_no=doc_no,
        line_no=line_no,
        doc_date=d_iso,
        report_month=_month(d_iso),
        inv_code=_s(r.get("inv_code")),
        inv_name=_s(r.get("inv_name")),
        qty=_f(r.get("qty")),
        unit_price=_f(r.get("unit_price")),
        total=_f(r.get("total")),
        dept=_s(r.get("dept")),
        store=_s(r.get("store")),
        mileage_raw=_s(r.get("mileage_raw")),
        mileage_km=_i(r.get("mileage_km")),
        mileage_status=_s(r.get("mileage_status")),
        mileage_src=_s(r.get("mileage_src")),
        remark=_s(r.get("remark")),
        confirm_flag="No" if str(r.get("confirm_flag")) == "No" else "Yes",
        flag_reason=_s(r.get("flag_reason")),
        source_file=_s(r.get("file")),
        src_row=_i(r.get("src_row")),
    )
    rec["row_hash"] = _hash_dict(rec)
    return rec


def _norm_wo_hdr(r):
    plate = _s(r.get("plate"))
    doc_no = _s(r.get("doc_no"))
    if not plate or not doc_no:
        return None
    d_iso = _iso_date(r.get("date"))
    rec = dict(
        plate=plate,
        brand=_s(r.get("brand")),
        vtype=_s(r.get("vtype")),
        doc_no=doc_no,
        doc_date=d_iso,
        report_month=_month(d_iso),
        vendor_en=_s(r.get("vendor_en")),
        vendor_th=_s(r.get("vendor_th")),
        remark=_s(r.get("remark")),
        amount=_f(r.get("amount")),
        discount=_f(r.get("discount")),
        net=_f(r.get("net")),
        line_sum=_f(r.get("line_sum")),
        n_lines=_i(r.get("n_lines")),
        has_vat_line=1 if r.get("has_vat_line") else 0,
        mileage_km=_i(r.get("mileage_km")),
        mileage_status=_s(r.get("mileage_status")),
        mileage_src=_s(r.get("mileage_src")),
        confirm_flag="No" if str(r.get("confirm_flag")) == "No" else "Yes",
        flag_reason=_s(r.get("flag_reason")),
        source_file=_s(r.get("file")),
        src_row=_i(r.get("src_row")),
    )
    rec["row_hash"] = _hash_dict(rec)
    return rec


def _norm_wo_line(r):
    plate = _s(r.get("plate"))
    doc_no = _s(r.get("doc_no"))
    line_no = _i(r.get("line_no"))
    if not plate or not doc_no or line_no is None:
        return None
    d_iso = _iso_date(r.get("date"))
    rec = dict(
        plate=plate,
        brand=_s(r.get("brand")),
        vtype=_s(r.get("vtype")),
        doc_no=doc_no,
        line_no=line_no,
        doc_date=d_iso,
        report_month=_month(d_iso),
        item_code=_s(r.get("item_code")),
        item_name=_s(r.get("item_name")),
        qty=_f(r.get("qty")),
        unit_price=_f(r.get("unit_price")),
        total=_f(r.get("total")),
        is_vat=1 if r.get("is_vat") else 0,
        vendor_en=_s(r.get("vendor_en")),
        mileage_km=_i(r.get("mileage_km")),
        mileage_status=_s(r.get("mileage_status")),
        confirm_flag="No" if str(r.get("confirm_flag")) == "No" else "Yes",
        flag_reason=_s(r.get("flag_reason")),
        source_file=_s(r.get("file")),
        src_row=_i(r.get("src_row")),
    )
    rec["row_hash"] = _hash_dict(rec)
    return rec


def _prepare(ds):
    """แปลง ds จาก core.scan_folder เป็น dict ที่ตัดแถวคีย์ซ้ำในรอบเดียวกันไว้แล้ว (last wins)"""
    issues, wo_hdr, wo_lines = {}, {}, {}
    skip_i = skip_h = skip_l = 0
    for r in ds.get("issues", []):
        rec = _norm_issue(r)
        if rec is None:
            skip_i += 1
        else:
            issues[(rec["plate"], rec["doc_no"], rec["line_no"])] = rec
    for r in ds.get("wo_hdr", []):
        rec = _norm_wo_hdr(r)
        if rec is None:
            skip_h += 1
        else:
            wo_hdr[(rec["plate"], rec["doc_no"])] = rec
    for r in ds.get("wo_lines", []):
        rec = _norm_wo_line(r)
        if rec is None:
            skip_l += 1
        else:
            wo_lines[(rec["plate"], rec["doc_no"], rec["line_no"])] = rec
    return dict(
        issues=issues, wo_hdr=wo_hdr, wo_lines=wo_lines,
        skipped=dict(issues=skip_i, wo_hdr=skip_h, wo_lines=skip_l),
    )


def analyze_dataset(db_path, ds):
    """ตรวจเทียบข้อมูลที่จะบันทึกกับฐานข้อมูล (อ่านอย่างเดียว ไม่แก้ไขฐานข้อมูล)"""
    prep = _prepare(ds)
    cur_i, cur_h, cur_l = {}, {}, {}
    if os.path.exists(db_path):
        con = connect(db_path)
        try:
            cur_i = {(r[0], r[1], r[2]): r[3] for r in con.execute("SELECT plate, doc_no, line_no, row_hash FROM vr_issue_line")}
            cur_h = {(r[0], r[1]): r[2] for r in con.execute("SELECT plate, doc_no, row_hash FROM vr_wo_header")}
            cur_l = {(r[0], r[1], r[2]): r[3] for r in con.execute("SELECT plate, doc_no, line_no, row_hash FROM vr_wo_line")}
        finally:
            con.close()

    def _diff(incoming, existing):
        new, upd, same = [], [], []
        for k, rec in incoming.items():
            if k not in existing:
                new.append(rec)
            elif existing[k] != rec["row_hash"]:
                upd.append(rec)
            else:
                same.append(rec)
        return dict(new=new, update=upd, unchanged=same)

    return dict(
        issues=_diff(prep["issues"], cur_i),
        wo_hdr=_diff(prep["wo_hdr"], cur_h),
        wo_lines=_diff(prep["wo_lines"], cur_l),
        skipped=prep["skipped"],
    )


_ISSUE_COLS = [
    "plate", "brand", "vtype", "doc_no", "line_no", "doc_date", "report_month",
    "inv_code", "inv_name", "qty", "unit_price", "total", "dept", "store",
    "mileage_raw", "mileage_km", "mileage_status", "mileage_src", "remark",
    "confirm_flag", "flag_reason", "source_file", "src_row", "row_hash",
]
_WO_HDR_COLS = [
    "plate", "brand", "vtype", "doc_no", "doc_date", "report_month",
    "vendor_en", "vendor_th", "remark", "amount", "discount", "net", "line_sum",
    "n_lines", "has_vat_line", "mileage_km", "mileage_status", "mileage_src",
    "confirm_flag", "flag_reason", "source_file", "src_row", "row_hash",
]
_WO_LINE_COLS = [
    "plate", "brand", "vtype", "doc_no", "line_no", "doc_date", "report_month",
    "item_code", "item_name", "qty", "unit_price", "total", "is_vat", "vendor_en",
    "mileage_km", "mileage_status", "confirm_flag", "flag_reason",
    "source_file", "src_row", "row_hash",
]


def _upsert_sql(table, cols, conflict_cols):
    all_cols = cols + ["first_batch_id", "last_batch_id"]
    placeholders = ",".join("?" * len(all_cols))
    upd_cols = [c for c in cols if c not in conflict_cols] + ["last_batch_id"]
    set_clause = ", ".join(f"{c}=excluded.{c}" for c in upd_cols) + ", updated_at=datetime('now','localtime')"
    return (
        f"INSERT INTO {table} ({','.join(all_cols)}) VALUES ({placeholders}) "
        f"ON CONFLICT({','.join(conflict_cols)}) DO UPDATE SET {set_clause}"
    )


def _build_pm_service_docs(ds):
    """แปลง ds เป็นรายการเอกสารสำหรับตาราง service_doc (PM Plan รายเดือน)"""
    import pm_plan as pm
    docs = {}
    wo_items = {}
    for l in ds.get("wo_lines", []):
        if pm.is_center_item(l.get("item_name")):
            no = _s(l.get("doc_no"))
            if no:
                wo_items.setdefault(no, []).append(str(l.get("item_name")).strip())
    for h in ds.get("wo_hdr", []):
        no = _s(h.get("doc_no"))
        plate = _s(h.get("plate"))
        if not no or not plate:
            continue
        items = wo_items.get(no, [])
        docs[no] = dict(
            doc_no=no, doc_kind="W/O", asset_code=plate, doc_date=_iso_date(h.get("date")),
            km=_i(h.get("mileage_km")), km_status=_s(h.get("mileage_status")),
            vendor=_s(h.get("vendor_th")) or _s(h.get("vendor_en")),
            pm_type="ศูนย์บริการ" if items else None,
            pm_items=" · ".join(sorted(set(items))) or None,
            amount=h.get("net"),
        )
    for l in ds.get("issues", []):
        no = _s(l.get("doc_no"))
        plate = _s(l.get("plate"))
        if not no or not plate:
            continue
        d = docs.get(no)
        if d is None:
            d = docs[no] = dict(
                doc_no=no, doc_kind="WX", asset_code=plate, doc_date=_iso_date(l.get("date")),
                km=None, km_status=None, vendor=_s(l.get("store")),
                pm_type=None, pm_items=None, amount=0.0,
            )
        km, stt = _i(l.get("mileage_km")), _s(l.get("mileage_status"))
        if km is not None and (d["km"] is None or (d["km_status"] != "ชัดเจน" and stt == "ชัดเจน")):
            d["km"], d["km_status"] = km, stt
        d["amount"] = (d["amount"] or 0.0) + (float(l.get("total") or 0.0))
        if pm.is_yard_oil(l.get("inv_name")):
            d["pm_type"] = "ลาน"
            its = set(filter(None, (d["pm_items"] or "").split(" · "))) | {str(l.get("inv_name")).strip()}
            d["pm_items"] = " · ".join(sorted(its))
    out = []
    for d in docs.values():
        d["amount"] = round(float(d["amount"] or 0.0), 2)
        d["row_hash"] = rr.file_sha256(json.dumps({k: d[k] for k in sorted(d)}, ensure_ascii=False, default=str).encode())[:20]
        out.append(d)
    return out


def commit_dataset(db_path, ds, user=None, sync_pm=True, note=None):
    """บันทึกข้อมูลรายงานซ่อมรายคันลง SQLite (สำรองฐานอัตโนมัติ + กันซ้ำด้วย UNIQUE INDEX + ซิงก์ PM Plan ได้)"""
    bpath = rr.backup_db(db_path)
    plan = analyze_dataset(db_path, ds)
    src_desc = str(ds.get("summary", {}).get("folder") or "Vehicle Reports")
    files_used = int(ds.get("summary", {}).get("n_used") or 0)

    con = connect(db_path)
    try:
        pm_stats = dict(inserted=0, updated=0, unchanged=0)
        with con:
            bid = con.execute(
                "INSERT INTO vr_import_batch(source_desc, imported_by, files_used, backup_path, note) VALUES (?,?,?,?,?)",
                (src_desc, user, files_used, bpath, note),
            ).lastrowid

            sql_i = _upsert_sql("vr_issue_line", _ISSUE_COLS, ("plate", "doc_no", "line_no"))
            for rec in plan["issues"]["new"] + plan["issues"]["update"]:
                con.execute(sql_i, [rec[c] for c in _ISSUE_COLS] + [bid, bid])

            sql_h = _upsert_sql("vr_wo_header", _WO_HDR_COLS, ("plate", "doc_no"))
            for rec in plan["wo_hdr"]["new"] + plan["wo_hdr"]["update"]:
                con.execute(sql_h, [rec[c] for c in _WO_HDR_COLS] + [bid, bid])

            sql_l = _upsert_sql("vr_wo_line", _WO_LINE_COLS, ("plate", "doc_no", "line_no"))
            for rec in plan["wo_lines"]["new"] + plan["wo_lines"]["update"]:
                con.execute(sql_l, [rec[c] for c in _WO_LINE_COLS] + [bid, bid])

            if sync_pm:
                import pm_plan as pm
                con.executescript(pm.PM_SQL)
                pm._ensure_composite_service_doc(con)
                n_synced = pm.sync_from_vr_tables(con)
                pm_stats = dict(inserted=n_synced, updated=0, unchanged=0)

            con.execute(
                """UPDATE vr_import_batch SET
                   issues_inserted=?, issues_updated=?, issues_unchanged=?,
                   wo_hdr_inserted=?, wo_hdr_updated=?, wo_hdr_unchanged=?,
                   wo_lines_inserted=?, wo_lines_updated=?, wo_lines_unchanged=?,
                   pm_synced=?
                   WHERE batch_id=?""",
                (
                    len(plan["issues"]["new"]), len(plan["issues"]["update"]), len(plan["issues"]["unchanged"]),
                    len(plan["wo_hdr"]["new"]), len(plan["wo_hdr"]["update"]), len(plan["wo_hdr"]["unchanged"]),
                    len(plan["wo_lines"]["new"]), len(plan["wo_lines"]["update"]), len(plan["wo_lines"]["unchanged"]),
                    pm_stats["inserted"] + pm_stats["updated"],
                    bid,
                ),
            )

        return dict(
            batch_id=bid,
            backup=bpath,
            issues={k: len(v) for k, v in plan["issues"].items()},
            wo_hdr={k: len(v) for k, v in plan["wo_hdr"].items()},
            wo_lines={k: len(v) for k, v in plan["wo_lines"].items()},
            pm_sync=pm_stats,
        )
    finally:
        con.close()


def db_summary(db_path):
    """สรุปภาพรวมข้อมูลรายงานซ่อมรายคันใน SQLite (ไม่สร้างไฟล์ใหม่ถ้ายังไม่มี)"""
    if not db_path or not os.path.exists(db_path):
        return None
    con = connect(db_path)
    try:
        ni, n_idocs, wx_sum, di1, di2 = con.execute(
            "SELECT COUNT(*), COUNT(DISTINCT plate||'|'||doc_no), COALESCE(SUM(total),0), MIN(doc_date), MAX(doc_date) FROM vr_issue_line"
        ).fetchone()
        nh, wo_net = con.execute("SELECT COUNT(*), COALESCE(SUM(net),0) FROM vr_wo_header").fetchone()
        nl, wo_sum, dw1, dw2 = con.execute(
            "SELECT COUNT(*), COALESCE(SUM(total),0), MIN(doc_date), MAX(doc_date) FROM vr_wo_line"
        ).fetchone()
        plates = [r[0] for r in con.execute(
            "SELECT DISTINCT plate FROM ("
            "SELECT plate FROM vr_issue_line UNION SELECT plate FROM vr_wo_header UNION SELECT plate FROM vr_wo_line"
            ") ORDER BY 1"
        )]
        months = [r[0] for r in con.execute(
            "SELECT DISTINCT report_month FROM ("
            "SELECT report_month FROM vr_issue_line UNION SELECT report_month FROM vr_wo_header UNION SELECT report_month FROM vr_wo_line"
            ") WHERE report_month IS NOT NULL ORDER BY 1 DESC"
        )]
        batches = [dict(r) for r in con.execute("SELECT * FROM vr_import_batch ORDER BY batch_id DESC LIMIT 30")]
        d_min = min([x for x in (di1, dw1) if x], default=None)
        d_max = max([x for x in (di2, dw2) if x], default=None)
        return dict(
            n_issue_lines=ni, n_issue_docs=n_idocs, wx_total=round(wx_sum, 2),
            n_wo_hdr=nh, wo_net=round(wo_net, 2),
            n_wo_lines=nl, wo_total=round(wo_sum, 2),
            grand_total=round(wx_sum + wo_sum, 2),
            n_plates=len(plates), plates=plates, months=months,
            date_min=d_min, date_max=d_max, batches=batches,
        )
    finally:
        con.close()


def query_rows(db_path, table_kind, months=None, plates=None, flag_only=False, search=None):
    """ดึงข้อมูลจาก vr_issue_line / vr_wo_header / vr_wo_line ตามเงื่อนไข"""
    tbl_map = {
        "issues": ("vr_issue_line", "doc_date DESC, plate, doc_no, line_no",
                   ("plate", "doc_no", "inv_code", "inv_name", "store", "dept", "remark")),
        "wo_hdr": ("vr_wo_header", "doc_date DESC, plate, doc_no",
                   ("plate", "doc_no", "vendor_en", "vendor_th", "remark")),
        "wo_lines": ("vr_wo_line", "doc_date DESC, plate, doc_no, line_no",
                     ("plate", "doc_no", "item_code", "item_name", "vendor_en")),
    }
    table, order_by, search_cols = tbl_map[table_kind]
    where, params = [], []
    if months:
        where.append(f"report_month IN ({','.join('?' * len(months))})")
        params.extend(months)
    if plates:
        where.append(f"plate IN ({','.join('?' * len(plates))})")
        params.extend(plates)
    if flag_only:
        where.append("confirm_flag = 'No'")
    if search and search.strip():
        q = f"%{search.strip().lower()}%"
        where.append("(" + " OR ".join(f"LOWER(COALESCE({c},'')) LIKE ?" for c in search_cols) + ")")
        params.extend([q] * len(search_cols))
    sql = f"SELECT * FROM {table}" + (" WHERE " + " AND ".join(where) if where else "") + f" ORDER BY {order_by}"
    con = connect(db_path)
    try:
        return [dict(r) for r in con.execute(sql, params)]
    finally:
        con.close()


def export_db_excel(db_path, months=None, plates=None, flag_only=False, search=None):
    """Export ข้อมูลรายงานซ่อมรายคันจาก SQLite (ตามตัวกรอง) เป็นไฟล์ Excel 4 ชีต"""
    import pandas as pd
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    iss = query_rows(db_path, "issues", months, plates, flag_only, search)
    woh = query_rows(db_path, "wo_hdr", months, plates, flag_only, search)
    wol = query_rows(db_path, "wo_lines", months, plates, flag_only, search)

    # สรุปรายคัน
    veh_map = {}
    for r in iss:
        v = veh_map.setdefault(r["plate"], dict(plate=r["plate"], brand=r["brand"] or "", vtype=r["vtype"] or "",
                                                wx_docs=set(), wx=0.0, wo_docs=set(), wo=0.0))
        v["wx_docs"].add(r["doc_no"])
        v["wx"] += float(r["total"] or 0.0)
    for r in wol:
        v = veh_map.setdefault(r["plate"], dict(plate=r["plate"], brand=r["brand"] or "", vtype=r["vtype"] or "",
                                                wx_docs=set(), wx=0.0, wo_docs=set(), wo=0.0))
        v["wo_docs"].add(r["doc_no"])
        v["wo"] += float(r["total"] or 0.0)
    veh_rows = [
        {
            "ทะเบียน": v["plate"], "ยี่ห้อ": v["brand"], "ประเภท": v["vtype"],
            "ใบเบิก (WX)": len(v["wx_docs"]), "ยอดเบิก (WX)": round(v["wx"], 2),
            "ใบสั่งซ่อม (W/O)": len(v["wo_docs"]), "ยอดสั่งซ่อม (W/O)": round(v["wo"], 2),
            "ยอดรวม (บาท)": round(v["wx"] + v["wo"], 2),
        }
        for v in sorted(veh_map.values(), key=lambda x: -(x["wx"] + x["wo"]))
    ]

    iss_cols = [
        ("plate", "ทะเบียน"), ("brand", "ยี่ห้อ"), ("doc_no", "เลขที่ใบเบิก"), ("line_no", "ลำดับ"),
        ("doc_date", "วันที่"), ("inv_code", "รหัสพัสดุ"), ("inv_name", "ชื่อพัสดุ"),
        ("qty", "จำนวน"), ("unit_price", "ราคา/หน่วย"), ("total", "ยอดสุทธิ"),
        ("dept", "แผนก"), ("store", "คลัง"), ("mileage_km", "เลขไมล์"), ("mileage_status", "สถานะเลขไมล์"),
        ("remark", "หมายเหตุ"), ("confirm_flag", "Confirm Flag"), ("flag_reason", "เหตุผล Flag"),
    ]
    woh_cols = [
        ("plate", "ทะเบียน"), ("brand", "ยี่ห้อ"), ("doc_no", "เลขที่ใบสั่งซ่อม"), ("doc_date", "วันที่"),
        ("vendor_th", "ผู้ขาย (TH)"), ("vendor_en", "ผู้ขาย (EN)"), ("amount", "ยอดก่อนลด"),
        ("discount", "ส่วนลด"), ("net", "ยอดสุทธิ"), ("line_sum", "ผลรวมบรรทัด"), ("n_lines", "จำนวนบรรทัด"),
        ("mileage_km", "เลขไมล์"), ("mileage_status", "สถานะเลขไมล์"), ("remark", "หมายเหตุ"),
        ("confirm_flag", "Confirm Flag"), ("flag_reason", "เหตุผล Flag"),
    ]
    wol_cols = [
        ("plate", "ทะเบียน"), ("brand", "ยี่ห้อ"), ("doc_no", "เลขที่ใบสั่งซ่อม"), ("line_no", "ลำดับ"),
        ("doc_date", "วันที่"), ("item_code", "รหัสรายการ"), ("item_name", "ชื่อรายการ"),
        ("qty", "จำนวน"), ("unit_price", "ราคา/หน่วย"), ("total", "ยอดสุทธิ"),
        ("vendor_en", "ผู้ขาย"), ("mileage_km", "เลขไมล์"), ("confirm_flag", "Confirm Flag"),
    ]

    import re as _re
    _BAD_XL_RE = _re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

    def _to_df(rows, cols):
        if not rows:
            return pd.DataFrame(columns=[h for _, h in cols])
        df = pd.DataFrame(rows)[[k for k, _ in cols]].rename(columns=dict(cols))
        for col_c in df.columns:
            if df[col_c].dtype == object:
                df[col_c] = df[col_c].map(lambda v: _BAD_XL_RE.sub("", v) if isinstance(v, str) else v)
        if "วันที่" in df.columns:
            df["วันที่"] = pd.to_datetime(df["วันที่"], errors="coerce")
        return df

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame(veh_rows).to_excel(xw, sheet_name="สรุปรายคัน", index=False)
        _to_df(iss, iss_cols).to_excel(xw, sheet_name="เบิกพัสดุ_WX", index=False)
        _to_df(woh, woh_cols).to_excel(xw, sheet_name="ใบสั่งซ่อม_หัว", index=False)
        _to_df(wol, wol_cols).to_excel(xw, sheet_name="ใบสั่งซ่อม_บรรทัด", index=False)

        fill_hdr = PatternFill("solid", fgColor="DCE9F5")
        money_hdrs = {"ยอดเบิก (WX)", "ยอดสั่งซ่อม (W/O)", "ยอดรวม (บาท)", "ราคา/หน่วย", "ยอดสุทธิ", "ยอดก่อนลด", "ส่วนลด", "ผลรวมบรรทัด"}
        for ws in xw.book.worksheets:
            ws.freeze_panes = "B2"
            for i, cell in enumerate(ws[1], 1):
                cell.font = Font(bold=True)
                cell.fill = fill_hdr
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ws.column_dimensions[get_column_letter(i)].width = 16
                if cell.value == "วันที่":
                    for c in ws[get_column_letter(i)][1:]:
                        c.number_format = "dd/mm/yyyy"
                elif cell.value in money_hdrs:
                    for c in ws[get_column_letter(i)][1:]:
                        c.number_format = "#,##0.00"
                elif cell.value in ("ชื่อพัสดุ", "ชื่อรายการ", "ผู้ขาย (TH)", "ผู้ขาย (EN)", "หมายเหตุ"):
                    ws.column_dimensions[get_column_letter(i)].width = 34
            if ws.max_row > 1:
                ws.auto_filter.ref = ws.dimensions
    return buf.getvalue()


# ============================================================================= PM vs CM Summary & 6-Sheet Excel Export
def compute_pm_cm(db_path, months=None, plates=None):
    """คำนวณสรุป PM vs CM จาก vr_issue_line / vr_wo_header / vr_wo_line (แยก W/O กับ WX ไม่ชนเลขเอกสาร)"""
    import re as _re
    import pm_plan as pm

    bad_xl = _re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
    clean = lambda s: bad_xl.sub("", s) if isinstance(s, str) else s

    where, params = [], []
    if months:
        where.append(f"report_month IN ({','.join('?' * len(months))})")
        params.extend(months)
    if plates:
        where.append(f"plate IN ({','.join('?' * len(plates))})")
        params.extend(plates)
    w_sql = (" WHERE " + " AND ".join(where)) if where else ""

    con = connect(db_path)
    try:
        wo_eng_items = {}
        wo_pm_items = {}
        for p, d, code, name in con.execute(f"SELECT plate, doc_no, item_code, item_name FROM vr_wo_line{w_sql}", params):
            if pm.is_engine_replace_item(name, code):
                wo_eng_items.setdefault((p, d), set()).add(clean(str(name or code).strip()))
            elif pm.is_center_item(name, code):
                wo_pm_items.setdefault((p, d), set()).add(clean(str(name or code).strip()))

        wx_pm_items = {}
        for p, d, code, name in con.execute(f"SELECT plate, doc_no, inv_code, inv_name FROM vr_issue_line{w_sql}", params):
            if pm.is_yard_oil(name, code):
                wx_pm_items.setdefault((p, d), set()).add(clean(str(name or code).strip()))

        docs = []
        for r in con.execute(
            f"""SELECT plate, brand, doc_no, doc_date, report_month,
                       COALESCE(vendor_th, vendor_en, ''), mileage_km, COALESCE(net, 0),
                       confirm_flag, COALESCE(remark, '')
                FROM vr_wo_header{w_sql}
                ORDER BY doc_date, plate, doc_no""",
            params,
        ):
            p, br, dno, dd, ym, vend, km, amt, cflg, rem = r
            is_eng = (p, dno) in wo_eng_items
            is_pm = (not is_eng) and ((p, dno) in wo_pm_items)
            if is_eng:
                m_grp, m_sub = "เปลี่ยนเครื่องยนต์", "เปลี่ยนเครื่องยนต์ (W/O)"
                m_items = " · ".join(sorted(wo_eng_items.get((p, dno), [])))
            elif is_pm:
                m_grp, m_sub = "PM", "PM ศูนย์บริการ (W/O)"
                m_items = " · ".join(sorted(wo_pm_items.get((p, dno), [])))
            else:
                m_grp, m_sub = "CM", "CM ศูนย์/อู่นอก (W/O)"
                m_items = ""
            docs.append(dict(
                maint=m_grp, sub=m_sub,
                kind="W/O", plate=p, brand=br or "", doc_no=dno,
                doc_date=dt.date.fromisoformat(dd) if dd else None,
                ym=ym or "", vendor=clean(vend), km=km, amount=round(float(amt), 2),
                pm_items=m_items,
                confirm_flag=cflg or "Yes", remark=clean(rem),
            ))

        for r in con.execute(
            f"""SELECT plate, MAX(brand), doc_no, MIN(doc_date), MIN(report_month),
                       MAX(COALESCE(store, '')), MAX(mileage_km), SUM(COALESCE(total, 0)),
                       MIN(confirm_flag), MAX(COALESCE(remark, ''))
                FROM vr_issue_line{w_sql}
                GROUP BY plate, doc_no
                ORDER BY MIN(doc_date), plate, doc_no""",
            params,
        ):
            p, br, dno, dd, ym, store, km, amt, cflg, rem = r
            is_pm = (p, dno) in wx_pm_items
            docs.append(dict(
                maint="PM" if is_pm else "CM",
                sub="PM ลาน (WX)" if is_pm else "CM เบิกพัสดุทั่วไป (WX)",
                kind="WX", plate=p, brand=br or "", doc_no=dno,
                doc_date=dt.date.fromisoformat(dd) if dd else None,
                ym=ym or "", vendor=clean(store), km=km, amount=round(float(amt), 2),
                pm_items=" · ".join(sorted(wx_pm_items.get((p, dno), []))),
                confirm_flag=cflg or "Yes", remark=clean(rem),
            ))

        ref_rows = []
        for r in con.execute(
            f"""SELECT report_month, doc_date, plate, brand, doc_no, line_no,
                       inv_code, inv_name, qty, unit_price, total, store, mileage_km, confirm_flag, source_file
                FROM vr_issue_line{w_sql}""",
            params,
        ):
            ym, dd, p, br, dno, lno, code, name, qty, up, tot, store, km, cflg, sfile = r
            is_pm = (p, dno) in wx_pm_items
            ref_rows.append((
                ym or "", dt.date.fromisoformat(dd) if dd else None, p, br or "",
                "PM" if is_pm else "CM",
                "PM ลาน (WX)" if is_pm else "CM เบิกพัสดุทั่วไป (WX)",
                "WX", dno, lno, clean(code or ""), clean(name or ""),
                qty, up, round(float(tot or 0), 2), clean(store or ""), km, cflg or "Yes", clean(sfile or ""),
            ))

        for r in con.execute(
            f"""SELECT report_month, doc_date, plate, brand, doc_no, line_no,
                       item_code, item_name, qty, unit_price, total, vendor_en, mileage_km, confirm_flag, source_file
                FROM vr_wo_line{w_sql}""",
            params,
        ):
            ym, dd, p, br, dno, lno, code, name, qty, up, tot, vend, km, cflg, sfile = r
            is_eng = (p, dno) in wo_eng_items
            is_pm = (not is_eng) and ((p, dno) in wo_pm_items)
            ref_rows.append((
                ym or "", dt.date.fromisoformat(dd) if dd else None, p, br or "",
                "เปลี่ยนเครื่องยนต์" if is_eng else ("PM" if is_pm else "CM"),
                "เปลี่ยนเครื่องยนต์ (W/O)" if is_eng else ("PM ศูนย์บริการ (W/O)" if is_pm else "CM ศูนย์/อู่นอก (W/O)"),
                "W/O", dno, lno, clean(code or ""), clean(name or ""),
                qty, up, round(float(tot or 0), 2), clean(vend or ""), km, cflg or "Yes", clean(sfile or ""),
            ))
    finally:
        con.close()

    ref_rows.sort(key=lambda x: (x[1] or dt.date(1900, 1, 1), x[2], x[6], x[7], x[8] or 0))
    grand_amt = round(sum(x["amount"] for x in docs), 2)
    dates = [x["doc_date"] for x in docs if x["doc_date"]]
    d_min = min(dates) if dates else None
    d_max = max(dates) if dates else None

    sub_order = [
        ("PM", "PM ศูนย์บริการ (W/O)"),
        ("PM", "PM ลาน (WX)"),
        ("เปลี่ยนเครื่องยนต์", "เปลี่ยนเครื่องยนต์ (W/O)"),
        ("CM", "CM ศูนย์/อู่นอก (W/O)"),
        ("CM", "CM เบิกพัสดุทั่วไป (WX)"),
    ]
    summary_rows = []
    for grp, sub in sub_order:
        sd = [x for x in docs if x["sub"] == sub]
        amt = round(sum(x["amount"] for x in sd), 2)
        summary_rows.append({
            "กลุ่ม": grp, "ประเภทเอกสาร": sub, "จำนวนเอกสาร": len(sd),
            "จำนวนรถ": len({x["plate"] for x in sd}), "ยอดเงินรวม": amt,
            "สัดส่วน": (amt / grand_amt) if grand_amt else 0.0,
        })

    monthly_rows = []
    for m in sorted({x["ym"] for x in docs if x["ym"]}):
        pm_d = [x for x in docs if x["ym"] == m and x["maint"] == "PM"]
        eng_d = [x for x in docs if x["ym"] == m and x["maint"] == "เปลี่ยนเครื่องยนต์"]
        cm_d = [x for x in docs if x["ym"] == m and x["maint"] == "CM"]
        pa = round(sum(x["amount"] for x in pm_d), 2)
        ea = round(sum(x["amount"] for x in eng_d), 2)
        ca = round(sum(x["amount"] for x in cm_d), 2)
        ta = round(pa + ea + ca, 2)
        monthly_rows.append(dict(
            เดือน=m,
            pm_docs=len(pm_d), pm_amt=pa, pm_pct=(pa / ta if ta else 0.0),
            eng_docs=len(eng_d), eng_amt=ea, eng_pct=(ea / ta if ta else 0.0),
            cm_docs=len(cm_d), cm_amt=ca, cm_pct=(ca / ta if ta else 0.0),
            total_amt=ta,
        ))

    vmap = {}
    for x in docs:
        v = vmap.setdefault(x["plate"], dict(plate=x["plate"], brand=x["brand"], pm_wo=0, pm_wx=0, pm_a=0.0, eng_cnt=0, eng_a=0.0, cm_wo=0, cm_wx=0, cm_a=0.0))
        if not v["brand"] and x["brand"]:
            v["brand"] = x["brand"]
        if x["maint"] == "PM":
            if x["kind"] == "W/O":
                v["pm_wo"] += 1
            else:
                v["pm_wx"] += 1
            v["pm_a"] += x["amount"]
        elif x["maint"] == "เปลี่ยนเครื่องยนต์":
            v["eng_cnt"] += 1
            v["eng_a"] += x["amount"]
        else:
            if x["kind"] == "W/O":
                v["cm_wo"] += 1
            else:
                v["cm_wx"] += 1
            v["cm_a"] += x["amount"]
    vehicle_rows = sorted(vmap.values(), key=lambda x: -(x["pm_a"] + x["eng_a"] + x["cm_a"]))

    pm_amt = round(sum(x["amount"] for x in docs if x["maint"] == "PM"), 2)
    eng_amt = round(sum(x["amount"] for x in docs if x["maint"] == "เปลี่ยนเครื่องยนต์"), 2)
    cm_amt = round(sum(x["amount"] for x in docs if x["maint"] == "CM"), 2)
    pm_cnt = sum(1 for x in docs if x["maint"] == "PM")
    eng_cnt = sum(1 for x in docs if x["maint"] == "เปลี่ยนเครื่องยนต์")
    cm_cnt = sum(1 for x in docs if x["maint"] == "CM")

    return dict(
        docs=docs, ref_rows=ref_rows, summary_rows=summary_rows,
        monthly_rows=monthly_rows, vehicle_rows=vehicle_rows,
        pm_amt=pm_amt, eng_amt=eng_amt, cm_amt=cm_amt,
        pm_cnt=pm_cnt, eng_cnt=eng_cnt, cm_cnt=cm_cnt,
        grand_amt=grand_amt, d_min=d_min, d_max=d_max,
    )


def export_pm_cm_excel(db_path, months=None, plates=None):
    """สร้างไฟล์ Excel สรุป PM vs CM ครบ 6 ชีต (รวมกราฟและชีตรายละเอียดอ้างอิงทุกบรรทัด)"""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, Reference

    res = compute_pm_cm(db_path, months=months, plates=plates)
    docs = res["docs"]
    ref_rows = res["ref_rows"]
    grand_amt = res["grand_amt"]

    wb = openpyxl.Workbook()
    f_title = Font(name="Arial", size=14, bold=True, color="1F4E78")
    f_sec = Font(name="Arial", size=11, bold=True, color="1F4E78")
    f_hdr = Font(name="Arial", size=10, bold=True)
    f_cell = Font(name="Arial", size=10)
    f_bold = Font(name="Arial", size=10, bold=True)
    fill_hdr = PatternFill("solid", fgColor="DCE9F5")
    fill_pm = PatternFill("solid", fgColor="E2EFDA")
    fill_cm = PatternFill("solid", fgColor="FCE4D6")
    fill_tot = PatternFill("solid", fgColor="F2F2F2")
    thin = Side(style="thin", color="BFBFBF")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Sheet 1: สรุป_PM_CM
    ws1 = wb.active
    ws1.title = "สรุป_PM_CM"
    ws1["A1"] = "สรุปค่าใช้จ่ายการซ่อมบำรุงแยก PM กับ CM"
    ws1["A1"].font = f_title
    dr_txt = f"{res['d_min']:%d/%m/%Y} – {res['d_max']:%d/%m/%Y}" if res["d_min"] else "-"
    ws1["A2"] = f"ช่วงข้อมูล: {dr_txt}   |   สร้างเมื่อ: {dt.datetime.now():%d/%m/%Y %H:%M}"
    ws1["A2"].font = Font(name="Arial", size=10, color="555555")

    ws1["A4"] = "1. ภาพรวม PM vs CM แยกตามประเภทเอกสาร"
    ws1["A4"].font = f_sec
    hdr1 = ["กลุ่ม", "ประเภทเอกสาร", "จำนวนเอกสาร (ใบ)", "จำนวนรถ (คัน)", "ยอดเงินรวม (บาท)", "สัดส่วนยอดเงิน"]
    for j, h in enumerate(hdr1, 1):
        c = ws1.cell(5, j, h)
        c.font = f_hdr; c.fill = fill_hdr; c.border = b_thin
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[5].height = 24

    r = 6
    fill_eng = PatternFill("solid", fgColor="FFF2CC")
    for srow in res["summary_rows"]:
        grp = srow["กลุ่ม"]
        vals = [grp, srow["ประเภทเอกสาร"], srow["จำนวนเอกสาร"], srow["จำนวนรถ"], srow["ยอดเงินรวม"], srow["สัดส่วน"]]
        for j, v in enumerate(vals, 1):
            c = ws1.cell(r, j, v)
            c.font = f_cell; c.border = b_thin
            c.fill = fill_pm if grp == "PM" else (fill_eng if grp == "เปลี่ยนเครื่องยนต์" else fill_cm)
            if j == 1: c.alignment = Alignment(horizontal="center", vertical="center")
            elif j in (3, 4): c.number_format = "#,##0"; c.alignment = Alignment(horizontal="right", vertical="center")
            elif j == 5: c.number_format = "#,##0.00"; c.alignment = Alignment(horizontal="right", vertical="center")
            elif j == 6: c.number_format = "0.0%"; c.alignment = Alignment(horizontal="right", vertical="center")
        r += 1

    for grp, lbl, fl in [
        ("PM", "รวม PM ทั้งหมด", fill_pm),
        ("เปลี่ยนเครื่องยนต์", "รวม เปลี่ยนเครื่องยนต์ ทั้งหมด", fill_eng),
        ("CM", "รวม CM ทั้งหมด", fill_cm),
        ("ALL", "รวมทั้งหมด (PM + เปลี่ยนเครื่องยนต์ + CM)", fill_tot),
    ]:
        gd = docs if grp == "ALL" else [x for x in docs if x["maint"] == grp]
        amt = round(sum(x["amount"] for x in gd), 2)
        vals = [grp if grp != "ALL" else "รวม", lbl, len(gd), len({x["plate"] for x in gd}), amt, (amt / grand_amt) if grand_amt else 0]
        for j, v in enumerate(vals, 1):
            c = ws1.cell(r, j, v)
            c.font = f_bold; c.fill = fl; c.border = b_thin
            if j == 1: c.alignment = Alignment(horizontal="center", vertical="center")
            elif j in (3, 4): c.number_format = "#,##0"; c.alignment = Alignment(horizontal="right", vertical="center")
            elif j == 5: c.number_format = "#,##0.00"; c.alignment = Alignment(horizontal="right", vertical="center")
            elif j == 6: c.number_format = "0.0%"; c.alignment = Alignment(horizontal="right", vertical="center")
        r += 1

    r += 1
    ws1.cell(r, 1, "2. สรุปค่าใช้จ่ายรายเดือน (PM vs CM vs เปลี่ยนเครื่องยนต์)").font = f_sec
    r += 1
    m_hdr_row = r
    hdr2 = ["เดือน", "PM (ใบ)", "ยอด PM (บาท)", "% PM", "CM (ใบ)", "ยอด CM (บาท)", "% CM",
            "เปลี่ยนเครื่อง (ใบ)", "ยอดเปลี่ยนเครื่อง (บาท)", "รวมทั้งหมด (บาท)"]
    for j, h in enumerate(hdr2, 1):
        c = ws1.cell(r, j, h)
        c.font = f_hdr; c.fill = fill_hdr; c.border = b_thin
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[r].height = 24
    r += 1
    m_start = r
    for mr_row in res["monthly_rows"]:
        vals = [mr_row["เดือน"], mr_row["pm_docs"], mr_row["pm_amt"], mr_row["pm_pct"],
                mr_row["cm_docs"], mr_row["cm_amt"], mr_row["cm_pct"],
                mr_row["eng_docs"], mr_row["eng_amt"], mr_row["total_amt"]]
        for j, v in enumerate(vals, 1):
            c = ws1.cell(r, j, v)
            c.font = f_cell; c.border = b_thin
            if j == 1: c.alignment = Alignment(horizontal="center", vertical="center")
            elif j in (2, 5, 8): c.number_format = "#,##0"; c.alignment = Alignment(horizontal="right", vertical="center")
            elif j in (3, 6, 9, 10): c.number_format = "#,##0.00"; c.alignment = Alignment(horizontal="right", vertical="center")
            elif j in (4, 7): c.number_format = "0.0%"; c.alignment = Alignment(horizontal="right", vertical="center")
        r += 1
    m_end = max(m_start, r - 1)

    if res["monthly_rows"]:
        tot_m = ["รวม", f"=SUM(B{m_start}:B{m_end})", f"=SUM(C{m_start}:C{m_end})", f"=IFERROR(C{r}/J{r},0)",
                 f"=SUM(E{m_start}:E{m_end})", f"=SUM(F{m_start}:F{m_end})", f"=IFERROR(F{r}/J{r},0)",
                 f"=SUM(H{m_start}:H{m_end})", f"=SUM(I{m_start}:I{m_end})", f"=SUM(J{m_start}:J{m_end})"]
        for j, v in enumerate(tot_m, 1):
            c = ws1.cell(r, j, v)
            c.font = f_bold; c.fill = fill_tot; c.border = b_thin
            if j == 1: c.alignment = Alignment(horizontal="center", vertical="center")
            elif j in (2, 5, 8): c.number_format = "#,##0"; c.alignment = Alignment(horizontal="right", vertical="center")
            elif j in (3, 6, 9, 10): c.number_format = "#,##0.00"; c.alignment = Alignment(horizontal="right", vertical="center")
            elif j in (4, 7): c.number_format = "0.0%"; c.alignment = Alignment(horizontal="right", vertical="center")

        chart = BarChart()
        chart.type = "col"; chart.grouping = "stacked"; chart.overlap = 100
        chart.title = "ค่าใช้จ่ายรายเดือน: PM vs CM (บาท)"
        chart.y_axis.title = "บาท"; chart.x_axis.title = "เดือน"
        cats = Reference(ws1, min_col=1, min_row=m_start, max_row=m_end)
        chart.add_data(Reference(ws1, min_col=3, min_row=m_hdr_row, max_row=m_end), titles_from_data=True)
        chart.add_data(Reference(ws1, min_col=6, min_row=m_hdr_row, max_row=m_end), titles_from_data=True)
        chart.set_categories(cats)
        chart.series[0].graphicalProperties.solidFill = "2E9E6A"
        chart.series[1].graphicalProperties.solidFill = "1F6FB2"
        chart.width = 18; chart.height = 11
        ws1.add_chart(chart, "L5")

    for j, w in enumerate([14, 28, 18, 15, 18, 18, 12, 16, 20, 18], 1):
        ws1.column_dimensions[get_column_letter(j)].width = w
    ws1.freeze_panes = "A6"

    # Sheet 2: สรุปรายคัน_PM_CM
    ws2 = wb.create_sheet("สรุปรายคัน_PM_CM")
    hdr_v = ["ลำดับ", "ทะเบียน", "ยี่ห้อ", "PM ศูนย์ (ใบ)", "PM ลาน (ใบ)", "ยอด PM (บาท)", "% PM",
             "เปลี่ยนเครื่อง (ใบ)", "ยอดเปลี่ยนเครื่อง (บาท)",
             "CM W/O (ใบ)", "CM WX (ใบ)", "ยอด CM (บาท)", "% CM", "รวมทั้งหมด (บาท)"]
    for j, h in enumerate(hdr_v, 1):
        c = ws2.cell(1, j, h)
        c.font = f_hdr; c.fill = fill_hdr; c.border = b_thin
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws2.row_dimensions[1].height = 26
    for idx, v in enumerate(res["vehicle_rows"], start=1):
        r = idx + 1
        tot = round(v["pm_a"] + v["eng_a"] + v["cm_a"], 2)
        vals = [idx, v["plate"], v["brand"], v["pm_wo"], v["pm_wx"], round(v["pm_a"], 2), (v["pm_a"] / tot) if tot else 0,
                v["eng_cnt"], round(v["eng_a"], 2),
                v["cm_wo"], v["cm_wx"], round(v["cm_a"], 2), (v["cm_a"] / tot) if tot else 0, tot]
        for j, val in enumerate(vals, 1):
            c = ws2.cell(r, j, val)
            c.font = f_cell; c.border = b_thin
            if j in (1, 2, 3): c.alignment = Alignment(horizontal="center", vertical="center")
            elif j in (4, 5, 8, 10, 11): c.number_format = "#,##0"; c.alignment = Alignment(horizontal="right", vertical="center")
            elif j in (6, 9, 12, 14): c.number_format = "#,##0.00"; c.alignment = Alignment(horizontal="right", vertical="center")
            elif j in (7, 13): c.number_format = "0.0%"; c.alignment = Alignment(horizontal="right", vertical="center")
    for j, w in enumerate([8, 12, 10, 13, 13, 16, 10, 14, 18, 13, 13, 16, 10, 18], 1):
        ws2.column_dimensions[get_column_letter(j)].width = w
    ws2.freeze_panes = "D2"
    if res["vehicle_rows"]:
        ws2.auto_filter.ref = f"A1:N{len(res['vehicle_rows']) + 1}"

    # Sheet 3, 4, 5: รายการ_PM, รายการ_เปลี่ยนเครื่องยนต์, รายการ_CM
    def _write_detail(title, rows, inc_pm):
        ws = wb.create_sheet(title)
        hdr = ["ลำดับ", "เดือน", "วันที่", "ทะเบียน", "ยี่ห้อ", "กลุ่ม", "ประเภทเอกสาร", "เลขที่เอกสาร",
               "ผู้ขาย / คลัง", "เลขไมล์", "ยอดสุทธิ (บาท)", "Confirm Flag"]
        widths = [8, 11, 13, 12, 10, 16, 24, 15, 34, 12, 16, 12]
        if inc_pm:
            hdr.append("รายการที่เข้าเกณฑ์"); widths.append(40)
        for j, (h, w) in enumerate(zip(hdr, widths), 1):
            c = ws.cell(1, j, h)
            c.font = f_hdr; c.fill = fill_hdr; c.border = b_thin
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(j)].width = w
        ws.row_dimensions[1].height = 24
        for idx, x in enumerate(rows, start=1):
            r = idx + 1
            vals = [idx, x["ym"], x["doc_date"], x["plate"], x["brand"], x["maint"], x["sub"],
                    x["doc_no"], x["vendor"], x["km"], x["amount"], x["confirm_flag"]]
            if inc_pm:
                vals.append(x["pm_items"])
            ws.append(vals)
            ws.cell(r, 3).number_format = "dd/mm/yyyy"
            ws.cell(r, 10).number_format = "#,##0"
            ws.cell(r, 11).number_format = "#,##0.00"
        ws.freeze_panes = "E2"
        if rows:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(hdr))}{len(rows) + 1}"

    pm_rows = [x for x in docs if x["maint"] == "PM"]
    eng_rows = [x for x in docs if x["maint"] == "เปลี่ยนเครื่องยนต์"]
    cm_rows = [x for x in docs if x["maint"] == "CM"]
    _write_detail("รายการ_PM", pm_rows, True)
    _write_detail("รายการ_เปลี่ยนเครื่องยนต์", eng_rows, True)
    _write_detail("รายการ_CM", cm_rows, False)

    # Sheet 5: รายละเอียดอ้างอิง (ทุกบรรทัด WX + W/O)
    ws5 = wb.create_sheet("รายละเอียดอ้างอิง")
    hdr_ref = [
        "ลำดับ", "เดือน", "วันที่", "ทะเบียน", "ยี่ห้อ", "กลุ่ม", "ประเภทเอกสาร", "ชนิด",
        "เลขที่เอกสาร", "บรรทัด", "รหัสรายการ/พัสดุ", "ชื่อรายการ / อะไหล่",
        "จำนวน", "ราคา/หน่วย", "ยอดสุทธิ (บาท)", "ผู้ขาย / คลัง", "เลขไมล์", "Confirm Flag", "ไฟล์ต้นทาง",
    ]
    w_ref = [8, 10, 12, 11, 9, 7, 23, 7, 14, 8, 18, 42, 10, 13, 15, 32, 11, 12, 16]
    ws5.row_dimensions[1].height = 24
    for j, (h, w) in enumerate(zip(hdr_ref, w_ref), 1):
        c = ws5.cell(1, j, h)
        c.font = f_hdr; c.fill = fill_hdr; c.border = b_thin
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws5.column_dimensions[get_column_letter(j)].width = w
    for idx, row in enumerate(ref_rows, start=1):
        r = idx + 1
        ws5.append([idx] + list(row))
        ws5.cell(r, 3).number_format = "dd/mm/yyyy"
        ws5.cell(r, 13).number_format = "#,##0.##"
        ws5.cell(r, 14).number_format = "#,##0.00"
        ws5.cell(r, 15).number_format = "#,##0.00"
        ws5.cell(r, 17).number_format = "#,##0"
    ws5.freeze_panes = "E2"
    if ref_rows:
        ws5.auto_filter.ref = f"A1:{get_column_letter(len(hdr_ref))}{len(ref_rows) + 1}"

    # Sheet 6: เกณฑ์การจำแนก
    ws6 = wb.create_sheet("เกณฑ์การจำแนก")
    notes = [
        ("ช่วงวันที่ข้อมูล", dr_txt),
        ("เกณฑ์ PM ศูนย์บริการ (W/O)", "ใบสั่งซ่อม (W/O) ที่มีรายการ 'น้ำมันเครื่อง' / 'ENGINE OIL' หรือ 'ไส้กรอง/กรองน้ำมันเครื่อง' (ไม่นับ ปะเก็น/ท่อ/ฝา/ถัง/อ่าง/ซีล และไม่นับใบเปลี่ยนเครื่องยนต์)"),
        ("เกณฑ์ PM ลาน (WX)", "ใบเบิกพัสดุ (WX) ที่มีรายการตัว 'น้ำมันเครื่อง' (ไม่นับไส้กรองและอะไหล่ประกอบ)"),
        ("เกณฑ์ เปลี่ยนเครื่องยนต์ (W/O)", "ใบสั่งซ่อม (W/O) ที่มีรายการเปลี่ยนเครื่องยนต์ดีเซล (WO-001-001-0198 HINO DIESEL ENGINE / WO-001-001-0202 KIT CHANGE COMPANANT NGV-DIESEL) — ใช้เป็นจุดเริ่มต้นวางแผน PM รอบ 20,000 กม. / 6 เดือน"),
        ("เกณฑ์ CM (ซ่อมทั่วไป)", "ใบสั่งซ่อม (W/O) และใบเบิกพัสดุ (WX) ทั้งหมดที่ไม่เข้าเกณฑ์ PM และเปลี่ยนเครื่องยนต์ข้างต้น"),
        ("หมายเหตุการคำนวณ", "แยกเลขที่เอกสาร W/O กับ WX ออกจากกันชัดเจน (ไม่ซ้อนกันกรณีเลขเอกสาร 8 หลักเหมือนกัน)"),
        ("ยอดกระทบยอดรวมทั้งหมด", f"{len(docs):,} เอกสาร (PM {len(pm_rows):,} / เปลี่ยนเครื่องยนต์ {len(eng_rows):,} / CM {len(cm_rows):,}) · รวม {grand_amt:,.2f} บาท · รายละเอียดอ้างอิง {len(ref_rows):,} บรรทัด"),
    ]
    ws6.cell(1, 1, "หัวข้อ").font = f_hdr; ws6.cell(1, 1).fill = fill_hdr; ws6.cell(1, 1).border = b_thin
    ws6.cell(1, 2, "รายละเอียด").font = f_hdr; ws6.cell(1, 2).fill = fill_hdr; ws6.cell(1, 2).border = b_thin
    for i, (k, v) in enumerate(notes, start=2):
        ws6.cell(i, 1, k).font = f_bold; ws6.cell(i, 1).border = b_thin
        ws6.cell(i, 2, v).font = f_cell; ws6.cell(i, 2).border = b_thin
    ws6.column_dimensions["A"].width = 28
    ws6.column_dimensions["B"].width = 95

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), res


# ============================================================================= Streamlit UI helpers
def render_vr_save_box(db_path, ds, user, full):
    """กล่องบันทึกข้อมูลที่อ่านได้จาก Crystal Reports ลง SQLite (กันซ้ำด้วย UNIQUE INDEX + สำรองอัตโนมัติ)"""
    import streamlit as st
    if not db_path or ds is None:
        return
    with st.expander("💾 บันทึกลงฐานข้อมูล SQLite (แยกตาราง vr_issue_line / vr_wo_header / vr_wo_line — กันซ้ำด้วย UNIQUE INDEX)", expanded=True):
        plan = analyze_dataset(db_path, ds)
        n_new = len(plan["issues"]["new"]) + len(plan["wo_hdr"]["new"]) + len(plan["wo_lines"]["new"])
        n_upd = len(plan["issues"]["update"]) + len(plan["wo_hdr"]["update"]) + len(plan["wo_lines"]["update"])
        n_same = len(plan["issues"]["unchanged"]) + len(plan["wo_hdr"]["unchanged"]) + len(plan["wo_lines"]["unchanged"])
        st.caption(
            f"**ฐานข้อมูลปลายทาง:** `{db_path}`  \n"
            f"- **เบิกพัสดุ (WX):** ใหม่ `{len(plan['issues']['new']):,}` · อัปเดต `{len(plan['issues']['update']):,}` · เหมือนเดิม `{len(plan['issues']['unchanged']):,}`  \n"
            f"- **ใบสั่งซ่อม_หัว (W/O):** ใหม่ `{len(plan['wo_hdr']['new']):,}` · อัปเดต `{len(plan['wo_hdr']['update']):,}` · เหมือนเดิม `{len(plan['wo_hdr']['unchanged']):,}`  \n"
            f"- **ใบสั่งซ่อม_บรรทัด:** ใหม่ `{len(plan['wo_lines']['new']):,}` · อัปเดต `{len(plan['wo_lines']['update']):,}` · เหมือนเดิม `{len(plan['wo_lines']['unchanged']):,}`"
        )
        c1, c2, c3 = st.columns([2.2, 2.2, 1.6])
        sync_pm = c1.checkbox("ซิงก์เข้าประวัติบริการ PM Plan (service_doc) พร้อมกัน", value=True, key="vr_sync_pm")
        confirm = c2.checkbox("ยืนยันบันทึก (สำรอง .sqlite อัตโนมัติก่อนบันทึก)", value=False, key="vr_save_ok")
        disabled = (not confirm) or (n_new == 0 and n_upd == 0 and not sync_pm)
        if c3.button("💾 บันทึกลงฐานข้อมูล", type="primary", disabled=disabled, key="vr_save_go", **full):
            with st.spinner("กำลังสำรองฐานข้อมูลและบันทึกรายงานซ่อมรายคัน…"):
                res = commit_dataset(db_path, ds, user=user, sync_pm=sync_pm)
            st.success(
                f"บันทึกสำเร็จ (Batch #{res['batch_id']}) · "
                f"WX +{res['issues']['new']:,}/{res['issues']['update']:,} · "
                f"W/O หัว +{res['wo_hdr']['new']:,}/{res['wo_hdr']['update']:,} · "
                f"W/O บรรทัด +{res['wo_lines']['new']:,}/{res['wo_lines']['update']:,}"
                + (f" · ซิงก์ PM +{res['pm_sync']['inserted']:,}/{res['pm_sync']['updated']:,}" if sync_pm else "")
            )
            if res.get("backup"):
                st.caption(f"สำรองฐานข้อมูลไว้ที่: `{res['backup']}`")


def render_vr_db_tab(db_path, is_local, app_root, full, kpi_fn, money_fn):
    """แท็บ '🗄️ ข้อมูลในฐาน (รายงานรายคัน)' — ค้นหา/กรอง/Export ย้อนหลังจาก SQLite"""
    import pandas as pd
    import streamlit as st
    import monthly_report as mr

    if not db_path or not os.path.exists(db_path):
        st.info("ยังไม่มีไฟล์ฐานข้อมูล — เลือกไฟล์ `.sqlite` ทางซ้าย แล้วอ่านรายงานและกดบันทึกเพื่อเริ่มเก็บข้อมูล")
        return

    sm = db_summary(db_path)
    if not sm or (sm["n_issue_lines"] == 0 and sm["n_wo_hdr"] == 0 and sm["n_wo_lines"] == 0):
        st.info("ยังไม่มีข้อมูลรายงานซ่อมรายคันในฐานข้อมูลนี้ — อ่านไฟล์รายงานทางซ้ายแล้วกดปุ่ม **💾 บันทึกลงฐานข้อมูล** ด้านบน")
        return

    def _fmt_d(s):
        if not s:
            return "-"
        try:
            return dt.date.fromisoformat(str(s)[:10]).strftime("%d/%m/%Y")
        except Exception:
            return str(s)

    k = st.columns(5)
    dr = f"{_fmt_d(sm['date_min'])} – {_fmt_d(sm['date_max'])}" if sm["date_min"] else "-"
    kpi_fn(k[0], "ทะเบียนในฐาน", f"{sm['n_plates']:,}", f"ช่วง {dr}")
    kpi_fn(k[1], "เบิกพัสดุ (WX)", money_fn(sm["wx_total"]), f"{sm['n_issue_docs']:,} ใบ · {sm['n_issue_lines']:,} บรรทัด")
    kpi_fn(k[2], "สั่งซ่อม (W/O)", money_fn(sm["wo_total"]), f"{sm['n_wo_hdr']:,} ใบ · {sm['n_wo_lines']:,} บรรทัด")
    kpi_fn(k[3], "ยอดรวมในฐาน (บาท)", money_fn(sm["grand_total"]), "WX + W/O", "ok")
    kpi_fn(k[4], "รอบการนำเข้า", f"{len(sm['batches']):,}", "UNIQUE INDEX กันซ้ำ")
    st.write("")

    f1, f2, f3, f4 = st.columns([1.6, 2.0, 2.0, 1.2])
    sel_m = f1.multiselect("เดือน", sm["months"], key="vrdb_m", placeholder="ทุกเดือน")
    sel_p = f2.multiselect("ทะเบียน", sm["plates"], key="vrdb_p", placeholder="ทุกทะเบียน")
    q_txt = f3.text_input("ค้นหา", key="vrdb_q", placeholder="เลขเอกสาร / อะไหล่ / ผู้ขาย / หมายเหตุ…")
    flag_only = f4.toggle("เฉพาะ Flag No", key="vrdb_flag")

    c_ex, c_sp = st.columns([1.8, 4.2])
    if c_ex.button("📤 สร้างไฟล์ Export จากฐาน (Excel)", key="vrdb_exp_btn", **full):
        with st.spinner("กำลังสร้างไฟล์ Excel จากฐานข้อมูล…"):
            blob = export_db_excel(db_path, months=sel_m or None, plates=sel_p or None, flag_only=flag_only, search=q_txt)
        fn = f"รายงานซ่อมรายคัน_DB_{dt.datetime.now():%Y%m%d_%H%M}.xlsx"
        saved = mr.save_local_export(app_root, fn, blob, is_local=is_local)
        st.session_state["vrdb_exp_blob"] = (fn, blob, saved)
    if st.session_state.get("vrdb_exp_blob"):
        fn, blob, saved = st.session_state["vrdb_exp_blob"]
        if saved:
            c_sp.success(f"💾 บันทึกแล้วที่: `{saved}`")
        st.download_button(
            f"⬇️ ดาวน์โหลด {fn}", blob, file_name=fn,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="vrdb_exp_dl",
        )

    st0, st1, st2, st3, st4 = st.tabs(["📊 สรุป PM / CM", "เบิกพัสดุ (WX)", "ใบสั่งซ่อม_หัว (W/O)", "ใบสั่งซ่อม_บรรทัด", "🕘 ประวัติการนำเข้า"])
    with st0:
        pmcm = compute_pm_cm(db_path, months=sel_m or None, plates=sel_p or None)
        ga = pmcm["grand_amt"] or 0.0
        pm_pct = (pmcm["pm_amt"] * 100.0 / ga) if ga else 0.0
        eng_pct = (pmcm["eng_amt"] * 100.0 / ga) if ga else 0.0
        cm_pct = (pmcm["cm_amt"] * 100.0 / ga) if ga else 0.0
        pk = st.columns(4)
        kpi_fn(pk[0], "ยอด PM (บาท)", money_fn(pmcm["pm_amt"]), f"{pmcm['pm_cnt']:,} ใบ · {pm_pct:.1f}%", "ok")
        kpi_fn(pk[1], "ยอด CM ซ่อมทั่วไป (บาท)", money_fn(pmcm["cm_amt"]), f"{pmcm['cm_cnt']:,} ใบ · {cm_pct:.1f}%", "blue")
        kpi_fn(pk[2], "เปลี่ยนเครื่องยนต์ (บาท)", money_fn(pmcm["eng_amt"]), f"{pmcm['eng_cnt']:,} คัน · {eng_pct:.1f}% (จุดเริ่มวาง PM)", "warn")
        kpi_fn(pk[3], "รวมทั้งหมด (บาท)", money_fn(ga), f"{pmcm['pm_cnt'] + pmcm['eng_cnt'] + pmcm['cm_cnt']:,} ใบ · {len(pmcm['ref_rows']):,} บรรทัด")
        st.write("")

        ep1, ep2 = st.columns([2.2, 3.8])
        if ep1.button("📊 Export สรุป PM / CM / เปลี่ยนเครื่องยนต์ (Excel 7 ชีต)", type="primary", key="vrdb_pmcm_exp_btn", **full):
            with st.spinner("กำลังสร้างไฟล์ Excel สรุป PM / CM / เปลี่ยนเครื่องยนต์ ครบ 7 ชีต…"):
                pblob, _ = export_pm_cm_excel(db_path, months=sel_m or None, plates=sel_p or None)
            pfn = f"สรุป_PM_vs_CM_{dt.datetime.now():%Y%m%d_%H%M}.xlsx"
            psaved = mr.save_local_export(app_root, pfn, pblob, is_local=is_local)
            st.session_state["vrdb_pmcm_blob"] = (pfn, pblob, psaved)
        if st.session_state.get("vrdb_pmcm_blob"):
            pfn, pblob, psaved = st.session_state["vrdb_pmcm_blob"]
            if psaved:
                ep2.success(f"💾 บันทึกแล้วที่: `{psaved}`")
            st.download_button(
                f"⬇️ ดาวน์โหลด {pfn}", pblob, file_name=pfn, type="primary",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="vrdb_pmcm_dl",
            )

        if pmcm["monthly_rows"]:
            import altair as alt
            mdf_chart = pd.DataFrame([
                {"เดือน": r["เดือน"], "กลุ่ม": "PM", "ยอดเงิน": r["pm_amt"]} for r in pmcm["monthly_rows"]
            ] + [
                {"เดือน": r["เดือน"], "กลุ่ม": "CM", "ยอดเงิน": r["cm_amt"]} for r in pmcm["monthly_rows"]
            ])
            ch = (
                alt.Chart(mdf_chart)
                .mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
                .encode(
                    x=alt.X("เดือน:N", title=None, axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("sum(ยอดเงิน):Q", title="บาท", axis=alt.Axis(format=",.0f")),
                    color=alt.Color("กลุ่ม:N", scale=alt.Scale(domain=["PM", "CM"], range=["#2e9e6a", "#1f6fb2"]),
                                    legend=alt.Legend(orient="bottom", title=None)),
                    tooltip=["เดือน", "กลุ่ม", alt.Tooltip("ยอดเงิน:Q", title="บาท", format=",.2f")],
                )
                .properties(height=300)
            )
            st.markdown("##### ค่าใช้จ่ายรายเดือน: PM vs CM (บาท)")
            st.altair_chart(ch, **full)

        c_s1, c_s2 = st.columns([1.0, 1.35])
        with c_s1:
            st.markdown("**1. ภาพรวมแยกตามประเภทเอกสาร (PM / เปลี่ยนเครื่องยนต์ / CM)**")
            sdf = pd.DataFrame(pmcm["summary_rows"])
            if len(sdf):
                sdf["สัดส่วน (%)"] = (sdf["สัดส่วน"] * 100.0).round(1)
                st.dataframe(
                    sdf[["กลุ่ม", "ประเภทเอกสาร", "จำนวนเอกสาร", "จำนวนรถ", "ยอดเงินรวม", "สัดส่วน (%)"]],
                    hide_index=True, **full,
                    column_config={
                        "ยอดเงินรวม": st.column_config.NumberColumn(format="%,.2f"),
                        "สัดส่วน (%)": st.column_config.NumberColumn(format="%.1f%%"),
                    },
                )
        with c_s2:
            st.markdown("**2. สรุปรายเดือน (PM vs CM vs เปลี่ยนเครื่องยนต์)**")
            mdf = pd.DataFrame([
                {
                    "เดือน": r["เดือน"], "PM (ใบ)": r["pm_docs"], "ยอด PM": r["pm_amt"],
                    "% PM": round(r["pm_pct"] * 100, 1),
                    "CM (ใบ)": r["cm_docs"], "ยอด CM": r["cm_amt"],
                    "% CM": round(r["cm_pct"] * 100, 1),
                    "เปลี่ยนเครื่อง (ใบ)": r["eng_docs"], "ยอดเปลี่ยนเครื่อง": r["eng_amt"],
                    "รวม (บาท)": r["total_amt"],
                }
                for r in pmcm["monthly_rows"]
            ])
            if len(mdf):
                st.dataframe(
                    mdf, hide_index=True, **full,
                    column_config={
                        "ยอด PM": st.column_config.NumberColumn(format="%,.2f"),
                        "ยอด CM": st.column_config.NumberColumn(format="%,.2f"),
                        "ยอดเปลี่ยนเครื่อง": st.column_config.NumberColumn(format="%,.2f"),
                        "รวม (บาท)": st.column_config.NumberColumn(format="%,.2f"),
                        "% PM": st.column_config.NumberColumn(format="%.1f%%"),
                        "% CM": st.column_config.NumberColumn(format="%.1f%%"),
                    },
                )

        st.markdown("**3. สรุปรายคัน (PM / เปลี่ยนเครื่องยนต์ / CM — เรียงตามยอดรวมสูงสุด)**")
        vdf = pd.DataFrame([
            {
                "ทะเบียน": v["plate"], "ยี่ห้อ": v["brand"],
                "PM ศูนย์ (ใบ)": v["pm_wo"], "PM ลาน (ใบ)": v["pm_wx"],
                "ยอด PM": round(v["pm_a"], 2),
                "เปลี่ยนเครื่อง (ใบ)": v["eng_cnt"], "ยอดเปลี่ยนเครื่อง": round(v["eng_a"], 2),
                "CM W/O (ใบ)": v["cm_wo"], "CM WX (ใบ)": v["cm_wx"],
                "ยอด CM": round(v["cm_a"], 2),
                "รวมทั้งหมด": round(v["pm_a"] + v["eng_a"] + v["cm_a"], 2),
            }
            for v in pmcm["vehicle_rows"]
        ])
        if len(vdf):
            st.dataframe(
                vdf, hide_index=True, height=380, **full,
                column_config={
                    "ยอด PM": st.column_config.NumberColumn(format="%,.2f"),
                    "ยอดเปลี่ยนเครื่อง": st.column_config.NumberColumn(format="%,.2f"),
                    "ยอด CM": st.column_config.NumberColumn(format="%,.2f"),
                    "รวมทั้งหมด": st.column_config.NumberColumn(format="%,.2f"),
                },
            )
    def _fmt_df_dates(df_):
        if "doc_date" in df_.columns:
            df_ = df_.copy()
            df_["doc_date"] = pd.to_datetime(df_["doc_date"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
        return df_

    _date_cfg = {"doc_date": st.column_config.TextColumn("วันที่")}

    with st1:
        rows = query_rows(db_path, "issues", sel_m or None, sel_p or None, flag_only, q_txt)
        df = _fmt_df_dates(pd.DataFrame(rows))
        tot = sum(float(r.get("total") or 0) for r in rows)
        st.caption(f"แสดง {len(rows):,} บรรทัด · ยอดรวม {tot:,.2f} บาท")
        if len(df):
            show = [c for c in ("doc_date", "plate", "doc_no", "line_no", "inv_code", "inv_name", "qty", "unit_price", "total", "store", "mileage_km", "confirm_flag") if c in df.columns]
            st.dataframe(df[show], hide_index=True, height=460, column_config=_date_cfg, **full)
    with st2:
        rows = query_rows(db_path, "wo_hdr", sel_m or None, sel_p or None, flag_only, q_txt)
        df = _fmt_df_dates(pd.DataFrame(rows))
        tot = sum(float(r.get("net") or 0) for r in rows)
        st.caption(f"แสดง {len(rows):,} ใบ · ยอดสุทธิรวม {tot:,.2f} บาท")
        if len(df):
            show = [c for c in ("doc_date", "plate", "doc_no", "vendor_th", "vendor_en", "amount", "discount", "net", "n_lines", "mileage_km", "confirm_flag", "flag_reason") if c in df.columns]
            st.dataframe(df[show], hide_index=True, height=460, column_config=_date_cfg, **full)
    with st3:
        rows = query_rows(db_path, "wo_lines", sel_m or None, sel_p or None, flag_only, q_txt)
        df = _fmt_df_dates(pd.DataFrame(rows))
        tot = sum(float(r.get("total") or 0) for r in rows)
        st.caption(f"แสดง {len(rows):,} บรรทัด · ยอดรวม {tot:,.2f} บาท")
        if len(df):
            show = [c for c in ("doc_date", "plate", "doc_no", "line_no", "item_code", "item_name", "qty", "unit_price", "total", "vendor_en", "confirm_flag") if c in df.columns]
            st.dataframe(df[show], hide_index=True, height=460, column_config=_date_cfg, **full)
    with st4:
        bdf = pd.DataFrame(sm["batches"])
        if len(bdf):
            if "imported_at" in bdf.columns:
                bdf["imported_at"] = pd.to_datetime(bdf["imported_at"], errors="coerce").dt.strftime("%d/%m/%Y %H:%M").fillna(bdf["imported_at"])
            cols = [c for c in ("batch_id", "imported_at", "imported_by", "source_desc", "files_used",
                                "issues_inserted", "issues_updated", "wo_hdr_inserted", "wo_hdr_updated",
                                "wo_lines_inserted", "wo_lines_updated", "pm_synced") if c in bdf.columns]
            st.dataframe(bdf[cols], hide_index=True, **full)

