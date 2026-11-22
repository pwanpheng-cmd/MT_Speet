# -*- coding: utf-8 -*-
"""
Repair Request DB engine v0.2.1 — ใบแจ้งซ่อม (หัวลาก / หางพ่วง / ตู้เย็น) -> SQLite
- อ่านไฟล์ได้ทั้งแบบ M1_ซ่อมรถ/M2_ซ่อมหาง/M3_ซ่อมตู้ และแบบรายเดือน หัวลาก/หางพ่วง/ตู้เย็น (หาแถวหัวตารางเอง)
- นำเข้าเฉพาะคอลัมน์ข้อมูลจริง: เลขที่เอกสาร, วันที่เปิดเอกสาร, วันที่แจ้งซ่อม, เวลา, ทะเบียน/เบอร์ตู้, รหัสพนักงาน, สถานะซ่อม
  (คอลัมน์ทะเบียน/ยี่ห้อ/เดือน/Index/ISO ในไฟล์เป็นสูตร -> คำนวณเองในระบบ)
- Business key = doc_no (UNIQUE) ; แถวที่ข้อมูลหลักไม่สอดคล้อง -> repair_request_review รอ DATA_OWNER (ไม่เข้าตารางหลัก)
"""
import os, re, json, shutil, sqlite3, hashlib, datetime as dt

ENGINE_VERSION = "0.2.1"

SCHEMA_SQL = """
PRAGMA foreign_keys = ON;
CREATE TABLE IF NOT EXISTS schema_info (key TEXT PRIMARY KEY, value TEXT);
INSERT OR IGNORE INTO schema_info VALUES ('schema_version','0.2.1');

CREATE TABLE IF NOT EXISTS import_batch (
    batch_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    file_name       TEXT NOT NULL,
    file_sha256     TEXT NOT NULL,
    imported_by     TEXT,
    imported_at     TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    is_base_load    INTEGER NOT NULL DEFAULT 0,
    rows_read       INTEGER NOT NULL DEFAULT 0,
    rows_blank      INTEGER NOT NULL DEFAULT 0,
    rows_dup_in_file INTEGER NOT NULL DEFAULT 0,
    rows_inserted   INTEGER NOT NULL DEFAULT 0,
    rows_unchanged  INTEGER NOT NULL DEFAULT 0,
    rows_updated    INTEGER NOT NULL DEFAULT 0,
    rows_review     INTEGER NOT NULL DEFAULT 0,
    backup_path     TEXT,
    note            TEXT
);

CREATE TABLE IF NOT EXISTS repair_request (
    request_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    doc_no          TEXT NOT NULL,
    asset_type      TEXT NOT NULL CHECK (asset_type IN ('หัวลาก','หางพ่วง','ตู้เย็น')),
    source_sheet    TEXT NOT NULL,
    open_date       TEXT,
    report_date     TEXT NOT NULL,
    report_month    TEXT NOT NULL,
    report_time_raw TEXT,
    report_time     TEXT,
    asset_raw       TEXT NOT NULL,
    asset_code      TEXT NOT NULL,
    brand           TEXT,
    employee_code   TEXT,
    status          TEXT NOT NULL CHECK (status IN ('Approved','Open','Cancel')),
    row_hash        TEXT NOT NULL,
    confirm_flag    TEXT NOT NULL DEFAULT 'Yes' CHECK (confirm_flag IN ('Yes','No')),
    flag_reason     TEXT,
    data_note       TEXT,
    source_row      INTEGER,
    first_batch_id  INTEGER NOT NULL REFERENCES import_batch(batch_id),
    last_batch_id   INTEGER NOT NULL REFERENCES import_batch(batch_id),
    created_at      TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    updated_at      TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE UNIQUE INDEX IF NOT EXISTS ux_repair_request_doc_no ON repair_request(doc_no);
CREATE INDEX IF NOT EXISTS ix_rr_kpi   ON repair_request(report_month, asset_type, asset_code);
CREATE INDEX IF NOT EXISTS ix_rr_asset ON repair_request(asset_code, report_date);

CREATE TABLE IF NOT EXISTS repair_request_history (
    hist_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    doc_no      TEXT NOT NULL,
    batch_id    INTEGER REFERENCES import_batch(batch_id),
    field       TEXT NOT NULL,
    old_value   TEXT,
    new_value   TEXT,
    changed_by  TEXT,
    changed_at  TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS repair_request_review (
    review_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id      INTEGER NOT NULL REFERENCES import_batch(batch_id),
    doc_no        TEXT NOT NULL,
    kind          TEXT NOT NULL CHECK (kind IN ('ข้อมูลไม่ผ่านตรวจ','ฟิลด์หลักเปลี่ยน','ซ้ำในไฟล์ค่าต่างกัน')),
    reason        TEXT NOT NULL,
    incoming_json TEXT NOT NULL,
    incoming_hash TEXT,
    decision      TEXT CHECK (decision IN ('นำเข้า','ไม่นำเข้า') OR decision IS NULL),
    decided_by    TEXT,
    decided_at    TEXT,
    created_at    TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS ix_review_open ON repair_request_review(decision);
CREATE INDEX IF NOT EXISTS ix_review_hash ON repair_request_review(doc_no, incoming_hash);

CREATE VIEW IF NOT EXISTS v_kpi_asset_month AS
SELECT report_month, asset_type, asset_code, COUNT(*) AS n_requests,
       MIN(report_date) AS first_report, MAX(report_date) AS last_report
FROM repair_request GROUP BY report_month, asset_type, asset_code;
"""

STATUS = {"Approved", "Open", "Cancel"}
CORE = ("asset_type", "report_date", "asset_code")
SOFT = ("status", "report_time_raw", "report_time", "employee_code", "open_date", "asset_raw", "brand",
        "source_sheet", "data_note")
FIELDS = ("doc_no", "asset_type", "source_sheet", "open_date", "report_date", "report_month", "report_time_raw",
          "report_time", "asset_raw", "asset_code", "brand", "employee_code", "status", "row_hash",
          "confirm_flag", "flag_reason", "data_note", "source_row")
HEADERS = {"doc_no": "เลขที่เอกสาร", "open_date": "วันที่เปิดเอกสาร", "report_date": "วันที่แจ้งซ่อม",
           "time": "เวลา", "asset_raw": "ทะเบียน / เบอร์ตู้", "employee": "รหัสพนักงาน", "status": "สถานะซ่อม"}


def sheet_asset_type(name):
    n = name.replace(" ", "")
    if "M1" in n or "ซ่อมรถ" in n or "หัวลาก" in n:
        return "หัวลาก"
    if "M2" in n or "ซ่อมหาง" in n or "หางพ่วง" in n or "ห่างพ่วง" in n:
        return "หางพ่วง"
    if "M3" in n or "ซ่อมตู้" in n or "ตู้เย็น" in n:
        return "ตู้เย็น"
    return None


def _d(v):
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    return None


def _norm_time(v):
    """คืน (raw_text, HH:MM | None, note | None) ; ค่าวันที่ปี 1900 = ตัวเลขที่ Excel จัดรูปแบบเป็นวันที่ -> แปลงกลับ"""
    note = None
    if isinstance(v, dt.datetime) and v.year in (1899, 1900):
        # Excel serial: ก่อน 1 มี.ค. 1900 ใช้ฐาน 31/12/1899 (บั๊กปีอธิกสุรทิน 1900 ของ Excel), หลังจากนั้นฐาน 30/12/1899
        base = dt.datetime(1899, 12, 31) if v < dt.datetime(1900, 3, 1) else dt.datetime(1899, 12, 30)
        num = round((v - base).total_seconds() / 86400, 4)
        note = f"เวลาในไฟล์เป็นรูปแบบวันที่ ({v:%d/%m/%Y %H:%M}) แปลงกลับเป็นตัวเลข {num:g}"
        v = num
    if v in (None, "", 0):
        return (None if v in (None, "") else "0"), None, note
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        raw = f"{v:g}"
        h = int(v)
        m = round((v - h) * 100)
        if 0 <= h < 24 and 0 <= m < 60:
            return raw, f"{h:02d}:{m:02d}", note
        return raw, None, ((note + " ; ") if note else "") + f"เวลาอ่านไม่ได้: {raw} (เก็บค่าต้นฉบับ ไม่เดา)"
    return str(v), None, ((note + " ; ") if note else "") + f"เวลาอ่านไม่ได้: {v}"


def _row_hash(rec):
    keys = ("doc_no", "asset_type", "open_date", "report_date", "report_time_raw", "asset_raw", "asset_code",
            "employee_code", "status")
    return hashlib.sha1(json.dumps([rec[k] for k in keys], ensure_ascii=False).encode("utf-8")).hexdigest()


def parse_row(sheet, asset_type, rown, get):
    raw = str(get("asset_raw") or "").strip()
    code = raw.split("/")[-1].strip() if asset_type == "ตู้เย็น" else raw[-7:].strip()
    flags, notes = [], []
    if asset_type == "ตู้เย็น":
        if not re.fullmatch(r"[A-Z]{4}\d{7}", code):
            flags.append(f"เบอร์ตู้ผิดรูปแบบ: {code or '(ว่าง)'}")
    elif not re.fullmatch(r"\d{2}-\d{4}", code):
        flags.append(f"ทะเบียนผิดรูปแบบ: {code or '(ว่าง)'}")
    rd = _d(get("report_date"))
    if not rd:
        flags.append("ไม่มีวันที่แจ้งซ่อม")
    status = str(get("status") or "").strip()
    if status not in STATUS:
        flags.append(f"สถานะไม่รู้จัก: {status or '(ว่าง)'}")
    t_raw, t, t_note = _norm_time(get("time"))
    if t_note:
        notes.append(t_note)
    emp = get("employee")
    emp = None if emp in (None, "", ".") else (str(int(emp)) if isinstance(emp, float) and emp.is_integer() else str(emp).strip())
    rec = dict(doc_no=str(get("doc_no")).strip(), asset_type=asset_type, source_sheet=sheet,
               open_date=_d(get("open_date")), report_date=rd, report_month=rd[:7] if rd else None,
               report_time_raw=t_raw, report_time=t, asset_raw=raw, asset_code=code,
               brand=(raw.split("/")[0].strip() or None) if raw else None, employee_code=emp, status=status,
               confirm_flag="No" if flags else "Yes", flag_reason="; ".join(flags) or None,
               data_note="; ".join(notes) or None, source_row=rown)
    rec["row_hash"] = _row_hash(rec)
    return rec


def read_workbook(src):
    """src = path หรือ file-like ; คืน dict(records, sheets:[...], errors:[...])"""
    import openpyxl
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    out = dict(records=[], sheets=[], errors=[])
    for ws in wb.worksheets:
        at = sheet_asset_type(ws.title)
        if not at:
            continue
        rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row or 0, 200000), values_only=True))
        hdr_i, colmap = None, {}
        for i, r in enumerate(rows[:6]):
            vals = [str(v).strip() if v is not None else "" for v in r]
            if HEADERS["doc_no"] in vals:
                hdr_i = i
                for key, h in HEADERS.items():
                    if h in vals:
                        colmap[key] = vals.index(h)
                break
        if hdr_i is None:
            continue
        missing = [HEADERS[k] for k in HEADERS if k not in colmap]
        if missing:
            out["errors"].append(f"ชีต {ws.title}: ไม่พบคอลัมน์ {', '.join(missing)}")
            continue
        n_data = n_blank = 0
        for j, r in enumerate(rows[hdr_i + 1:], start=hdr_i + 2):
            get = lambda k, r=r: r[colmap[k]] if colmap[k] < len(r) else None
            if get("doc_no") in (None, ""):
                n_blank += 1
                continue
            n_data += 1
            out["records"].append(parse_row(ws.title, at, j, get))
        out["sheets"].append(dict(sheet=ws.title, asset_type=at, header_row=hdr_i + 1, rows=n_data, blank=n_blank))
    if not out["sheets"] and not out["errors"]:
        out["errors"].append("ไม่พบชีตใบแจ้งซ่อม (ต้องมีชื่อ หัวลาก/หางพ่วง/ตู้เย็น หรือ M1/M2/M3 และมีหัวคอลัมน์ 'เลขที่เอกสาร')")
    return out


# ----------------------------------------------------------------------------- DB
def connect(db_path):
    os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True)
    con = sqlite3.connect(db_path)
    con.row_factory = sqlite3.Row
    con.executescript(SCHEMA_SQL)
    return con


def analyze(con, records):
    """จำลองการนำเข้า (ยังไม่เขียน) -> plan"""
    plan = dict(new=[], unchanged=[], update=[], review=[], review_known=[], dup_in_file=[])
    seen = {}
    by_doc = {}
    for rec in records:
        by_doc.setdefault(rec["doc_no"], []).append(rec)
    conflict_docs = {d for d, l in by_doc.items() if len({r["row_hash"] for r in l}) > 1}
    for rec in records:
        k = rec["doc_no"]
        if k in seen:
            if k not in conflict_docs:
                plan["dup_in_file"].append(rec)
            continue
        seen[k] = True
        if k in conflict_docs:
            variants = by_doc[k]
            ih = hashlib.sha1("|".join(sorted(v["row_hash"] for v in variants)).encode()).hexdigest()
            if _review_exists(con, k, ih):
                plan["review_known"].append(rec)
                continue
            diff = sorted({f for f in ("asset_code", "report_date", "report_time_raw", "status", "employee_code", "open_date")
                           if len({str(v[f]) for v in variants}) > 1})
            plan["review"].append(dict(kind="ซ้ำในไฟล์ค่าต่างกัน", rec=rec, ihash=ih,
                                       reason=f"เลขเอกสารซ้ำ {len(variants)} แถวในไฟล์ ค่าต่างกันที่: {', '.join(diff)}",
                                       variants=variants))
            continue
        if _review_exists(con, k, rec["row_hash"]):
            plan["review_known"].append(rec)        # เคยส่งตรวจแล้ว (รอตัดสิน หรือ ตัดสินไม่นำเข้า) -> ไม่ซ้ำ
            continue
        old = con.execute("SELECT * FROM repair_request WHERE doc_no=?", (k,)).fetchone()
        if old is None:
            if rec["confirm_flag"] == "No":
                plan["review"].append(dict(kind="ข้อมูลไม่ผ่านตรวจ", rec=rec, reason=rec["flag_reason"]))
            else:
                plan["new"].append(rec)
            continue
        old = dict(old)
        if old["row_hash"] == rec["row_hash"]:
            plan["unchanged"].append(rec)
            continue
        core = [c for c in CORE if old[c] != rec[c]]
        if core or rec["confirm_flag"] == "No":
            why = ("ฟิลด์หลักเปลี่ยน: " + ", ".join(f"{c} {old[c]} → {rec[c]}" for c in core)) if core else rec["flag_reason"]
            plan["review"].append(dict(kind="ฟิลด์หลักเปลี่ยน" if core else "ข้อมูลไม่ผ่านตรวจ", rec=rec, reason=why))
            continue
        changes = {c: (old[c], rec[c]) for c in SOFT if old[c] != rec[c]}
        plan["update"].append(dict(rec=rec, changes=changes))
    return plan


def _review_exists(con, doc_no, ihash):
    return con.execute("SELECT 1 FROM repair_request_review WHERE doc_no=? AND incoming_hash=? AND "
                       "(decision IS NULL OR decision='ไม่นำเข้า')", (doc_no, ihash)).fetchone() is not None


def _insert(con, rec, batch_id):
    cols = list(FIELDS) + ["first_batch_id", "last_batch_id"]
    vals = [rec.get(c) for c in FIELDS] + [batch_id, batch_id]
    con.execute(f"INSERT INTO repair_request({','.join(cols)}) VALUES ({','.join('?' * len(cols))})", vals)


def _update(con, rec, fields, batch_id, user):
    old = dict(con.execute("SELECT * FROM repair_request WHERE doc_no=?", (rec["doc_no"],)).fetchone())
    for c in fields:
        if old[c] != rec[c]:
            con.execute("INSERT INTO repair_request_history(doc_no,batch_id,field,old_value,new_value,changed_by) VALUES (?,?,?,?,?,?)",
                        (rec["doc_no"], batch_id, c, old[c], rec[c], user))
    sets = list(fields) + ["row_hash", "confirm_flag", "flag_reason", "source_row"]
    con.execute("UPDATE repair_request SET " + ",".join(f"{c}=?" for c in sets) +
                ", last_batch_id=?, updated_at=datetime('now','localtime') WHERE doc_no=?",
                [rec[c] for c in sets] + [batch_id, rec["doc_no"]])


def backup_db(db_path):
    if not os.path.exists(db_path):
        return None
    bdir = os.path.join(os.path.dirname(os.path.abspath(db_path)), "backup")
    os.makedirs(bdir, exist_ok=True)
    dst = os.path.join(bdir, f"{os.path.splitext(os.path.basename(db_path))[0]}_{dt.datetime.now():%Y%m%d_%H%M%S}.sqlite")
    src = sqlite3.connect(db_path)
    out = sqlite3.connect(dst)
    with out:
        src.backup(out)
    src.close(); out.close()
    return dst


def commit(db_path, records, file_name, file_bytes_sha, user=None, n_blank=0):
    bpath = backup_db(db_path)
    con = connect(db_path)
    try:
        is_base = con.execute("SELECT COUNT(*) FROM repair_request").fetchone()[0] == 0
        plan = analyze(con, records)
        with con:
            cur = con.execute("INSERT INTO import_batch(file_name,file_sha256,imported_by,is_base_load,backup_path) VALUES (?,?,?,?,?)",
                              (file_name, file_bytes_sha, user, int(is_base), bpath))
            bid = cur.lastrowid
            for rec in plan["new"]:
                _insert(con, rec, bid)
            for u in plan["update"]:
                _update(con, u["rec"], SOFT, bid, user)
            for rv in plan["review"]:
                con.execute("INSERT INTO repair_request_review(batch_id,doc_no,kind,reason,incoming_json,incoming_hash) VALUES (?,?,?,?,?,?)",
                            (bid, rv["rec"]["doc_no"], rv["kind"], rv["reason"],
                             json.dumps(rv.get("variants") or rv["rec"], ensure_ascii=False),
                             rv.get("ihash") or rv["rec"]["row_hash"]))
            stats = dict(rows_read=len(records), rows_blank=n_blank, rows_dup_in_file=len(plan["dup_in_file"]),
                         rows_inserted=len(plan["new"]), rows_unchanged=len(plan["unchanged"]),
                         rows_updated=len(plan["update"]), rows_review=len(plan["review"]))
            if plan["review_known"]:
                con.execute("UPDATE import_batch SET note=? WHERE batch_id=?",
                            (f"รายการที่เคยส่งตรวจแล้ว {len(plan['review_known'])} รายการ (ไม่ส่งซ้ำ)", bid))
            con.execute("UPDATE import_batch SET " + ",".join(f"{k}=?" for k in stats) + " WHERE batch_id=?",
                        list(stats.values()) + [bid])
        return dict(batch_id=bid, backup=bpath, is_base=is_base, review_known=len(plan["review_known"]), **stats)
    finally:
        con.close()


def decide_review(db_path, review_id, decision, user, chosen_index=0):
    """decision = 'นำเข้า' | 'ไม่นำเข้า' ; ซ้ำในไฟล์ค่าต่างกัน -> chosen_index เลือกแถวที่ใช้"""
    con = connect(db_path)
    try:
        rv = con.execute("SELECT * FROM repair_request_review WHERE review_id=?", (review_id,)).fetchone()
        if rv is None or rv["decision"]:
            return False, "รายการนี้ถูกตัดสินแล้วหรือไม่พบ"
        with con:
            if decision == "นำเข้า":
                data = json.loads(rv["incoming_json"])
                rec = data[chosen_index] if isinstance(data, list) else data
                rec["confirm_flag"] = "Yes"
                rec["flag_reason"] = f"อนุมัตินำเข้าโดย {user}: {rv['reason']}"
                exists = con.execute("SELECT 1 FROM repair_request WHERE doc_no=?", (rec["doc_no"],)).fetchone()
                if exists:
                    _update(con, rec, CORE + SOFT, rv["batch_id"], user)
                else:
                    _insert(con, rec, rv["batch_id"])
            con.execute("UPDATE repair_request_review SET decision=?, decided_by=?, decided_at=datetime('now','localtime') WHERE review_id=?",
                        (decision, user, review_id))
        return True, "บันทึกแล้ว"
    finally:
        con.close()


def summary(db_path):
    if not os.path.exists(db_path):
        return None
    con = connect(db_path)
    try:
        q = lambda s, *a: con.execute(s, a).fetchall()
        return dict(
            by_type=[dict(r) for r in q("SELECT asset_type, COUNT(*) n, COUNT(DISTINCT asset_code) assets, MIN(report_date) d1, MAX(report_date) d2 FROM repair_request GROUP BY asset_type")],
            total=q("SELECT COUNT(*) FROM repair_request")[0][0],
            review_open=q("SELECT COUNT(*) FROM repair_request_review WHERE decision IS NULL")[0][0],
            batches=[dict(r) for r in q("SELECT * FROM import_batch ORDER BY batch_id DESC")],
        )
    finally:
        con.close()


def kpi_table(db_path):
    con = connect(db_path)
    try:
        return [dict(r) for r in con.execute(
            "SELECT report_month, asset_type, COUNT(*) AS vehicles, SUM(n_requests) AS requests FROM v_kpi_asset_month GROUP BY 1,2 ORDER BY 1,2")]
    finally:
        con.close()


def fetch(db_path, sql, *args):
    con = connect(db_path)
    try:
        return [dict(r) for r in con.execute(sql, args)]
    finally:
        con.close()


def file_sha256(data_bytes):
    return hashlib.sha256(data_bytes).hexdigest()


# ============================================================================= เลขไมล์ (mileage)
MILEAGE_SQL = """
CREATE TABLE IF NOT EXISTS mileage_batch (
    batch_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    file_name     TEXT NOT NULL,
    file_sha256   TEXT NOT NULL,
    imported_by   TEXT,
    imported_at   TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    rows_read     INTEGER NOT NULL DEFAULT 0,
    rows_dup_in_file INTEGER NOT NULL DEFAULT 0,
    rows_inserted INTEGER NOT NULL DEFAULT 0,
    rows_unchanged INTEGER NOT NULL DEFAULT 0,
    rows_review   INTEGER NOT NULL DEFAULT 0,
    backup_path   TEXT,
    note          TEXT
);
CREATE TABLE IF NOT EXISTS mileage_reading (
    reading_id    INTEGER PRIMARY KEY AUTOINCREMENT,
    asset_code    TEXT NOT NULL,
    reading_date  TEXT NOT NULL,
    km            INTEGER NOT NULL CHECK (km >= 0),
    data_note     TEXT,
    source_file   TEXT,
    source_sheet  TEXT,
    source_row    INTEGER,
    batch_id      INTEGER NOT NULL REFERENCES mileage_batch(batch_id),
    approved_by   TEXT,
    created_at    TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE UNIQUE INDEX IF NOT EXISTS ux_mileage_asset_date ON mileage_reading(asset_code, reading_date);
CREATE TABLE IF NOT EXISTS mileage_review (
    review_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id      INTEGER NOT NULL REFERENCES mileage_batch(batch_id),
    asset_code    TEXT NOT NULL,
    reading_date  TEXT,
    km_raw        TEXT,
    reason        TEXT NOT NULL,
    incoming_json TEXT NOT NULL,
    incoming_hash TEXT NOT NULL,
    decision      TEXT CHECK (decision IN ('นำเข้า','ไม่นำเข้า') OR decision IS NULL),
    decided_by    TEXT,
    decided_at    TEXT,
    created_at    TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE VIEW IF NOT EXISTS v_mileage_latest AS
SELECT m.asset_code, m.reading_date, m.km, m.data_note
FROM mileage_reading m
JOIN (SELECT asset_code, MAX(reading_date) d FROM mileage_reading GROUP BY asset_code) x
  ON x.asset_code = m.asset_code AND x.d = m.reading_date;
"""

_PLATE_HDR = ("registration", "ทะเบียน", "ทะเบียนรถ", "plate")
_DATE_HDR = ("date", "วันที่")
_KM_HDR = ("km.", "km", "เลขไมล์", "กม.", "mileage", "odometer")


def connect_mileage(db_path):
    con = connect(db_path)
    con.executescript(MILEAGE_SQL)
    return con


def read_mileage_workbook(src):
    """อ่านทุกชีต: หาแถวหัวตาราง (Registration/Date/km.) หรือถ้าไม่มีหัว ใช้รูปแบบแถว (ทะเบียน, วันที่, ตัวเลข)"""
    import openpyxl
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    out = dict(records=[], sheets=[], errors=[])
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        cp = cd = ck = None
        start = 0
        for i, r in enumerate(rows[:10]):
            low = [str(v).strip().lower() if v is not None else "" for v in r]
            if any(x in _PLATE_HDR for x in low) and any(x in _KM_HDR for x in low):
                cp = next(j for j, x in enumerate(low) if x in _PLATE_HDR)
                ck = next(j for j, x in enumerate(low) if x in _KM_HDR)
                cd = next((j for j, x in enumerate(low) if x in _DATE_HDR), None)
                start = i + 1
                break
        headerless = cp is None
        if headerless:          # ไม่มีหัวตาราง: ยอมรับเมื่อแถวแรกที่มีข้อมูลเป็น (ทะเบียน, วันที่, ตัวเลข)
            first = next((r for r in rows if r and r[0] not in (None, "")), None)
            if not (first and len(first) >= 3 and re.fullmatch(r"\d{2}-\d{4}", str(first[0]).strip())
                    and isinstance(first[1], (dt.datetime, dt.date))):
                continue
            cp, cd, ck = 0, 1, 2
        n = 0
        for j, r in enumerate(rows[start:], start=start + 1):
            if not r or cp >= len(r) or r[cp] in (None, ""):
                continue
            plate = str(r[cp]).strip()
            rd = _d(r[cd]) if cd is not None and cd < len(r) else None
            kmv = r[ck] if ck < len(r) else None
            flags, notes = [], []
            if not re.fullmatch(r"\d{2}-\d{4}", plate):
                flags.append(f"ทะเบียนผิดรูปแบบ: {plate}")
            if not rd:
                flags.append("ไม่มีวันที่")
            km = None
            if isinstance(kmv, bool) or kmv in (None, ""):
                flags.append("ไม่มีเลขไมล์")
            elif isinstance(kmv, (int, float)) and float(kmv).is_integer():
                km = int(kmv)
                if km < 10:
                    flags.append(f"เลขไมล์ {km} กม. (น่าจะไม่ใช่ค่าจริง)")
                elif km < 1000:
                    notes.append(f"เลขไมล์ต่ำ {km:,} กม. (อาจเปลี่ยนเครื่อง/เรือนไมล์ใหม่)")
                elif km >= 1_000_000:
                    notes.append("เกิน 1,000,000 กม.")
            else:
                flags.append(f"เลขไมล์ไม่ใช่ตัวเลข: {kmv}")
            rec = dict(asset_code=plate, reading_date=rd, km=km, km_raw=None if kmv is None else str(kmv),
                       source_sheet=ws.title, source_row=j, flags="; ".join(flags) or None, data_note="; ".join(notes) or None)
            rec["hash"] = hashlib.sha1(f"{plate}|{rd}|{rec['km_raw']}".encode()).hexdigest()
            out["records"].append(rec)
            n += 1
        out["sheets"].append(dict(sheet=ws.title, rows=n, header=not headerless))
    if not out["records"]:
        out["errors"].append("ไม่พบข้อมูลเลขไมล์ (ต้องมีคอลัมน์ ทะเบียน/Registration, วันที่/Date, km./เลขไมล์)")
    return out


def analyze_mileage(con, records):
    plan = dict(new=[], unchanged=[], review=[], review_known=[], dup_in_file=[])
    groups = {}
    for r in records:
        groups.setdefault((r["asset_code"], r["reading_date"]), []).append(r)
    known_assets = {x[0] for x in con.execute("SELECT DISTINCT asset_code FROM repair_request")}
    done = set()
    for r in records:
        key = (r["asset_code"], r["reading_date"])
        if key in done:
            if len({x["km_raw"] for x in groups[key]}) == 1:
                plan["dup_in_file"].append(r)
            continue
        done.add(key)
        variants = groups[key]
        kms = sorted({x["km_raw"] for x in variants})
        if len(kms) > 1:
            ih = hashlib.sha1("|".join(x["hash"] for x in variants).encode()).hexdigest()
            rv = dict(rec=r, reason=f"วันเดียวกันในไฟล์มีหลายค่า: {', '.join(kms)}", hash=ih, variants=variants)
        elif r["flags"]:
            rv = dict(rec=r, reason=r["flags"], hash=r["hash"])
        else:
            rv = None
            old = con.execute("SELECT km FROM mileage_reading WHERE asset_code=? AND reading_date=?", key).fetchone()
            if old is not None:
                if old[0] == r["km"]:
                    plan["unchanged"].append(r)
                    continue
                rv = dict(rec=r, reason=f"มีเลขไมล์วันนี้อยู่แล้ว {old[0]:,} แต่ไฟล์เป็น {r['km']:,}", hash=r["hash"])
            else:
                prev = con.execute("SELECT reading_date, km FROM mileage_reading WHERE asset_code=? AND reading_date<? "
                                   "ORDER BY reading_date DESC LIMIT 1", key).fetchone()
                if prev is not None and r["km"] < prev[1]:
                    rv = dict(rec=r, reason=f"ต่ำกว่าครั้งก่อน ({prev[1]:,} กม. เมื่อ {prev[0]}) — เรือนไมล์ถอยหลัง/เปลี่ยนเครื่อง?", hash=r["hash"])
        if rv:
            if con.execute("SELECT 1 FROM mileage_review WHERE asset_code=? AND incoming_hash=? AND (decision IS NULL OR decision='ไม่นำเข้า')",
                           (r["asset_code"], rv["hash"])).fetchone():
                plan["review_known"].append(r)
            else:
                plan["review"].append(rv)
            continue
        if known_assets and r["asset_code"] not in known_assets:
            r = dict(r, data_note="; ".join(x for x in (r["data_note"], "ไม่พบทะเบียนนี้ในฐานใบแจ้งซ่อม") if x))
        plan["new"].append(r)
    return plan


def _ins_mileage(con, r, bid, file_name, approved_by=None):
    con.execute("INSERT INTO mileage_reading(asset_code,reading_date,km,data_note,source_file,source_sheet,source_row,batch_id,approved_by) "
                "VALUES (?,?,?,?,?,?,?,?,?)", (r["asset_code"], r["reading_date"], r["km"], r["data_note"], file_name,
                                               r["source_sheet"], r["source_row"], bid, approved_by))


def commit_mileage(db_path, records, file_name, sha, user=None):
    bpath = backup_db(db_path)
    con = connect_mileage(db_path)
    try:
        plan = analyze_mileage(con, records)
        with con:
            bid = con.execute("INSERT INTO mileage_batch(file_name,file_sha256,imported_by,backup_path) VALUES (?,?,?,?)",
                              (file_name, sha, user, bpath)).lastrowid
            for r in plan["new"]:
                _ins_mileage(con, r, bid, file_name)
            for rv in plan["review"]:
                r = rv["rec"]
                con.execute("INSERT INTO mileage_review(batch_id,asset_code,reading_date,km_raw,reason,incoming_json,incoming_hash) VALUES (?,?,?,?,?,?,?)",
                            (bid, r["asset_code"], r["reading_date"], r["km_raw"], rv["reason"],
                             json.dumps(rv.get("variants") or r, ensure_ascii=False), rv["hash"]))
            st = dict(rows_read=len(records), rows_dup_in_file=len(plan["dup_in_file"]), rows_inserted=len(plan["new"]),
                      rows_unchanged=len(plan["unchanged"]), rows_review=len(plan["review"]))
            con.execute("UPDATE mileage_batch SET " + ",".join(f"{k}=?" for k in st) + " WHERE batch_id=?", list(st.values()) + [bid])
        return dict(batch_id=bid, backup=bpath, review_known=len(plan["review_known"]), **st)
    finally:
        con.close()


def decide_mileage_review(db_path, review_id, decision, user, km_override=None, chosen_index=0):
    """นำเข้า: ใช้ค่าตามไฟล์ (หรือ km_override ที่ผู้ตัดสินกรอกจากเอกสารจริง) ; ไม่นำเข้า: ปิดรายการ"""
    con = connect_mileage(db_path)
    try:
        rv = con.execute("SELECT * FROM mileage_review WHERE review_id=?", (review_id,)).fetchone()
        if rv is None or rv["decision"]:
            return False, "รายการนี้ถูกตัดสินแล้วหรือไม่พบ"
        with con:
            if decision == "นำเข้า":
                data = json.loads(rv["incoming_json"])
                r = data[chosen_index] if isinstance(data, list) else data
                km = km_override if km_override is not None else r.get("km")
                if km is None or not r.get("reading_date") or not re.fullmatch(r"\d{2}-\d{4}", r["asset_code"]):
                    return False, "นำเข้าไม่ได้: ต้องมีทะเบียนถูกรูปแบบ วันที่ และเลขไมล์เป็นตัวเลข (กรอกเลขไมล์ที่ถูกต้อง)"
                note = f"อนุมัติโดย {user}: {rv['reason']}" + (f" ; แก้เลขไมล์เป็น {km:,}" if km_override is not None else "")
                r = dict(r, km=int(km), data_note=note)
                con.execute("DELETE FROM mileage_reading WHERE asset_code=? AND reading_date=?", (r["asset_code"], r["reading_date"]))
                file_name = con.execute("SELECT file_name FROM mileage_batch WHERE batch_id=?", (rv["batch_id"],)).fetchone()[0]
                _ins_mileage(con, r, rv["batch_id"], file_name, approved_by=user)
            con.execute("UPDATE mileage_review SET decision=?, decided_by=?, decided_at=datetime('now','localtime') WHERE review_id=?",
                        (decision, user, review_id))
        return True, "บันทึกแล้ว"
    finally:
        con.close()


def mileage_fetch(db_path, sql, *args):
    con = connect_mileage(db_path)
    try:
        return [dict(r) for r in con.execute(sql, args)]
    finally:
        con.close()


# ============================================================================= Export รูปแบบชีต M1/M2/M3
M_SHEETS = {
    "หัวลาก": dict(name="M1_ซ่อมรถ",
                   headers=["ลำดับ", "เลขที่เอกสาร", "วันที่เปิดเอกสาร", "วันที่แจ้งซ่อม", "เวลา", "ทะเบียน / เบอร์ตู้",
                            "รหัสพนักงาน", "สถานะซ่อม", "ทะเบียน", "ยี่ห้อ", "เดือน", "Index", "ISO"],
                   formulas=['=RIGHT(F{r}, 7)', '=LEFT(F{r}, 6)', '=TEXT(D{r},"MMm-YY")', '=K{r}&":"&I{r}',
                             '=_xlfn.XLOOKUP(I{r},pm!B:B,pm!I:I,"No",0)'],
                   widths=[9.2, 18.1, 14.1, 14.1, 11.9, 29.2, 18.1, 12, 11.8, 9, 9, 12.2, 8]),
    "หางพ่วง": dict(name="M2_ซ่อมหาง",
                    headers=["ลำดับ", "เลขที่เอกสาร", "วันที่เปิดเอกสาร", "วันที่แจ้งซ่อม", "เวลา", "ทะเบียน / เบอร์ตู้",
                             "รหัสพนักงาน", "สถานะซ่อม", "ทะเบียน", "ทะเบียน", "เดือน", "Index"],
                    formulas=['=RIGHT(F{r}, 7)', '=LEFT(F{r}, 6)', '=TEXT(D{r},"MM-YY")', '=K{r}&":"&I{r}'],
                    widths=[9.2, 18.1, 14.1, 14.1, 11.9, 43.8, 18.1, 12, 11.8, 9, 8.7, 14.8]),
    "ตู้เย็น": dict(name="M3_ซ่อมตู้",
                    headers=["ลำดับ", "เลขที่เอกสาร", "วันที่เปิดเอกสาร", "วันที่แจ้งซ่อม", "เวลา", "ทะเบียน / เบอร์ตู้",
                             "รหัสพนักงาน", "สถานะซ่อม", "ทะเบียน", "Month", "Index"],
                    formulas=['=IF(F{r}="","",TRIM(RIGHT(SUBSTITUTE(F{r},"/",REPT(" ",200)),200)))', '=TEXT(D{r},"MM-YY")',
                              '=J{r}&"-"&I{r}'],
                    widths=[8.7, 20.1, 14.7, 14.7, 9.4, 44.2, 11.8, 20.9, 12.7, 13.4, 17.9]),
}


def _num_or_text(v):
    if v in (None, ""):
        return None
    try:
        f = float(v)
        return int(f) if f.is_integer() else f
    except (TypeError, ValueError):
        return v


def export_m_sheets(db_path, out_path=None):
    """สร้างไฟล์ Excel ชีต M1_ซ่อมรถ/M2_ซ่อมหาง/M3_ซ่อมตู้ จากฐานข้อมูล (1 แถว = 1 เลขเอกสาร, เรียงตามวันที่แจ้งซ่อม)
    คอลัมน์ B–H = ข้อมูล ; I–M = สูตรเดียวกับไฟล์ต้นฉบับ (ISO อ้างชีต pm ของไฟล์หลัก) ; คืน bytes ถ้าไม่ระบุ out_path"""
    import io
    import openpyxl
    from openpyxl.styles import Font, Alignment
    con = connect(db_path)
    try:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        base_font = Font(name="Angsana New", size=16)
        hdr_font = Font(name="Angsana New", size=16, bold=True)
        counts = {}
        for at, spec in M_SHEETS.items():
            ws = wb.create_sheet(spec["name"])
            for j, h in enumerate(spec["headers"], 1):
                c = ws.cell(2, j, h)
                c.font = hdr_font
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            rows = con.execute("SELECT * FROM repair_request WHERE asset_type=? ORDER BY report_date, doc_no", (at,)).fetchall()
            r = 3
            for row in rows:
                vals = [None, row["doc_no"],
                        dt.datetime.fromisoformat(row["open_date"]) if row["open_date"] else None,
                        dt.datetime.fromisoformat(row["report_date"]),
                        _num_or_text(row["report_time_raw"]) if row["report_time_raw"] is not None else 0,
                        row["asset_raw"],
                        _num_or_text(row["employee_code"]) if row["employee_code"] else ".",
                        row["status"]]
                for j, v in enumerate(vals, 1):
                    c = ws.cell(r, j, v)
                    c.font = base_font
                    if j in (3, 4):
                        c.number_format = "d/m/yyyy"
                for k, f in enumerate(spec["formulas"]):
                    c = ws.cell(r, 9 + k, f.format(r=r))
                    c.font = base_font
                r += 1
            for j, w in enumerate(spec["widths"], 1):
                ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
            ws.freeze_panes = "C3"
            ws.auto_filter.ref = f"A2:{openpyxl.utils.get_column_letter(len(spec['headers']))}{max(r - 1, 2)}"
            counts[spec["name"]] = len(rows)
        # ข้อมูลประกอบ
        info = wb.create_sheet("ข้อมูลการ Export")
        rv = con.execute("SELECT doc_no, kind, reason FROM repair_request_review WHERE decision IS NULL").fetchall()
        lines = [("สร้างจากฐานข้อมูล", os.path.basename(db_path)), ("เวลา Export", dt.datetime.now().strftime("%d/%m/%Y %H:%M"))]
        lines += [(f"จำนวนแถว {k}", v) for k, v in counts.items()]
        lines += [("หมายเหตุ", "1 แถว = 1 เลขที่เอกสาร (ตัดแถวซ้ำแล้ว) เรียงตามวันที่แจ้งซ่อม ; คอลัมน์ I–M เป็นสูตรเดียวกับไฟล์หลัก"),
                  ("หมายเหตุ ISO", "สูตร ISO อ้างชีต pm ในไฟล์หลัก — ในไฟล์นี้จะแสดง error จนกว่าจะวางลงไฟล์หลัก"),
                  ("วิธีอัปเดต", "ในไฟล์หลัก ลบข้อมูลเดิมตั้งแต่แถว 3 ของแต่ละชีต แล้วคัดลอกแถว 3 ลงไปทั้งหมด (คอลัมน์ A–M) จากไฟล์นี้ไปวางที่แถว 3"),
                  ("รายการรอตรวจที่ไม่ได้อยู่ในไฟล์นี้", len(rv))]
        for k, v in lines:
            info.append([k, v])
        for x in rv:
            info.append([x["doc_no"], f"{x['kind']}: {x['reason']}"])
        info.column_dimensions["A"].width = 36
        info.column_dimensions["B"].width = 110
        if out_path:
            wb.save(out_path)
            return out_path, counts
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue(), counts
    finally:
        con.close()
