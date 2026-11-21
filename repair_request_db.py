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
