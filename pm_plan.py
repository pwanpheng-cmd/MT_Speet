# -*- coding: utf-8 -*-
"""PM Plan engine — ประวัติบริการ (จากไฟล์ Export รายงานซ่อมรายคัน) + คำนวณกำหนด PM รายคัน

แหล่งข้อมูล (ไฟล์ฐานข้อมูลเดียวกับใบแจ้งซ่อม/เลขไมล์):
  service_doc      เอกสาร W/O (ใบสั่งซ่อมภายนอก) และ WX (ใบเบิกพัสดุ) รายคัน พร้อมวันที่ เลขไมล์ และธง PM
  mileage_reading  เลขไมล์จากไฟล์เลขไมล์ (ตารางเดิม)
  repair_request   วันที่แจ้งซ่อมล่าสุด (ข้อมูลประกอบ)

เกณฑ์ (ยืนยันโดย Owner 22/11/2026):
  PM ศูนย์บริการ = W/O ที่มีรายการน้ำมันเครื่อง/ไส้กรองน้ำมันเครื่อง  → รอบถัดไปทุก 20,000 กม. หรือ 6 เดือน
  PM ลาน          = ใบเบิก WX ที่มีน้ำมันเครื่อง                       → รอบถัดไปทุก 30,000 กม. หรือ 6 เดือน
  ครบกำหนด = ถึงอย่างใดอย่างหนึ่งก่อน
  รอบถัดไปใช้ระยะตามประเภท PM ครั้งล่าสุดของคันนั้น (ศูนย์ → 20,000 · ลาน → 30,000)
"""
import os, re, json, sqlite3, datetime as dt

import repair_request_db as rr

PM_SQL = """
CREATE TABLE IF NOT EXISTS service_batch (
    batch_id INTEGER PRIMARY KEY AUTOINCREMENT,
    file_name TEXT, file_sha256 TEXT, imported_by TEXT,
    imported_at TEXT DEFAULT (datetime('now','localtime')),
    rows_read INTEGER, rows_inserted INTEGER, rows_updated INTEGER, rows_unchanged INTEGER,
    backup_path TEXT
);
CREATE TABLE IF NOT EXISTS service_doc (
    doc_no TEXT PRIMARY KEY,               -- เลขที่ใบสั่งซ่อม (W/O) หรือ เลขที่ใบเบิก (WX)
    doc_kind TEXT NOT NULL CHECK (doc_kind IN ('W/O','WX')),
    asset_code TEXT NOT NULL,              -- ทะเบียน
    doc_date TEXT,                         -- YYYY-MM-DD
    km INTEGER,                            -- เลขไมล์ตามเอกสาร (ค่าดิบ)
    km_status TEXT,                        -- ชัดเจน / ตีความ / ต้องตรวจ (จากตัวอ่านรายงาน)
    vendor TEXT,
    pm_type TEXT,                          -- 'ศูนย์บริการ' / 'ลาน' / NULL
    pm_items TEXT,                         -- รายการที่ทำให้เป็น PM
    amount REAL,
    row_hash TEXT,
    source_file TEXT,
    batch_id INTEGER,
    updated_at TEXT DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS ix_service_doc_asset ON service_doc(asset_code, doc_date);
"""

ODO_WRAP = 1_000_000          # เลขไมล์บางแหล่งเป็น 6 หลัก (วนรอบที่ 1,000,000) — คำนวณระยะแบบ mod
DEFAULTS = dict(center_km=20000, yard_km=30000, months=6, window_days=365, max_km_day=1000, min_km_day=20)

_OIL_EXCL = ("ปะเก็น", "ประเก็น", "ท่อ", "ฝา", "ถัง", "อ่าง", "ซีล")


def is_center_item(name):
    """W/O: น้ำมันเครื่อง หรือ ไส้กรอง/กรองน้ำมันเครื่อง (ไม่นับ ปะเก็น/ท่อ/ฝา/ถัง)"""
    n = str(name or "")
    return "น้ำมันเครื่อง" in n and not any(x in n for x in _OIL_EXCL)


def is_yard_oil(name):
    """WX: ตัวน้ำมันเครื่อง (ไม่นับไส้กรองและอะไหล่ประกอบ)"""
    n = str(name or "")
    return "น้ำมันเครื่อง" in n and "กรอง" not in n and not any(x in n for x in _OIL_EXCL)


def connect(db_path):
    con = rr.connect_mileage(db_path)       # สร้างตารางหลัก + เลขไมล์ ถ้ายังไม่มี
    con.executescript(PM_SQL)
    return con


def _date(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y"):
        try:
            return dt.datetime.strptime(s[:19] if "-" in s else s, fmt).date().isoformat()
        except ValueError:
            pass
    return None


def _int(v):
    try:
        if v is None or v == "":
            return None
        f = float(v)
        return None if f != f else int(round(f))
    except (TypeError, ValueError):
        return None


def read_report_export(src):
    """อ่านไฟล์ Export จากหน้า 'รายงานซ่อมรายคัน' (Vehicle_Reports_*.xlsx): ชีต ใบสั่งซ่อม_หัว / ใบสั่งซ่อม_บรรทัด / เบิกพัสดุ"""
    import openpyxl
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    need = ["ใบสั่งซ่อม_หัว", "ใบสั่งซ่อม_บรรทัด", "เบิกพัสดุ"]
    miss = [s for s in need if s not in wb.sheetnames]
    if miss:
        return dict(docs=[], errors=[f"ไม่พบชีต {', '.join(miss)} — ใช้ไฟล์ที่ Export จากหน้า 'รายงานซ่อมรายคัน'"], stats={})

    def rows(name):
        it = wb[name].iter_rows(values_only=True)
        hdr = [str(h).strip() if h is not None else "" for h in next(it)]
        for r in it:
            if r and any(v not in (None, "") for v in r):
                yield dict(zip(hdr, r))

    docs, errors = {}, []
    wo_items = {}
    for r in rows("ใบสั่งซ่อม_บรรทัด"):
        if is_center_item(r.get("ชื่อรายการ")):
            wo_items.setdefault(str(r.get("เลขที่ใบสั่งซ่อม")).strip(), []).append(str(r.get("ชื่อรายการ")).strip())
    for r in rows("ใบสั่งซ่อม_หัว"):
        no = str(r.get("เลขที่ใบสั่งซ่อม") or "").strip()
        plate = str(r.get("ทะเบียน") or "").strip()
        if not no or not plate:
            continue
        items = wo_items.get(no, [])
        docs[no] = dict(doc_no=no, doc_kind="W/O", asset_code=plate, doc_date=_date(r.get("วันที่")),
                        km=_int(r.get("เลขไมล์")), km_status=(r.get("สถานะเลขไมล์") or None),
                        vendor=(r.get("ผู้ขาย (TH)") or r.get("ผู้ขาย (EN)") or None),
                        pm_type="ศูนย์บริการ" if items else None, pm_items=" · ".join(sorted(set(items))) or None,
                        amount=r.get("ยอดสุทธิ"))
    for r in rows("เบิกพัสดุ"):
        no = str(r.get("เลขที่ใบเบิก") or "").strip()
        plate = str(r.get("ทะเบียน") or "").strip()
        if not no or not plate:
            continue
        d = docs.get(no)
        if d is None:
            d = docs[no] = dict(doc_no=no, doc_kind="WX", asset_code=plate, doc_date=_date(r.get("วันที่")), km=None,
                                km_status=None, vendor=(r.get("คลัง") or None), pm_type=None, pm_items=None, amount=0.0)
        km, stt = _int(r.get("เลขไมล์")), r.get("สถานะเลขไมล์")
        if km is not None and (d["km"] is None or (d["km_status"] != "ชัดเจน" and stt == "ชัดเจน")):
            d["km"], d["km_status"] = km, stt
        d["amount"] = (d["amount"] or 0) + (r.get("ยอดสุทธิ") or 0)
        if is_yard_oil(r.get("ชื่อพัสดุ")):
            d["pm_type"] = "ลาน"
            its = set(filter(None, (d["pm_items"] or "").split(" · "))) | {str(r.get("ชื่อพัสดุ")).strip()}
            d["pm_items"] = " · ".join(sorted(its))
    out = []
    for d in docs.values():
        d["amount"] = round(float(d["amount"] or 0), 2)
        d["row_hash"] = rr.file_sha256(json.dumps({k: d[k] for k in sorted(d)}, ensure_ascii=False, default=str).encode())[:20]
        out.append(d)
    st = dict(docs=len(out), wo=sum(d["doc_kind"] == "W/O" for d in out), wx=sum(d["doc_kind"] == "WX" for d in out),
              pm_center=sum(d["pm_type"] == "ศูนย์บริการ" for d in out), pm_yard=sum(d["pm_type"] == "ลาน" for d in out),
              plates=len({d["asset_code"] for d in out}))
    return dict(docs=out, errors=errors, stats=st)


def analyze_service(db_path, docs):
    con = connect(db_path)
    try:
        cur = {r["doc_no"]: r["row_hash"] for r in con.execute("SELECT doc_no, row_hash FROM service_doc")}
    finally:
        con.close()
    new = [d for d in docs if d["doc_no"] not in cur]
    upd = [d for d in docs if d["doc_no"] in cur and cur[d["doc_no"]] != d["row_hash"]]
    same = [d for d in docs if cur.get(d["doc_no"]) == d["row_hash"]]
    return dict(new=new, update=upd, unchanged=same)


COLS = ["doc_no", "doc_kind", "asset_code", "doc_date", "km", "km_status", "vendor", "pm_type", "pm_items", "amount", "row_hash"]


def commit_service(db_path, docs, file_name, sha, user=None):
    bpath = rr.backup_db(db_path)
    plan = analyze_service(db_path, docs)
    con = connect(db_path)
    try:
        with con:
            bid = con.execute("INSERT INTO service_batch(file_name,file_sha256,imported_by,backup_path) VALUES (?,?,?,?)",
                              (file_name, sha, user, bpath)).lastrowid
            for d in plan["new"] + plan["update"]:
                con.execute(f"INSERT OR REPLACE INTO service_doc({','.join(COLS)},source_file,batch_id,updated_at) "
                            f"VALUES ({','.join('?' * len(COLS))},?,?,datetime('now','localtime'))",
                            [d[c] for c in COLS] + [file_name, bid])
            st = dict(rows_read=len(docs), rows_inserted=len(plan["new"]), rows_updated=len(plan["update"]),
                      rows_unchanged=len(plan["unchanged"]))
            con.execute("UPDATE service_batch SET " + ",".join(f"{k}=?" for k in st) + " WHERE batch_id=?", list(st.values()) + [bid])
        return dict(batch_id=bid, backup=bpath, **st)
    finally:
        con.close()


def fetch(db_path, sql, *args):
    con = connect(db_path)
    try:
        return [dict(r) for r in con.execute(sql, args)]
    finally:
        con.close()


# ----------------------------------------------------------------------------- คำนวณ
def _d(s):
    return dt.date.fromisoformat(s) if isinstance(s, str) else s


def _inc(a, b):
    """ระยะจาก a ไป b (กม.) — ถ้า b < a นับว่าวนรอบ (เลขไมล์ 6 หลัก) เฉพาะเมื่อ a ≥ 900,000 และ b < 100,000 ; ไม่เช่นนั้น = ผิดปกติ (None)"""
    if b >= a:
        return b - a
    if a >= ODO_WRAP - 100_000 and b < 100_000:
        return b + ODO_WRAP - a
    return None


def _consistent(p, q, max_km_day):
    (d1, k1), (d2, k2) = (p, q) if p[0] <= q[0] else (q, p)
    days = (d2 - d1).days
    if days == 0:
        return abs(k2 - k1) <= 500
    inc = _inc(k1, k2)
    return inc is not None and inc <= max_km_day * days


def clean_series(points, max_km_day):
    """ตัดจุดเลขไมล์ที่ขัดกับจุดอื่น (เช่น พิมพ์ผิด/ค่า 1) — ตัดจุดที่ขัดแย้งมากที่สุดทีละจุดจนสอดคล้องกันทั้งหมด"""
    pts = sorted(set(points))
    removed = []
    while len(pts) > 1:
        bad = [sum(not _consistent(p, q, max_km_day) for q in pts if q is not p) for p in pts]
        worst = max(bad)
        if worst == 0:
            break
        i = bad.index(worst)
        removed.append(pts.pop(i))
    return pts, removed


def daily_rate(pts, window_days, today):
    """กม./วัน = ระยะรวม ÷ จำนวนวัน ของช่วงเลขไมล์ต่อเนื่องในหน้าต่างล่าสุด
    ตัดช่วงที่เร็วผิดปกติ (> 3 เท่าของค่ากลางช่วงอื่น + 100) คืน (rate, จำนวนช่วงที่ตัด)"""
    if len(pts) < 2:
        return None, 0
    last = pts[-1][0]
    win = [p for p in pts if (last - p[0]).days <= window_days]
    if len(win) < 2:
        win = pts[-2:]
    segs = []
    for a, b in zip(win, win[1:]):
        days = (b[0] - a[0]).days
        inc = _inc(a[1], b[1])
        if days > 0 and inc is not None:
            segs.append((days, inc))
    if not segs:
        return None, 0
    rs = sorted(inc / days for days, inc in segs if days >= 3) or sorted(inc / days for days, inc in segs)
    med = rs[len(rs) // 2]
    keep = [(d, k) for d, k in segs if k / d <= 3 * med + 100]
    days = sum(d for d, _ in keep)
    if days < 7:
        return None, len(segs) - len(keep)
    return sum(k for _, k in keep) / days, len(segs) - len(keep)


def add_months(d, n):
    y, m = divmod(d.month - 1 + n, 12)
    y, m = d.year + y, m + 1
    import calendar
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def compute_plan(db_path, today=None, **opt):
    o = {**DEFAULTS, **{k: v for k, v in opt.items() if v is not None}}
    today = today or dt.date.today()
    con = connect(db_path)
    try:
        plates = {r[0] for r in con.execute("SELECT DISTINCT asset_code FROM repair_request WHERE asset_type='หัวลาก' AND in_iso='Yes'")}
        plates |= {r[0] for r in con.execute("SELECT DISTINCT asset_code FROM service_doc")}
        last_rr = {r[0]: r[1] for r in con.execute("SELECT asset_code, MAX(report_date) FROM repair_request WHERE asset_type='หัวลาก' AND in_iso='Yes' GROUP BY 1")}
        pts = {}
        for a, d, k in con.execute("SELECT asset_code, reading_date, km FROM mileage_reading WHERE km IS NOT NULL"):
            pts.setdefault(a, []).append((_d(d), int(k), "เลขไมล์"))
        for a, d, k, s, no in con.execute("SELECT asset_code, doc_date, km, km_status, doc_no FROM service_doc "
                                          "WHERE km IS NOT NULL AND doc_date IS NOT NULL AND IFNULL(km_status,'') <> 'ต้องตรวจ'"):
            if k >= 1000:
                pts.setdefault(a, []).append((_d(d), int(k), no))
        events = {}
        for r in con.execute("SELECT * FROM service_doc WHERE pm_type IS NOT NULL AND doc_date IS NOT NULL ORDER BY doc_date"):
            events.setdefault(r["asset_code"], []).append(dict(r))
    finally:
        con.close()

    rows, rates = [], {}
    series = {}
    for a in plates:
        raw = pts.get(a, [])
        clean, removed = clean_series([(d, k) for d, k, _ in raw], o["max_km_day"])
        r, nskip = daily_rate(clean, o["window_days"], today)
        series[a] = (clean, removed, len(raw), nskip, r)
        if r is not None and o["min_km_day"] <= r <= o["max_km_day"]:
            rates[a] = r
    vals = sorted(rates.values())
    fleet = vals[len(vals) // 2] if vals else None

    for a in sorted(plates):
        clean, removed, nraw, nskip, r0 = series[a]
        rate = rates.get(a)
        notes = []
        rate_src = "ประวัติเลขไมล์"
        if rate is None:
            rate, rate_src = fleet, "ค่ากลางทั้งกอง"
            if r0 is not None:
                notes.append(f"กม./วันจากประวัติ ({r0:,.0f}) อยู่นอกช่วง {o['min_km_day']}–{o['max_km_day']:,} ใช้ค่ากลางของกองรถ")
            else:
                notes.append("ข้อมูลเลขไมล์ไม่พอ ใช้ กม./วัน ค่ากลางของกองรถ")
        if removed:
            notes.append(f"ตัดเลขไมล์ผิดปกติ {len(removed)} ค่า")
        if nskip:
            notes.append(f"ตัดช่วงวิ่งเร็วผิดปกติ {nskip} ช่วง")
        ev = events.get(a, [])
        last_c = next((e for e in reversed(ev) if e["pm_type"] == "ศูนย์บริการ"), None)
        last_y = next((e for e in reversed(ev) if e["pm_type"] == "ลาน"), None)
        last = ev[-1] if ev else None
        cur_date, cur_km = (clean[-1] if clean else (None, None))
        row = dict(asset_code=a, last_repair=last_rr.get(a),
                   last_center=last_c["doc_date"] if last_c else None, last_center_km=last_c["km"] if last_c else None,
                   last_yard=last_y["doc_date"] if last_y else None, last_yard_km=last_y["km"] if last_y else None,
                   last_km_date=cur_date.isoformat() if cur_date else None, last_km=cur_km,
                   km_points=len(clean), km_day=round(rate, 1) if rate else None, rate_src=rate_src)
        if last is None:
            row.update(pm_type="ยังไม่มีประวัติ PM", interval_km=None, last_pm=None, last_pm_doc=None, run_km=None,
                       remain_km=None, due_time=None, due_km_date=None, due_date=None, due_by="ไม่มีประวัติ",
                       status="ไม่มีประวัติ PM", notes="; ".join(notes + ["ไม่พบ PM ในข้อมูลที่นำเข้า — ควรนัดตรวจเช็ค"]))
            rows.append(row)
            continue
        ptype = last["pm_type"]
        interval = o["center_km"] if ptype == "ศูนย์บริการ" else o["yard_km"]
        ld = _d(last["doc_date"])
        # ระยะที่วิ่งไปแล้วตั้งแต่ PM ล่าสุด
        run, how = None, ""
        ek = last["km"] if (last["km"] and last["km"] >= 1000 and last["km_status"] != "ต้องตรวจ") else None
        if ek is not None and clean and all(_consistent((ld, ek), p, o["max_km_day"]) for p in clean[-3:]):
            run = (_inc(ek, cur_km) or 0) + (rate or 0) * max((today - cur_date).days, 0)
            how = "เลขไมล์ PM + เลขไมล์ล่าสุด"
        elif rate:
            run = rate * max((today - ld).days, 0)
            how = "กม./วัน × จำนวนวันตั้งแต่ PM"
            if ek is None:
                notes.append("PM ล่าสุดไม่มีเลขไมล์ที่ใช้ได้")
            else:
                notes.append("เลขไมล์ตอน PM ขัดกับประวัติ")
        due_time = add_months(ld, o["months"])
        due_km_date, remain = None, None
        if run is not None:
            remain = interval - run
            if rate:
                due_km_date = today + dt.timedelta(days=int(remain // rate))
        if due_km_date and due_km_date < due_time:
            due, by = due_km_date, f"ระยะ {interval:,} กม."
        else:
            due, by = due_time, f"ครบ {o['months']} เดือน"
        status = "เกินกำหนด" if due < today else ("ครบกำหนดเดือนนี้" if (due.year, due.month) == (today.year, today.month) else "ยังไม่ถึง")
        row.update(pm_type=ptype, interval_km=interval, last_pm=last["doc_date"], last_pm_doc=last["doc_no"],
                   run_km=int(run) if run is not None else None, remain_km=int(remain) if remain is not None else None,
                   run_method=how, due_time=due_time.isoformat(), due_km_date=due_km_date.isoformat() if due_km_date else None,
                   due_date=due.isoformat(), due_by=by, status=status, notes="; ".join(notes))
        rows.append(row)
    return dict(rows=rows, fleet_rate=fleet, options=o, today=today.isoformat())


def km_history(db_path, plate):
    """ประวัติเลขไมล์ของคันหนึ่ง (รวมทุกแหล่ง) พร้อมธงว่าถูกตัดออกจากการคำนวณหรือไม่"""
    o = DEFAULTS
    con = connect(db_path)
    try:
        raw = [(r[0], int(r[1]), "ไฟล์เลขไมล์", "") for r in
               con.execute("SELECT reading_date, km FROM mileage_reading WHERE asset_code=? AND km IS NOT NULL", (plate,))]
        raw += [(r[0], int(r[1]), r[2], r[3]) for r in
                con.execute("SELECT doc_date, km, doc_kind, doc_no FROM service_doc WHERE asset_code=? AND km IS NOT NULL "
                            "AND doc_date IS NOT NULL AND IFNULL(km_status,'')<>'ต้องตรวจ' AND km>=1000", (plate,))]
    finally:
        con.close()
    clean, removed = clean_series([(_d(d), k) for d, k, _, _ in raw], o["max_km_day"])
    rem = set(removed)
    return sorted([dict(date=d, km=k, source=s, doc_no=n, used="ตัดออก (ขัดกับค่าอื่น)" if (_d(d), k) in rem else "ใช้คำนวณ")
                   for d, k, s, n in raw], key=lambda x: x["date"])
