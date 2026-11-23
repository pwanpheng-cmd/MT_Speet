# -*- coding: utf-8 -*-
"""
Vehicle Report Importer v1.0.0
อ่านรายงาน Crystal "Finish Goods Takeoff and Work Order Report by Vehicle" (rpt_SumWO_Date.rpt)
รายคัน (.xls / .xlsx) ทั้งโฟลเดอร์ -> แสดงใน GUI -> Export Excel

- อ่านไฟล์ต้นทางแบบอ่านอย่างเดียว (ไม่แก้ไฟล์ต้นทาง ไม่แตะฐานข้อมูล Fleet_DB)
- กระทบยอดต่อไฟล์ 6 จุด, ตรวจไฟล์ซ้ำ (เนื้อหาเหมือนกัน), ติด Confirm Flag Yes/No พร้อมเหตุผล
- วันที่เก็บเป็น datetime แสดง DD/MM/YYYY
"""
import os, sys, re, hashlib, json, threading, queue, traceback, warnings
from datetime import datetime, timedelta, date

warnings.filterwarnings("ignore")
APP_VERSION = "1.1.0"
DEFAULT_SRC = r"D:\Vehicle_Reports"
DEFAULT_OUT = r"D:\MT_Speed\Export_File"

# ----------------------------------------------------------------------------- parser
def _load_rows(path):
    if path.lower().endswith(".xls"):
        import xlrd
        sh = xlrd.open_workbook(path).sheet_by_index(0)
        return [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = [[("" if v is None else v) for v in row] for row in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def _g(row, c):
    return row[c] if c < len(row) else ""


def _num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and 30000 < v < 80000:
        return (datetime(1899, 12, 30) + timedelta(days=float(v))).date()
    return None


def _clean_rem(s):
    return re.sub(r"\x7f|Rem", "", str(s)).strip()


def _blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


_KM = r"(\d{1,3}(?:,\d{3})+|\d{1,7})"


def extract_mileage(text, from_column=False):
    """ดึงเลขไมล์จากข้อความ คืน (ค่า, ที่มา, สถานะ) — สถานะ: ชัดเจน / ตีความ / ต้องตรวจ ; ไม่พบ = (None, "", "")"""
    s = str(text or "").strip()
    if not s:
        return None, "", ""
    toint = lambda x: int(re.sub(r"[,\s]", "", x))
    if from_column:
        m = re.fullmatch(r"([\d,]+)\s*(?:km\.?|กม\.?)?", s, re.I)
        if m and re.sub(r"\D", "", m.group(1)):
            v = toint(m.group(1))
            return v, "คอลัมน์เลขไมล์", ("ชัดเจน" if v >= 100 else "ต้องตรวจ")
        m = re.fullmatch(r"(\d{1,4})\.(\d{3})\s*(?:km\.?|กม\.?)?", s, re.I)
        if m:
            return int(m.group(1) + m.group(2)), "คอลัมน์เลขไมล์ (จุดคั่นหลักพัน)", "ตีความ"
        return None, f"คอลัมน์เลขไมล์อ่านไม่ได้: {s[:30]}", "ต้องตรวจ"
    pats = [
        (r"(?<![\d.\-])" + _KM + r"\s*(?:กม|km)\.?", "หมายเหตุ: …กม./KM"),
        (r"เลขไมล์\s*:?\s*" + _KM + r"(?![\d.])", "หมายเหตุ: เลขไมล์ …"),
        (r"^/?\s*" + _KM + r"\s*[@&]\s*MT[DS]", "หมายเหตุ: …@MTD-SR"),
        (r"MTD-SR\d+\s*/*\s*(\d{4,7})\s*$", "หมายเหตุ: MTD-SR…//…"),
        (r"//\s*(\d{1,7})\s*(?=///|$|\s|\()", "หมายเหตุ: //…"),
    ]
    for p, src in pats:
        m = re.search(p, s, re.I)
        if m:
            v = toint(m.group(1))
            return v, src, ("ชัดเจน" if v >= 100 else "ต้องตรวจ")
    m = re.search(r"//\s*(\d{1,4})\.(\d{3})(?:\.(\d{3}))?\s*(?=///|$|\s|\()", s)
    if m:
        return int("".join(g for g in m.groups() if g)), "หมายเหตุ: //…(จุดคั่นหลักพัน)", "ตีความ"
    m = re.fullmatch(r"/?\s*(\d{4,7})", s)
    if m:
        return int(m.group(1)), "หมายเหตุ: ตัวเลขล้วน", "ต้องตรวจ"
    return None, "", ""


def issue_line_mileage(raw_col, remark):
    a = extract_mileage(raw_col, from_column=True) if not _blank(raw_col) else (None, "", "")
    b = extract_mileage(remark)
    if a[0] is not None and b[0] is not None and a[0] != b[0] and a[2] == "ชัดเจน" and b[2] == "ชัดเจน":
        return a[0], f"คอลัมน์ {a[0]:,} ≠ หมายเหตุ {b[0]:,}", "ต้องตรวจ"
    if a[0] is not None:
        return a
    if b[0] is not None:
        return b[0], b[1], b[2] if not a[1] else "ต้องตรวจ"
    return None, a[1], a[2]


def vehicle_info(desc):
    """แยกยี่ห้อ/ประเภทจากคำอธิบายรถ; คืน (brand, type, standard_format)"""
    d = (desc or "").strip()
    m = re.search(r"\(([^)]*)\)\s*$", d)
    typ_raw = m.group(1).strip() if m else ""
    typ = {"prime mover": "หัวลาก"}.get(typ_raw.lower(), typ_raw)
    brand = d.split("/")[0].strip().upper() if "/" in d else ""
    std = bool(m) and bool(brand)
    if not std:  # รูปแบบไม่มาตรฐาน -> หาเฉพาะคำที่มีจริงในข้อความ (ไม่เดา)
        up = d.upper()
        brand = next((b for b in ("HINO", "SCANIA", "ISUZU", "VOLVO", "MITSUBISHI", "FUSO", "UD") if b in up), "")
        typ = "หัวลาก" if "หัวลาก" in d else typ
    return brand, typ, std


def parse_report(path):
    rows = _load_rows(path)
    o = dict(file=os.path.basename(path), path=path, plate="", vehicle_desc="", title="", print_ts="",
             rpt="", issues_hdr=[], issues=[], wo_hdr=[], wo_lines=[], totals={}, anomalies=[])
    sec = "A"
    cur_doc = None
    cur_wo = None
    for i, row in enumerate(rows):
        cells = [(c, v) for c, v in enumerate(row) if not _blank(v)]
        if not cells:
            continue
        txt = " | ".join(str(v) for _, v in cells)
        if "Finish Goods Takeoff" in txt:
            o["title"] = str(cells[0][1]).strip(); continue
        if "License Plate" in txt:
            o["plate"] = str(_g(row, 8)).strip(); o["vehicle_desc"] = str(_g(row, 10)).strip(); continue
        if "Doc Code" in txt:
            continue
        if "Work Order List" in txt:
            sec = "B"; continue
        if "WX. Total" in txt:
            o["totals"]["wx"] = _num(_g(row, 27)); continue
        if "W/O Total" in txt:
            o["totals"]["wo"] = _num(_g(row, 24)); continue
        if "Total Amount" in txt:
            o["totals"]["all"] = _num(_g(row, 27)); continue
        if "Print" in txt and re.search(r"\d\d/\d\d/\d{4}", txt):
            o["print_ts"] = re.sub(r"\s+", " ", txt.replace("Print", "").replace(":", "", 1)).strip(); continue
        if txt.lower().endswith(".rpt"):
            o["rpt"] = txt.strip(); continue
        if sec == "A":
            if not _blank(_g(row, 7)):
                o["issues"].append(dict(src_row=i + 1, doc_no=cur_doc, date=_date(_g(row, 5)),
                                        inv_code=str(_g(row, 7)).strip(), inv_name=str(_g(row, 10)).strip(),
                                        qty=_num(_g(row, 16)), unit_price=_num(_g(row, 21)), total=_num(_g(row, 27)),
                                        dept=(str(_g(row, 30)).strip() or str(_g(row, 29)).strip()),
                                        store=str(_g(row, 34)).strip(), mileage=_g(row, 36)))
            elif not _blank(_g(row, 2)):
                cur_doc = str(_g(row, 2)).strip()
                o["issues_hdr"].append(dict(src_row=i + 1, doc_no=cur_doc, remark=_clean_rem(_g(row, 5)),
                                            doc_total=_num(_g(row, 27))))
            else:
                o["anomalies"].append((i + 1, "ส่วนเบิกพัสดุ: แถวไม่รู้จัก", txt[:150]))
        else:
            if not _blank(_g(row, 7)):
                o["wo_lines"].append(dict(src_row=i + 1, hdr_idx=len(o["wo_hdr"]) - 1, doc_no=str(_g(row, 2)).strip(),
                                          date=_date(_g(row, 5)), item_code=str(_g(row, 7)).strip(),
                                          item_name=str(_g(row, 11)).strip() or str(_g(row, 12)).strip(),
                                          qty=_num(_g(row, 16)), unit_price=_num(_g(row, 21)), total=_num(_g(row, 27))))
            elif _date(_g(row, 5)):
                cur_wo = dict(src_row=i + 1, date=_date(_g(row, 5)), remark=str(_g(row, 8)).strip(),
                              amount=_num(_g(row, 14)), discount=_num(_g(row, 18)), net=_num(_g(row, 26)),
                              vendor_en=str(_g(row, 32)).strip(), vendor_th="")
                o["wo_hdr"].append(cur_wo)
            elif cur_wo is not None and cur_wo["amount"] is None and _num(_g(row, 14)) is not None:
                cur_wo.update(amount=_num(_g(row, 14)), discount=_num(_g(row, 18)), net=_num(_g(row, 26)))
            elif cur_wo is not None and len(cells) == 1 and cells[0][0] == 32:
                if not cur_wo["vendor_en"]:
                    cur_wo["vendor_en"] = str(cells[0][1]).strip()
                elif not cur_wo["vendor_th"]:
                    cur_wo["vendor_th"] = str(cells[0][1]).strip()
            else:
                o["anomalies"].append((i + 1, "ส่วนใบสั่งซ่อม: แถวไม่รู้จัก", txt[:150]))

    # link WO header -> doc no (from its lines), line numbers
    for k, h in enumerate(o["wo_hdr"]):
        ls = [l for l in o["wo_lines"] if l["hdr_idx"] == k]
        docs = sorted({l["doc_no"] for l in ls})
        h["doc_no"] = docs[0] if len(docs) == 1 else ",".join(docs)
        h["line_sum"] = round(sum(l["total"] or 0 for l in ls), 2)
        h["n_lines"] = len(ls)
        h["has_vat_line"] = any("ภาษีมูลค่าเพิ่ม" in l["item_name"] for l in ls)
        for n, l in enumerate(ls, 1):
            l["line_no"] = n
            l["vendor_en"] = h["vendor_en"]
            l["is_vat"] = "ภาษีมูลค่าเพิ่ม" in l["item_name"]
    seq = {}
    for l in o["issues"]:
        seq[l["doc_no"]] = seq.get(l["doc_no"], 0) + 1
        l["line_no"] = seq[l["doc_no"]]
    hdr_by_doc = {h["doc_no"]: h for h in o["issues_hdr"]}
    for l in o["issues"]:
        l["remark"] = hdr_by_doc.get(l["doc_no"], {}).get("remark", "")
        l["mileage_raw"] = "" if _blank(l["mileage"]) else str(l["mileage"]).strip()
        l["mileage_km"], l["mileage_src"], l["mileage_status"] = issue_line_mileage(l["mileage"], l["remark"])
    for h in o["wo_hdr"]:
        h["mileage_km"], h["mileage_src"], h["mileage_status"] = extract_mileage(h["remark"])
    for k, h in enumerate(o["wo_hdr"]):
        for l in o["wo_lines"]:
            if l["hdr_idx"] == k:
                l["mileage_km"], l["mileage_src"], l["mileage_status"] = h["mileage_km"], h["mileage_src"], h["mileage_status"]

    r2 = lambda x: round(x or 0, 2)
    t = o["totals"]
    chk = {
        "บรรทัดเบิก = WX Total": r2(sum(x["total"] or 0 for x in o["issues"])) == r2(t.get("wx")),
        "หัวเบิก = WX Total": r2(sum(x["doc_total"] or 0 for x in o["issues_hdr"])) == r2(t.get("wx")),
        "บรรทัด WO = W/O Total": r2(sum(x["total"] or 0 for x in o["wo_lines"])) == r2(t.get("wo")),
        "หัว WO = W/O Total": r2(sum(x["net"] or 0 for x in o["wo_hdr"])) == r2(t.get("wo")),
        "WX + W/O = Total Amount": r2((t.get("wx") or 0) + (t.get("wo") or 0)) == r2(t.get("all")),
        "ทะเบียนตรงชื่อไฟล์": bool(o["plate"]) and o["file"].startswith(o["plate"]),
    }
    o["checks"] = chk
    o["brand"], o["vtype"], o["desc_std"] = vehicle_info(o["vehicle_desc"])
    sig = json.dumps([o["plate"], o["totals"],
                      [(l["doc_no"], l["inv_code"], l["qty"], l["total"]) for l in o["issues"]],
                      [(l["doc_no"], l["item_code"], l["qty"], l["total"]) for l in o["wo_lines"]]],
                     default=str, ensure_ascii=False)
    o["content_hash"] = hashlib.sha1(sig.encode("utf-8")).hexdigest()[:12]
    return o


def scan_folder(folder, progress=None, dedupe=True):
    files = sorted(f for f in os.listdir(folder)
                   if f.lower().endswith((".xls", ".xlsx")) and not f.startswith("~$"))
    reports, skipped = [], []
    for n, f in enumerate(files, 1):
        p = os.path.join(folder, f)
        if progress:
            progress(n, len(files), f)
        try:
            o = parse_report(p)
            if not o["plate"]:
                skipped.append((f, "ไม่พบหัวรายงาน 'ทะเบียนรถ/License Plate' (ไม่ใช่รูปแบบรายงานรายคัน หรือไฟล์ว่าง)"))
                continue
            reports.append(o)
        except Exception as e:
            skipped.append((f, f"อ่านไม่ได้: {type(e).__name__}: {e}"))
    # duplicates
    first = {}
    for o in reports:
        h = o["content_hash"]
        if h in first:
            o["duplicate_of"] = first[h]["file"]
        else:
            o["duplicate_of"] = ""
            first[h] = o
    plate_files = {}
    for o in reports:
        plate_files.setdefault(o["plate"], []).append(o)
    for o in reports:
        others = [x["file"] for x in plate_files[o["plate"]] if x is not o and x["content_hash"] != o["content_hash"]]
        o["same_plate_diff_content"] = others
    return build_dataset(folder, reports, skipped, dedupe)


def build_dataset(folder, reports, skipped, dedupe):
    used = [o for o in reports if not (dedupe and o["duplicate_of"])]
    files_rows, issues, wo_hdr, wo_lines, flags, per_vehicle = [], [], [], [], [], []

    def flag(level, file, plate, kind, ref, detail):
        flags.append(dict(level=level, file=file, plate=plate, kind=kind, ref=ref, detail=detail))

    for f, why in skipped:
        flag("ข้าม", f, "", "ไฟล์ไม่ได้นำเข้า", "", why)

    for o in reports:
        failed = [k for k, v in o["checks"].items() if not v]
        recon_fail = [k for k in failed if k != "ทะเบียนตรงชื่อไฟล์"]
        status = "ซ้ำ (ไม่นำเข้า)" if (dedupe and o["duplicate_of"]) else ("PASSED" if not failed else "มีประเด็น")
        files_rows.append(dict(file=o["file"], plate=o["plate"], vehicle_desc=o["vehicle_desc"], brand=o["brand"],
                               vtype=o["vtype"], print_ts=o["print_ts"], n_issue_docs=len(o["issues_hdr"]),
                               n_issue_lines=len(o["issues"]), n_wo=len(o["wo_hdr"]), n_wo_lines=len(o["wo_lines"]),
                               wx=o["totals"].get("wx"), wo=o["totals"].get("wo"), all=o["totals"].get("all"),
                               status=status, failed=", ".join(failed), duplicate_of=o["duplicate_of"],
                               content_hash=o["content_hash"]))
        if o["duplicate_of"]:
            flag("ข้อมูล", o["file"], o["plate"], "ไฟล์ซ้ำ (เนื้อหาเหมือนกันทุกบรรทัด)", "",
                 f"ซ้ำกับ {o['duplicate_of']}" + (" — ตัดออก นำเข้าครั้งเดียว" if dedupe else ""))
        if not o["checks"]["ทะเบียนตรงชื่อไฟล์"]:
            flag("เตือน", o["file"], o["plate"], "ชื่อไฟล์ไม่ตรงทะเบียนในรายงาน", "",
                 f"ชื่อไฟล์ {o['file']} แต่ทะเบียนในรายงาน {o['plate']} — รายงานของทะเบียนตามชื่อไฟล์อาจยังไม่ได้ export")
        if o["same_plate_diff_content"]:
            flag("No", o["file"], o["plate"], "ทะเบียนเดียวกันหลายไฟล์ เนื้อหาต่างกัน", "",
                 "ไฟล์อื่น: " + ", ".join(o["same_plate_diff_content"]))
        if not o["desc_std"]:
            flag("No", o["file"], o["plate"], "คำอธิบายรถไม่ตามรูปแบบ", "", o["vehicle_desc"])
        for k in recon_fail:
            flag("No" if k != "หัว WO = W/O Total" else "เตือน", o["file"], o["plate"], "กระทบยอดไม่ผ่าน", k, "")
        for a in o["anomalies"]:
            flag("No", o["file"], o["plate"], "แถวที่อ่านไม่ออก", f"แถว {a[0]}", f"{a[1]}: {a[2]}")
        if dedupe and o["duplicate_of"]:
            continue

        file_flag = []
        if any(k in recon_fail for k in ("บรรทัดเบิก = WX Total", "หัวเบิก = WX Total", "บรรทัด WO = W/O Total",
                                         "WX + W/O = Total Amount")):
            file_flag.append("ไฟล์กระทบยอดไม่ผ่าน")
        if not o["desc_std"]:
            file_flag.append("คำอธิบายรถไม่ตามรูปแบบ")
        if o["same_plate_diff_content"]:
            file_flag.append("ทะเบียนซ้ำหลายไฟล์")
        base = dict(plate=o["plate"], brand=o["brand"], vtype=o["vtype"], file=o["file"])

        for l in o["issues"]:
            reasons = list(file_flag)
            if (l["total"] or 0) <= 0:
                reasons.append("ยอด ≤ 0")
            issues.append({**base, **{k: l[k] for k in ("doc_no", "line_no", "date", "inv_code", "inv_name", "qty",
                                                         "unit_price", "total", "dept", "store", "mileage_raw",
                                                         "mileage_km", "mileage_status", "mileage_src",
                                                         "remark", "src_row")},
                           "confirm_flag": "No" if reasons else "Yes", "flag_reason": "; ".join(reasons)})
        for h in o["wo_hdr"]:
            reasons = list(file_flag)
            if round(h["net"] or 0, 2) != h["line_sum"]:
                reasons.append(f"ยอดสุทธิหัว {h['net']:,.2f} ≠ ผลรวมบรรทัด {h['line_sum']:,.2f}")
                flag("No", o["file"], o["plate"], "หัว WO ยอดสุทธิไม่ตรงผลรวมบรรทัด", h["doc_no"],
                     f"ยอด {h['amount']:,.2f} / ส่วนลด {h['discount'] or 0:,.2f} / สุทธิ {h['net']:,.2f} / "
                     f"ผลรวมบรรทัด {h['line_sum']:,.2f} (ต่าง {h['line_sum'] - (h['net'] or 0):,.2f})")
            if "," in str(h["doc_no"]) or not h["doc_no"]:
                reasons.append("หัว WO ผูกเลขเอกสารไม่ได้ 1:1")
            h["confirm_flag"] = "No" if reasons else "Yes"
            h["flag_reason"] = "; ".join(reasons)
            wo_hdr.append({**base, **{k: h[k] for k in ("doc_no", "date", "vendor_en", "vendor_th", "remark", "amount",
                                                         "discount", "net", "line_sum", "n_lines", "has_vat_line",
                                                         "mileage_km", "mileage_status", "mileage_src",
                                                         "src_row", "confirm_flag", "flag_reason")}})
        hdr_flag = {h["doc_no"]: h for h in o["wo_hdr"]}
        for l in o["wo_lines"]:
            h = hdr_flag.get(l["doc_no"], {})
            wo_lines.append({**base, **{k: l[k] for k in ("doc_no", "line_no", "date", "item_code", "item_name", "qty",
                                                           "unit_price", "total", "is_vat", "vendor_en",
                                                           "mileage_km", "mileage_status", "src_row")},
                             "confirm_flag": h.get("confirm_flag", "No"), "flag_reason": h.get("flag_reason", "")})
        reads = [(x["date"], x["mileage_km"]) for x in o["issues"] + o["wo_hdr"]
                 if x.get("mileage_km") is not None and x.get("mileage_status") in ("ชัดเจน", "ตีความ") and x["date"]]
        last = max(reads) if reads else (None, None)
        per_vehicle.append(dict(last_mileage=last[1], last_mileage_date=last[0], plate=o["plate"], brand=o["brand"], vtype=o["vtype"], file=o["file"],
                                n_issue_docs=len(o["issues_hdr"]), wx=o["totals"].get("wx") or 0,
                                n_wo=len(o["wo_hdr"]), wo=o["totals"].get("wo") or 0,
                                all=o["totals"].get("all") or 0,
                                first_date=min([x["date"] for x in o["issues"] + o["wo_lines"] if x["date"]], default=None),
                                last_date=max([x["date"] for x in o["issues"] + o["wo_lines"] if x["date"]], default=None)))

    # cross-file: issue doc shared by several plates (info)
    shared = {}
    for l in issues:
        shared.setdefault(l["doc_no"], set()).add(l["plate"])
    n_shared = sum(1 for v in shared.values() if len(v) > 1)

    summary = dict(
        folder=folder, run_at=datetime.now(), n_files_found=len(reports) + len(skipped), n_reports=len(reports),
        n_skipped=len(skipped), n_dup=sum(1 for o in reports if o["duplicate_of"]),
        n_used=len(used), n_plates=len({o["plate"] for o in used}),
        n_pass=sum(1 for o in reports if all(o["checks"].values())),
        wx=round(sum(o["totals"].get("wx") or 0 for o in used), 2),
        wo=round(sum(o["totals"].get("wo") or 0 for o in used), 2),
        n_issue_lines=len(issues), n_issue_docs=len({(l["plate"], l["doc_no"]) for l in issues}),
        n_wo=len(wo_hdr), n_wo_lines=len(wo_lines),
        n_flag_no=sum(1 for r in issues + wo_lines if r["confirm_flag"] == "No"),
        n_shared_issue_docs=n_shared, dedupe=dedupe,
        date_min=min([r["date"] for r in issues + wo_lines if r["date"]], default=None),
        date_max=max([r["date"] for r in issues + wo_lines if r["date"]], default=None),
    )
    summary["all"] = round(summary["wx"] + summary["wo"], 2)
    def _mc(rows):
        c = {"ชัดเจน": 0, "ตีความ": 0, "ต้องตรวจ": 0, "ไม่มี": 0}
        for r in rows:
            c[r.get("mileage_status") or "ไม่มี"] = c.get(r.get("mileage_status") or "ไม่มี", 0) + 1
        return c
    summary["mileage_issue"] = _mc(issues)
    summary["mileage_wo"] = _mc(wo_hdr)
    # post-check: dataset totals = sum of file totals
    summary["chk_issue_total"] = round(sum(r["total"] or 0 for r in issues), 2) == summary["wx"]
    summary["chk_wo_total"] = round(sum(r["total"] or 0 for r in wo_lines), 2) == summary["wo"]
    return dict(summary=summary, files=files_rows, vehicles=per_vehicle, issues=issues,
                wo_hdr=wo_hdr, wo_lines=wo_lines, flags=flags)


# ----------------------------------------------------------------------------- column specs
# (key, header, kind[text|num|money|date|int], width)
COLS = {
    "files": [
        ('file', 'ชื่อไฟล์', 'text', 16),
        ('plate', 'ทะเบียน', 'text', 11),
        ('brand', 'ยี่ห้อ', 'text', 9),
        ('vtype', 'ประเภท', 'text', 9),
        ('vehicle_desc', 'คำอธิบายรถ', 'text', 32),
        ('print_ts', 'เวลาพิมพ์รายงาน', 'text', 22),
        ('n_issue_docs', 'ใบเบิก', 'int', 8),
        ('n_issue_lines', 'บรรทัดเบิก', 'int', 9),
        ('n_wo', 'ใบสั่งซ่อม', 'int', 9),
        ('n_wo_lines', 'บรรทัด WO', 'int', 9),
        ('wx', 'ยอดเบิก (WX)', 'money', 14),
        ('wo', 'ยอดสั่งซ่อม (W/O)', 'money', 15),
        ('all', 'ยอดรวม', 'money', 15),
        ('status', 'สถานะ', 'text', 14),
        ('failed', 'จุดที่ไม่ผ่าน', 'text', 30),
        ('duplicate_of', 'ซ้ำกับไฟล์', 'text', 16),
        ('content_hash', 'Hash เนื้อหา', 'text', 14),
    ],
    "vehicles": [
        ('plate', 'ทะเบียน', 'text', 11),
        ('brand', 'ยี่ห้อ', 'text', 9),
        ('vtype', 'ประเภท', 'text', 9),
        ('last_mileage', 'เลขไมล์ล่าสุด', 'km', 12),
        ('last_mileage_date', 'วันที่เลขไมล์ล่าสุด', 'date', 13),
        ('file', 'ไฟล์', 'text', 16),
        ('first_date', 'วันที่แรก', 'date', 12),
        ('last_date', 'วันที่ล่าสุด', 'date', 12),
        ('n_issue_docs', 'ใบเบิก', 'int', 8),
        ('wx', 'ยอดเบิกพัสดุ', 'money', 15),
        ('n_wo', 'ใบสั่งซ่อม', 'int', 9),
        ('wo', 'ยอดสั่งซ่อมภายนอก', 'money', 16),
        ('all', 'ยอดรวม', 'money', 16),
    ],
    "issues": [
        ('plate', 'ทะเบียน', 'text', 11),
        ('doc_no', 'เลขที่ใบเบิก', 'text', 12),
        ('line_no', 'บรรทัด', 'int', 7),
        ('date', 'วันที่', 'date', 12),
        ('mileage_km', 'เลขไมล์', 'km', 11),
        ('mileage_status', 'สถานะเลขไมล์', 'text', 10),
        ('inv_code', 'รหัสพัสดุ', 'text', 17),
        ('inv_name', 'ชื่อพัสดุ', 'text', 40),
        ('qty', 'จำนวน', 'num', 9),
        ('unit_price', 'ราคา/หน่วย', 'money', 12),
        ('total', 'ยอดสุทธิ', 'money', 13),
        ('dept', 'หน่วยงาน', 'text', 24),
        ('store', 'คลัง', 'text', 26),
        ('mileage_src', 'ที่มาเลขไมล์', 'text', 26),
        ('mileage_raw', 'เลขไมล์ (ต้นฉบับ)', 'text', 12),
        ('remark', 'หมายเหตุใบเบิก', 'text', 36),
        ('brand', 'ยี่ห้อ', 'text', 8),
        ('vtype', 'ประเภท', 'text', 8),
        ('confirm_flag', 'Confirm Flag', 'text', 8),
        ('flag_reason', 'เหตุผล Flag', 'text', 30),
        ('file', 'ไฟล์ต้นทาง', 'text', 15),
        ('src_row', 'แถวต้นทาง', 'int', 8),
    ],
    "wo_hdr": [
        ('plate', 'ทะเบียน', 'text', 11),
        ('doc_no', 'เลขที่ใบสั่งซ่อม', 'text', 13),
        ('date', 'วันที่', 'date', 12),
        ('mileage_km', 'เลขไมล์', 'km', 11),
        ('mileage_status', 'สถานะเลขไมล์', 'text', 10),
        ('vendor_en', 'ผู้ขาย (EN)', 'text', 40),
        ('vendor_th', 'ผู้ขาย (TH)', 'text', 36),
        ('remark', 'หมายเหตุ', 'text', 36),
        ('mileage_src', 'ที่มาเลขไมล์', 'text', 22),
        ('amount', 'ยอด', 'money', 13),
        ('discount', 'ส่วนลด', 'money', 10),
        ('net', 'ยอดสุทธิ', 'money', 13),
        ('line_sum', 'ผลรวมบรรทัด', 'money', 13),
        ('n_lines', 'จำนวนบรรทัด', 'int', 9),
        ('has_vat_line', 'มีบรรทัด VAT', 'text', 10),
        ('brand', 'ยี่ห้อ', 'text', 8),
        ('vtype', 'ประเภท', 'text', 8),
        ('confirm_flag', 'Confirm Flag', 'text', 8),
        ('flag_reason', 'เหตุผล Flag', 'text', 34),
        ('file', 'ไฟล์ต้นทาง', 'text', 15),
        ('src_row', 'แถวต้นทาง', 'int', 8),
    ],
    "wo_lines": [
        ('plate', 'ทะเบียน', 'text', 11),
        ('doc_no', 'เลขที่ใบสั่งซ่อม', 'text', 13),
        ('line_no', 'บรรทัด', 'int', 7),
        ('date', 'วันที่', 'date', 12),
        ('mileage_km', 'เลขไมล์', 'km', 11),
        ('mileage_status', 'สถานะเลขไมล์', 'text', 10),
        ('item_code', 'รหัสรายการ', 'text', 17),
        ('item_name', 'ชื่อรายการ', 'text', 40),
        ('qty', 'จำนวน', 'num', 9),
        ('unit_price', 'ราคา/หน่วย', 'money', 12),
        ('total', 'ยอด', 'money', 13),
        ('is_vat', 'บรรทัด VAT', 'text', 9),
        ('vendor_en', 'ผู้ขาย', 'text', 36),
        ('confirm_flag', 'Confirm Flag', 'text', 8),
        ('flag_reason', 'เหตุผล Flag', 'text', 30),
        ('file', 'ไฟล์ต้นทาง', 'text', 15),
        ('src_row', 'แถวต้นทาง', 'int', 8),
    ],
    "flags": [
        ('level', 'ระดับ', 'text', 8),
        ('kind', 'ประเด็น', 'text', 34),
        ('file', 'ไฟล์', 'text', 22),
        ('plate', 'ทะเบียน', 'text', 11),
        ('ref', 'อ้างอิง', 'text', 20),
        ('detail', 'รายละเอียด', 'text', 80),
    ],
}
SHEETS = [("files", "สรุปไฟล์"), ("vehicles", "สรุปรายคัน"), ("issues", "เบิกพัสดุ"),
          ("wo_hdr", "ใบสั่งซ่อม_หัว"), ("wo_lines", "ใบสั่งซ่อม_บรรทัด"), ("flags", "ประเด็น_Flag")]


def fmt_cell(v, kind):
    if v is None or v == "":
        return ""
    if isinstance(v, bool):
        return "ใช่" if v else ""
    if kind == "date" and isinstance(v, date):
        return v.strftime("%d/%m/%Y")
    if kind == "money" and isinstance(v, (int, float)):
        return f"{v:,.2f}"
    if kind == "km" and isinstance(v, (int, float)):
        return f"{v:,.0f}"
    if kind == "num" and isinstance(v, (int, float)):
        return f"{v:,.4f}".rstrip("0").rstrip(".")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


# ----------------------------------------------------------------------------- export
def export_excel(ds, out_dir):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    os.makedirs(out_dir, exist_ok=True)
    fn = os.path.join(out_dir, f"Vehicle_Reports_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ข้อมูลการนำเข้า"
    s = ds["summary"]
    hfont = Font(bold=True, color="FFFFFF")
    hfill = PatternFill("solid", fgColor="1F4E78")
    nofill = PatternFill("solid", fgColor="FCE4D6")
    info = [
        ("โปรแกรม", f"Vehicle Report Importer v{APP_VERSION}"),
        ("โฟลเดอร์ต้นทาง", s["folder"]),
        ("เวลาอ่านข้อมูล", s["run_at"]),
        ("ไฟล์ที่พบ (.xls/.xlsx)", s["n_files_found"]),
        ("ไฟล์รายงานที่อ่านได้", s["n_reports"]),
        ("ไฟล์ที่ข้าม (ไม่ใช่รูปแบบรายงาน/อ่านไม่ได้)", s["n_skipped"]),
        ("ไฟล์ซ้ำ (เนื้อหาเหมือนกัน)", s["n_dup"]),
        ("ตัดไฟล์ซ้ำออก", "ใช่" if s["dedupe"] else "ไม่"),
        ("ไฟล์ที่ใช้ในข้อมูล", s["n_used"]),
        ("จำนวนทะเบียน", s["n_plates"]),
        ("ไฟล์ที่ผ่านกระทบยอดครบ 6 จุด", s["n_pass"]),
        ("ช่วงวันที่ข้อมูล", f"{s['date_min']:%d/%m/%Y} – {s['date_max']:%d/%m/%Y}" if s["date_min"] else ""),
        ("ใบเบิก / บรรทัดเบิก", f"{s['n_issue_docs']:,} / {s['n_issue_lines']:,}"),
        ("ใบสั่งซ่อม / บรรทัด", f"{s['n_wo']:,} / {s['n_wo_lines']:,}"),
        ("ยอดเบิกพัสดุ (WX)", s["wx"]),
        ("ยอดสั่งซ่อมภายนอก (W/O)", s["wo"]),
        ("ยอดรวม", s["all"]),
        ("ตรวจ: ผลรวมบรรทัดเบิก = ยอด WX รวม", "PASSED" if s["chk_issue_total"] else "FAILED"),
        ("ตรวจ: ผลรวมบรรทัด WO = ยอด W/O รวม", "PASSED" if s["chk_wo_total"] else "FAILED"),
        ("บรรทัด Confirm Flag = No", s["n_flag_no"]),
        ("เลขใบเบิกที่ใช้กับหลายทะเบียน", s["n_shared_issue_docs"]),
        ("เลขไมล์ — บรรทัดเบิก", " / ".join(f"{k} {v:,}" for k, v in s["mileage_issue"].items())),
        ("เลขไมล์ — ใบสั่งซ่อม", " / ".join(f"{k} {v:,}" for k, v in s["mileage_wo"].items())),
        ("วิธีได้เลขไมล์", "ใบเบิก: คอลัมน์ 'เลขไมล์' ของรายงานก่อน ถ้าว่างจึงดึงจากหมายเหตุ ; ใบสั่งซ่อม: ดึงจากหมายเหตุ "
                          "(รูปแบบ '…กม.', 'เลขไมล์ …', '…@MTD-SR', '//…') ; ชัดเจน = รูปแบบตรง ; "
                          "ตีความ = '//884.215' อ่านจุดเป็นตัวคั่นหลักพัน ; ต้องตรวจ = คอลัมน์กับหมายเหตุไม่ตรงกัน / ค่า < 100 / "
                          "ตัวเลขล้วนไม่มีหน่วย ; ช่องว่าง = ไม่พบเลขไมล์ (ไม่เดา)"),
        ("หมายเหตุ", "ยอดเงินตามรายงานต้นทาง (ใบสั่งซ่อมมีบรรทัด 'ภาษีมูลค่าเพิ่ม' แยก) ; "
                     "แถว Flag No ต้องให้คุณแขกตรวจก่อนนำเข้าฐานข้อมูลหลัก ; ไฟล์นี้ไม่ใช่ฐานข้อมูลหลัก"),
    ]
    ws.append(["รายการ", "ค่า"])
    for c in ws[1]:
        c.font = hfont; c.fill = hfill
    for k, v in info:
        ws.append([k, v])
        cell = ws.cell(ws.max_row, 2)
        if isinstance(v, datetime):
            cell.number_format = "DD/MM/YYYY HH:MM:SS"
        elif isinstance(v, float):
            cell.number_format = "#,##0.00"
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 70

    for key, title in SHEETS:
        w = wb.create_sheet(title)
        cols = COLS[key]
        w.append([c[1] for c in cols])
        for c in w[1]:
            c.font = hfont; c.fill = hfill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        flag_col = next((i for i, c in enumerate(cols, 1) if c[0] == "confirm_flag"), None)
        for r in ds[key]:
            vals = []
            for k, _, kind, _ in cols:
                v = r.get(k)
                if isinstance(v, bool):
                    v = "ใช่" if v else ""
                if kind == "text" and isinstance(v, float) and v.is_integer():
                    v = str(int(v))
                if isinstance(v, str):
                    v = ILLEGAL_CHARACTERS_RE.sub("", v)
                vals.append(v)
            w.append(vals)
            if flag_col and r.get("confirm_flag") == "No":
                w.cell(w.max_row, flag_col).fill = nofill
        for i, (k, _, kind, width) in enumerate(cols, 1):
            L = get_column_letter(i)
            w.column_dimensions[L].width = width
            fmt = {"date": "DD/MM/YYYY", "money": "#,##0.00", "num": "#,##0.####", "int": "0", "km": "#,##0"}.get(kind)
            if fmt:
                for cell in w[L][1:]:
                    cell.number_format = fmt
        w.freeze_panes = "B2" if key != "flags" else "A2"
        if w.max_row > 1:
            w.auto_filter.ref = w.dimensions
    wb.save(fn)
    return fn


# ----------------------------------------------------------------------------- GUI
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    root = tk.Tk()
    root.title(f"Vehicle Report Importer v{APP_VERSION} — รายงานซ่อม/เบิกพัสดุรายคัน")
    root.geometry("1280x760")
    try:
        root.state("zoomed")
    except Exception:
        pass
    style = ttk.Style()
    try:
        style.theme_use("vista" if sys.platform == "win32" else "clam")
    except Exception:
        pass
    base_font = ("Tahoma", 10)
    style.configure("Treeview", font=base_font, rowheight=22)
    style.configure("Treeview.Heading", font=("Tahoma", 10, "bold"))
    style.configure("TLabel", font=base_font)
    style.configure("TButton", font=base_font)
    style.configure("Sum.TLabel", font=("Tahoma", 10, "bold"), foreground="#1F4E78")

    state = {"ds": None}
    q = queue.Queue()

    # --- top: folders
    top = ttk.Frame(root, padding=8)
    top.pack(fill="x")
    src_var = tk.StringVar(value=DEFAULT_SRC if os.path.isdir(DEFAULT_SRC) else "")
    out_var = tk.StringVar(value=DEFAULT_OUT)
    dedupe_var = tk.BooleanVar(value=True)

    ttk.Label(top, text="โฟลเดอร์รายงาน:").grid(row=0, column=0, sticky="w")
    ttk.Entry(top, textvariable=src_var, width=70).grid(row=0, column=1, sticky="we", padx=4)
    ttk.Button(top, text="Browse…", command=lambda: _browse(src_var, "เลือกโฟลเดอร์รายงานรายคัน")).grid(row=0, column=2)
    btn_load = ttk.Button(top, text="① อ่านข้อมูล", width=16)
    btn_load.grid(row=0, column=3, padx=(10, 0))
    ttk.Checkbutton(top, text="ตัดไฟล์ซ้ำ (เนื้อหาเหมือนกัน)", variable=dedupe_var).grid(row=0, column=4, padx=8)

    ttk.Label(top, text="โฟลเดอร์ Export:").grid(row=1, column=0, sticky="w", pady=(6, 0))
    ttk.Entry(top, textvariable=out_var, width=70).grid(row=1, column=1, sticky="we", padx=4, pady=(6, 0))
    ttk.Button(top, text="Browse…", command=lambda: _browse(out_var, "เลือกโฟลเดอร์เก็บไฟล์ Excel")).grid(row=1, column=2, pady=(6, 0))
    btn_export = ttk.Button(top, text="② Export Excel", width=16, state="disabled")
    btn_export.grid(row=1, column=3, padx=(10, 0), pady=(6, 0))
    btn_open = ttk.Button(top, text="เปิดโฟลเดอร์ Export", state="normal")
    btn_open.grid(row=1, column=4, padx=8, pady=(6, 0), sticky="w")
    top.columnconfigure(1, weight=1)

    prog = ttk.Progressbar(root, mode="determinate")
    prog.pack(fill="x", padx=8)
    status_var = tk.StringVar(value="เลือกโฟลเดอร์รายงาน แล้วกด ① อ่านข้อมูล")
    ttk.Label(root, textvariable=status_var).pack(fill="x", padx=8, pady=(2, 0))
    sum_var = tk.StringVar(value="")
    ttk.Label(root, textvariable=sum_var, style="Sum.TLabel", wraplength=1250, justify="left").pack(fill="x", padx=8, pady=4)

    # --- filter
    ff = ttk.Frame(root, padding=(8, 0))
    ff.pack(fill="x")
    ttk.Label(ff, text="ค้นหา (ทะเบียน/เลขเอกสาร/ข้อความ):").pack(side="left")
    find_var = tk.StringVar()
    ent_find = ttk.Entry(ff, textvariable=find_var, width=36)
    ent_find.pack(side="left", padx=4)
    only_no_var = tk.BooleanVar(value=False)
    ttk.Checkbutton(ff, text="เฉพาะ Flag No / ประเด็น", variable=only_no_var,
                    command=lambda: refresh_current()).pack(side="left", padx=8)
    count_var = tk.StringVar()
    ttk.Label(ff, textvariable=count_var).pack(side="right")

    nb = ttk.Notebook(root)
    nb.pack(fill="both", expand=True, padx=8, pady=6)
    trees = {}
    for key, title in SHEETS:
        fr = ttk.Frame(nb)
        nb.add(fr, text=title)
        cols = COLS[key]
        tv = ttk.Treeview(fr, columns=[c[0] for c in cols], show="headings", selectmode="extended")
        for k, h, kind, w in cols:
            tv.heading(k, text=h, command=lambda k=k, key=key: sort_by(key, k))
            tv.column(k, width=int(w * 8.5), anchor=("e" if kind in ("money", "num", "int", "km") else "w"), stretch=False)
        vs = ttk.Scrollbar(fr, orient="vertical", command=tv.yview)
        hs = ttk.Scrollbar(fr, orient="horizontal", command=tv.xview)
        tv.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        tv.grid(row=0, column=0, sticky="nsew"); vs.grid(row=0, column=1, sticky="ns"); hs.grid(row=1, column=0, sticky="ew")
        fr.rowconfigure(0, weight=1); fr.columnconfigure(0, weight=1)
        tv.tag_configure("no", background="#FCE4D6")
        tv.tag_configure("dup", foreground="#888888")
        trees[key] = tv
    sort_state = {}

    def _browse(var, title):
        d = filedialog.askdirectory(title=title, initialdir=var.get() or "D:\\")
        if d:
            var.set(os.path.normpath(d))

    def current_key():
        return SHEETS[nb.index(nb.select())][0]

    def row_is_issue(key, r):
        if key in ("issues", "wo_hdr", "wo_lines"):
            return r.get("confirm_flag") == "No"
        if key == "files":
            return r.get("status") != "PASSED"
        if key == "flags":
            return r.get("level") in ("No", "ข้าม")
        return False

    def fill(key):
        tv = trees[key]
        tv.delete(*tv.get_children())
        ds = state["ds"]
        if not ds:
            return
        rows = ds[key]
        txt = find_var.get().strip().lower()
        cols = COLS[key]
        shown = 0
        for r in rows:
            if only_no_var.get() and not row_is_issue(key, r):
                continue
            vals = [fmt_cell(r.get(k), kind) for k, _, kind, _ in cols]
            if txt and not any(txt in str(v).lower() for v in vals):
                continue
            tag = ()
            if key == "files" and str(r.get("status", "")).startswith("ซ้ำ"):
                tag = ("dup",)
            elif row_is_issue(key, r):
                tag = ("no",)
            tv.insert("", "end", values=vals, tags=tag)
            shown += 1
        count_var.set(f"แสดง {shown:,} / {len(rows):,} แถว")

    def refresh_current(*_):
        fill(current_key())

    def sort_by(key, col):
        ds = state["ds"]
        if not ds:
            return
        rev = sort_state.get((key, col), False)
        kind = next(c[2] for c in COLS[key] if c[0] == col)

        def kf(r):
            v = r.get(col)
            if kind in ("money", "num", "int", "km"):
                return (v is None, v or 0)
            if kind == "date":
                return (v is None, v or date.min)
            return (v is None, str(v or ""))
        ds[key].sort(key=kf, reverse=rev)
        sort_state[(key, col)] = not rev
        fill(key)

    nb.bind("<<NotebookTabChanged>>", refresh_current)
    find_var.trace_add("write", lambda *_: root.after(300, refresh_current))

    def show_summary():
        s = state["ds"]["summary"]
        dr = f"{s['date_min']:%d/%m/%Y} – {s['date_max']:%d/%m/%Y}" if s["date_min"] else "-"
        sum_var.set(
            f"ไฟล์ {s['n_files_found']} | อ่านได้ {s['n_reports']} | ข้าม {s['n_skipped']} | ซ้ำ {s['n_dup']} | "
            f"ใช้ {s['n_used']} ไฟล์ = {s['n_plates']} ทะเบียน | กระทบยอดผ่านครบ {s['n_pass']}/{s['n_reports']} | "
            f"ช่วงวันที่ {dr}\n"
            f"เบิกพัสดุ {s['n_issue_docs']:,} ใบ/{s['n_issue_lines']:,} บรรทัด = {s['wx']:,.2f} | "
            f"สั่งซ่อม {s['n_wo']:,} ใบ/{s['n_wo_lines']:,} บรรทัด = {s['wo']:,.2f} | รวม {s['all']:,.2f} บาท | "
            f"Flag No {s['n_flag_no']:,} บรรทัด | ประเด็น {len(state['ds']['flags'])} รายการ")

    def do_load():
        folder = src_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("โฟลเดอร์ไม่ถูกต้อง", "กรุณาเลือกโฟลเดอร์รายงาน")
            return
        btn_load.config(state="disabled"); btn_export.config(state="disabled")
        prog["value"] = 0
        status_var.set("กำลังอ่าน…")
        dd = dedupe_var.get()

        def work():
            try:
                ds = scan_folder(folder, progress=lambda n, t, f: q.put(("prog", n, t, f)), dedupe=dd)
                q.put(("done", ds))
            except Exception:
                q.put(("err", traceback.format_exc()))
        threading.Thread(target=work, daemon=True).start()

    def poll():
        try:
            while True:
                m = q.get_nowait()
                if m[0] == "prog":
                    _, n, t, f = m
                    prog["maximum"] = t; prog["value"] = n
                    status_var.set(f"กำลังอ่าน {n}/{t}: {f}")
                elif m[0] == "done":
                    state["ds"] = m[1]
                    for key, _ in SHEETS:
                        fill(key)
                    refresh_current()
                    show_summary()
                    status_var.set("อ่านข้อมูลเสร็จ — ตรวจแท็บ 'ประเด็น_Flag' แล้วกด ② Export Excel")
                    btn_load.config(state="normal"); btn_export.config(state="normal")
                elif m[0] == "err":
                    btn_load.config(state="normal")
                    status_var.set("เกิดข้อผิดพลาด")
                    messagebox.showerror("ผิดพลาด", m[1][-1500:])
                elif m[0] == "exported":
                    btn_export.config(state="normal")
                    status_var.set(f"Export แล้ว: {m[1]}")
                    if messagebox.askyesno("Export เสร็จ", f"บันทึกไฟล์แล้ว\n{m[1]}\n\nเปิดไฟล์เลยหรือไม่?"):
                        _open(m[1])
                elif m[0] == "export_err":
                    btn_export.config(state="normal")
                    status_var.set("Export ไม่สำเร็จ")
                    messagebox.showerror("Export ไม่สำเร็จ", m[1][-1500:])
        except queue.Empty:
            pass
        root.after(100, poll)

    def do_export():
        if not state["ds"]:
            return
        out = out_var.get().strip() or DEFAULT_OUT
        btn_export.config(state="disabled")
        status_var.set("กำลังเขียน Excel…")

        def work():
            try:
                q.put(("exported", export_excel(state["ds"], out)))
            except Exception:
                q.put(("export_err", traceback.format_exc()))
        threading.Thread(target=work, daemon=True).start()

    def _open(p):
        try:
            if sys.platform == "win32":
                os.startfile(p)
            else:
                import subprocess
                subprocess.Popen(["xdg-open", p])
        except Exception as e:
            messagebox.showerror("เปิดไม่ได้", str(e))

    def open_out():
        out = out_var.get().strip() or DEFAULT_OUT
        os.makedirs(out, exist_ok=True)
        _open(out)

    btn_load.config(command=do_load)
    btn_export.config(command=do_export)
    btn_open.config(command=open_out)
    root.after(100, poll)
    root.mainloop()


# ----------------------------------------------------------------------------- CLI (ทดสอบ/ใช้งานไม่ผ่าน GUI)
if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "--cli":
        src = sys.argv[2]
        out = sys.argv[3] if len(sys.argv) > 3 else DEFAULT_OUT
        ds = scan_folder(src, progress=None, dedupe="--keep-dup" not in sys.argv)
        print(json.dumps(ds["summary"], default=str, ensure_ascii=False, indent=1))
        print("flags:", len(ds["flags"]))
        print("saved:", export_excel(ds, out))
    else:
        run_gui()
