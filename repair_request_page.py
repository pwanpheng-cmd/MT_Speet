# -*- coding: utf-8 -*-
"""หน้า 'ฐานข้อมูลใบแจ้งซ่อม' (หัวลาก / หางพ่วง / ตู้เย็น) — ใช้ใน vehicle_report_web.py"""
import os, io, sqlite3, tempfile
from datetime import datetime

import pandas as pd
import streamlit as st

import repair_request_db as rr

TYPES = ["หัวลาก", "หางพ่วง", "ตู้เย็น"]


def _df(rows, cols=None):
    df = pd.DataFrame(rows)
    if cols:
        df = df[[c for c in cols if c in df.columns]]
    return df


COLS_TH = {"doc_no": "เลขที่เอกสาร", "asset_type": "ประเภท", "report_date": "วันที่แจ้งซ่อม", "report_time": "เวลา",
           "report_time_raw": "เวลา (ต้นฉบับ)", "open_date": "วันที่เปิดเอกสาร", "asset_code": "ทะเบียน/เบอร์ตู้",
           "asset_raw": "ทะเบียน/เบอร์ตู้ (ต้นฉบับ)", "brand": "ยี่ห้อ", "employee_code": "รหัสพนักงาน",
           "status": "สถานะซ่อม", "in_iso": "ใน ISO", "report_month": "เดือน", "source_sheet": "ชีตต้นทาง", "source_row": "แถวต้นทาง",
           "data_note": "หมายเหตุข้อมูล", "flag_reason": "เหตุผล", "confirm_flag": "Confirm Flag"}
SHOW = ["doc_no", "asset_type", "report_date", "report_time", "asset_code", "in_iso", "brand", "status", "employee_code",
        "open_date", "data_note", "source_sheet", "source_row"]


# ----------------------------------------------------------------------------- ① เลือกไฟล์ฐานข้อมูล (ใช้ร่วม 2 หน้า)
def _pick_db_file(initial):
    """เปิดหน้าต่าง Browse ไฟล์ .sqlite ของ Windows (เฉพาะรันบนเครื่อง)"""
    import subprocess, sys
    init_dir = os.path.dirname(initial) if initial else ""
    code = ("import tkinter as tk; from tkinter import filedialog; r=tk.Tk(); r.withdraw(); "
            "r.wm_attributes('-topmost',1); "
            f"p=filedialog.askopenfilename(initialdir={init_dir!r}, title='เลือกไฟล์ฐานข้อมูล', "
            "filetypes=[('SQLite database','*.sqlite *.db *.sqlite3'),('All files','*.*')]); print(p or '', end='')")
    try:
        res = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, timeout=600, encoding="utf-8")
        return os.path.normpath(res.stdout.strip()) if res.stdout.strip() else None
    except Exception as e:
        st.sidebar.error(f"เปิดหน้าต่างเลือกไฟล์ไม่ได้: {e} — พิมพ์ path เองได้")
        return None


def _sha(path):
    import hashlib
    if not path or not os.path.exists(path):
        return None
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def db_selector(is_local, app_dir, full):
    """แสดงส่วน '① ไฟล์ฐานข้อมูล' ใน sidebar แล้วคืน path ฐานข้อมูลที่ใช้
    - บนเครื่อง: Browse ไฟล์ .sqlite → อ่าน/บันทึกที่ไฟล์นั้นโดยตรง (สำรองอัตโนมัติก่อนนำเข้า)
    - บนเว็บ:   Browse (อัปโหลด) ไฟล์ .sqlite → ทำงานบนสำเนาชั่วคราว → ดาวน์โหลดกลับไปวางทับไฟล์เดิม"""
    ss = st.session_state
    with st.sidebar:
        st.markdown("**① ไฟล์ฐานข้อมูล**")
        if is_local:
            ss.setdefault("rr_db", os.path.join(app_dir, "data", "Repair_Request_v2.sqlite"))   # ไม่ใช่ key ของ widget → คงค่าข้ามหน้า
            c1, c2 = st.columns([4, 1.4])
            if c2.button("📁", key="b_db", help="Browse ไฟล์ฐานข้อมูล (.sqlite)", **full):
                p = _pick_db_file(ss["rr_db"])
                if p:
                    ss["rr_db"] = p
                    st.rerun()
            v = c1.text_input("ไฟล์ฐานข้อมูล (.sqlite)", value=ss["rr_db"], label_visibility="collapsed").strip().strip('"')
            if v and v != ss["rr_db"]:
                ss["rr_db"] = v
            db = ss["rr_db"]
            if os.path.exists(db):
                st.caption(f"ใช้ไฟล์นี้โดยตรง · {os.path.getsize(db)/1024/1024:,.1f} MB · สำรองอัตโนมัติที่ backup\\ ก่อนนำเข้าทุกครั้ง")
            else:
                st.warning("ยังไม่มีไฟล์นี้ — จะสร้างฐานใหม่เมื่อนำเข้าไฟล์ชุดแรก")
            return db

        # ---- เว็บ
        if "rr_db_tmp" not in ss:
            ss["rr_db_tmp"] = os.path.join(tempfile.mkdtemp(prefix="rrdb_"), "Repair_Request_v2.sqlite")
            ss["rr_db_name"] = None
        db = ss["rr_db_tmp"]
        up = st.file_uploader("Browse ไฟล์ฐานข้อมูล (.sqlite) จากเครื่อง", type=["sqlite", "db", "sqlite3"], key="rr_db_up")
        if up is not None:
            fid = getattr(up, "file_id", None) or f"{up.name}:{up.size}"
            if ss.get("rr_db_fid") != fid:
                data = up.getvalue()
                if not data.startswith(b"SQLite format 3"):
                    st.error("ไฟล์นี้ไม่ใช่ฐานข้อมูล SQLite")
                else:
                    with open(db, "wb") as fh:
                        fh.write(data)
                    ss["rr_db_fid"], ss["rr_db_name"] = fid, up.name
                    ss["rr_db_sha0"] = _sha(db)
                    ss["rr_db_at"] = datetime.now().strftime("%d/%m/%Y %H:%M")
        if ss.get("rr_db_name"):
            st.caption(f"กำลังใช้: **{ss['rr_db_name']}** (เปิดเมื่อ {ss['rr_db_at']}) — เป็นสำเนาชั่วคราวบนเว็บ")
            changed = _sha(db) != ss.get("rr_db_sha0")
            if changed:
                st.warning("มีการเปลี่ยนแปลงในฐานข้อมูล — ดาวน์โหลดแล้ววางทับไฟล์เดิมในเครื่อง (สำรองไฟล์เดิมก่อน) "
                           "ไม่เช่นนั้นข้อมูลจะหายเมื่อปิดเว็บ")
            with open(db, "rb") as fh:
                st.download_button("ดาวน์โหลดฐานข้อมูล (บันทึกกลับเครื่อง)", fh.read(), file_name=ss["rr_db_name"],
                                   icon="⬇️", type="primary" if changed else "secondary", key="rr_db_dl", **full)
        else:
            st.info("ยังไม่ได้เลือกไฟล์ — Browse ไฟล์ Repair_Request_v2.sqlite จากเครื่อง (เช่น D:\\MT_Speed\\data) "
                    "หรือใช้ฐานว่างชั่วคราวเพื่อตรวจไฟล์")
            if os.path.exists(db):
                with open(db, "rb") as fh:
                    st.download_button("ดาวน์โหลดฐานข้อมูลชั่วคราว", fh.read(), file_name="Repair_Request_v2.sqlite",
                                       icon="⬇️", key="rr_db_dl0", **full)
        return db


def _th(df):
    df = df.copy()
    for c in ("report_date", "open_date"):
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce").dt.strftime("%d/%m/%Y")
    return df.astype(object).where(df.notna(), "").rename(columns=COLS_TH)


def _analyze(db_path, records):
    if os.path.exists(db_path):
        con = rr.connect(db_path)
    else:
        con = sqlite3.connect(":memory:")
        con.row_factory = sqlite3.Row
        con.executescript(rr.SCHEMA_SQL)
    try:
        return rr.analyze(con, records)
    finally:
        con.close()


def _excel_bytes(sheets):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, df in sheets.items():
            (df if len(df) else pd.DataFrame({"ไม่มีข้อมูล": []})).to_excel(xw, sheet_name=name[:31], index=False)
    return buf.getvalue()


KPI_DETAIL_SHEETS = [("หัวลาก", "M1_ซ่อมรถ"), ("หางพ่วง", "M2_ซ่อมหาง"), ("ตู้เย็น", "M3_ซ่อมตู้")]
KPI_DETAIL_COLS = [("doc_no", "เลขที่เอกสาร", 18), ("open_date", "วันที่เปิดเอกสาร", 14), ("report_date", "วันที่แจ้งซ่อม", 14),
                   ("report_time", "เวลา", 8), ("report_time_raw", "เวลา (ต้นฉบับ)", 12), ("asset_raw", "ทะเบียน / เบอร์ตู้", 34),
                   ("asset_code", "ทะเบียน/เบอร์ตู้ (รหัส)", 18), ("in_iso", "ใน ISO", 9), ("brand", "ยี่ห้อ", 12), ("employee_code", "รหัสพนักงาน", 13),
                   ("status", "สถานะซ่อม", 11), ("report_month", "เดือน", 9), ("data_note", "หมายเหตุข้อมูล", 30)]


def _kpi_export_bytes(db, pv, pr):
    """Export KPI + รายละเอียดใบแจ้งซ่อม M1/M2/M3 แยกชีต (วันที่เป็นวันที่จริง รูปแบบ DD/MM/YYYY)"""
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    counts = {}
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        tot_v = pv.copy(); tot_r = pr.copy()
        for df_, name in ((tot_v, "คันต่อเดือน"), (tot_r, "ใบแจ้งซ่อมต่อเดือน")):
            out = df_.rename_axis("เดือน").reset_index()
            out["รวม"] = out[TYPES].sum(axis=1)
            if name == "ใบแจ้งซ่อมต่อเดือน":
                out.loc[len(out)] = ["รวม"] + [int(out[c].sum()) for c in TYPES + ["รวม"]]
            out.to_excel(xw, sheet_name=name, index=False)
        for t, sheet in KPI_DETAIL_SHEETS:
            rows = rr.fetch(db, "SELECT * FROM repair_request WHERE asset_type=? AND in_iso='Yes' ORDER BY report_date, doc_no", t)
            df = pd.DataFrame(rows, columns=[c for c, _, _ in KPI_DETAIL_COLS]) if not rows else \
                pd.DataFrame(rows)[[c for c, _, _ in KPI_DETAIL_COLS]]
            for c in ("open_date", "report_date"):
                df[c] = pd.to_datetime(df[c], errors="coerce")
            df.columns = [h for _, h, _ in KPI_DETAIL_COLS]
            df.to_excel(xw, sheet_name=sheet, index=False)
            counts[sheet] = len(df)
        note = pd.DataFrame({"รายการ": ["สร้างเมื่อ", "ไฟล์ฐานข้อมูล", "คันต่อเดือน", "ใบแจ้งซ่อมต่อเดือน"] + [s_ for _, s_ in KPI_DETAIL_SHEETS] + ["หมายเหตุ"],
                             "รายละเอียด": [datetime.now().strftime("%d/%m/%Y %H:%M"), os.path.basename(db),
                                            "จำนวนคัน/ตู้ไม่ซ้ำต่อเดือน (คิดเฉพาะ ISO = Yes · ตามวันที่แจ้งซ่อม ทุกสถานะ)", "จำนวนใบแจ้งซ่อมต่อเดือน (เฉพาะ ISO = Yes)"]
                                           + [f"รายละเอียดใบแจ้งซ่อม{t} {counts[s_]:,} แถว" + (" (เฉพาะ ISO = Yes)" if t == "หัวลาก" else "") for t, s_ in KPI_DETAIL_SHEETS]
                                           + ["ไม่รวมรายการที่ยังรอตรวจ · หัวลากไม่รวมรายการที่ ISO = No"]})
        note.to_excel(xw, sheet_name="ข้อมูลการ Export", index=False)
        hdr_fill = PatternFill("solid", fgColor="DCE9F5")
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.font = Font(bold=True); cell.fill = hdr_fill; cell.alignment = Alignment(horizontal="center", vertical="center")
            if ws.title in dict((s_, 1) for _, s_ in KPI_DETAIL_SHEETS):
                for i, (_, _, wdt) in enumerate(KPI_DETAIL_COLS, 1):
                    ws.column_dimensions[get_column_letter(i)].width = wdt
                for col in ("B", "C"):
                    for cell in ws[col][1:]:
                        cell.number_format = "dd/mm/yyyy"
                if ws.max_row > 1:
                    ws.auto_filter.ref = ws.dimensions
            else:
                for i in range(1, ws.max_column + 1):
                    ws.column_dimensions[get_column_letter(i)].width = 22 if ws.title != "ข้อมูลการ Export" else (24 if i == 1 else 60)
    return buf.getvalue(), counts


def render(is_local, user, full, app_dir, kpi_card):
    ss = st.session_state
    db = db_selector(is_local, app_dir, full)

    st.markdown('<div class="app-hero"><h1>ฐานข้อมูลใบแจ้งซ่อม</h1>'
                '<p>หัวลาก · หางพ่วง · ตู้เย็น — นำเข้าแบบกันซ้ำด้วยเลขที่เอกสาร · แถวที่ข้อมูลไม่สอดคล้องรอ DATA_OWNER ตัดสิน</p></div>',
                unsafe_allow_html=True)

    s = rr.summary(db)
    empty = not s or s["total"] == 0
    k = st.columns(5)
    if empty:
        kpi_card(k[0], "สถานะฐานข้อมูล", "ยังว่าง", "นำเข้าไฟล์ชุดแรกเพื่อสร้างฐานหลัก", "warn")
    else:
        bt = {r["asset_type"]: r for r in s["by_type"]}
        kpi_card(k[0], "ใบแจ้งซ่อมทั้งหมด", f"{s['total']:,}", f"{min(r['d1'] for r in s['by_type'])} – {max(r['d2'] for r in s['by_type'])}")
        for i, t in enumerate(TYPES, 1):
            r = bt.get(t)
            kpi_card(k[i], t, f"{r['n']:,}" if r else "0", f"{r['assets']:,} คัน/ตู้" if r else "")
        kpi_card(k[4], "รอตรวจ", f"{s['review_open']:,}", "รอ DATA_OWNER ตัดสิน", "warn" if s["review_open"] else "ok")
    st.write("")

    t_imp, t_data, t_kpi, t_rev, t_log = st.tabs(["📥 นำเข้าไฟล์", "📋 ข้อมูลในฐาน", "📊 KPI คัน/เดือน",
                                                  f"⚠️ รอตรวจ ({0 if empty else s['review_open']})", "🕘 ประวัติ"])

    # ------------------------------------------------------------------ import
    with t_imp:
        st.markdown("**1) เลือกไฟล์** — รองรับชีต `หัวลาก / หางพ่วง / ตู้เย็น` หรือ `M1_ซ่อมรถ / M2_ซ่อมหาง / M3_ซ่อมตู้` "
                    "(ใช้เฉพาะคอลัมน์: เลขที่เอกสาร, วันที่เปิดเอกสาร, วันที่แจ้งซ่อม, เวลา, ทะเบียน/เบอร์ตู้, รหัสพนักงาน, สถานะซ่อม)")
        up = st.file_uploader("ไฟล์ข้อมูลรถ (.xlsx)", type=["xlsx", "xlsm"], key="rr_up")
        go = st.button("2) ตรวจไฟล์ (ยังไม่บันทึก)", icon="🔍", disabled=up is None)
        if go and up is not None:
            data = up.getvalue()
            with st.spinner("กำลังอ่านและเทียบกับฐานข้อมูล…"):
                wb = rr.read_workbook(io.BytesIO(data))
                plan = _analyze(db, wb["records"]) if wb["records"] else None
            ss["rr_an"] = dict(name=up.name, sha=rr.file_sha256(data), wb=wb, plan=plan,
                               blank=sum(x["blank"] for x in wb["sheets"]), db=db)

        an = ss.get("rr_an")
        if an and an["db"] == db:
            wb, plan = an["wb"], an["plan"]
            for e in wb["errors"]:
                st.error(e)
            if wb["sheets"]:
                st.caption("ชีตที่อ่าน: " + " · ".join(f"{x['sheet']} → {x['asset_type']} ({x['rows']:,} แถว"
                                                        + (f", ว่าง {x['blank']:,}" if x['blank'] else "") + ")" for x in wb["sheets"]))
            if plan:
                m = st.columns(6)
                kpi_card(m[0], "ใหม่ (จะเพิ่ม)", f"{len(plan['new']):,}", "", "ok")
                kpi_card(m[1], "มีแล้ว ไม่เปลี่ยน", f"{len(plan['unchanged']):,}", "ข้าม")
                kpi_card(m[2], "อัปเดต", f"{len(plan['update']):,}", "สถานะ/เวลา/รหัสพนักงาน")
                kpi_card(m[3], "รอตรวจ", f"{len(plan['review']):,}", "ไม่เข้าตารางหลัก", "warn" if plan["review"] else "ok")
                kpi_card(m[4], "ซ้ำในไฟล์", f"{len(plan['dup_in_file']):,}", "ข้าม")
                kpi_card(m[5], "เคยส่งตรวจแล้ว", f"{len(plan['review_known']):,}", "ไม่ส่งซ้ำ")
                st.write("")
                new_df = _th(_df(plan["new"], SHOW))
                upd_df = pd.DataFrame([dict(เลขที่เอกสาร=u["rec"]["doc_no"], ฟิลด์=COLS_TH.get(f, f), ค่าเดิม=o, ค่าใหม่=n)
                                       for u in plan["update"] for f, (o, n) in u["changes"].items()])
                rev_df = pd.DataFrame([dict(เลขที่เอกสาร=r["rec"]["doc_no"], ประเภท=r["rec"]["asset_type"], ชนิด=r["kind"],
                                            เหตุผล=r["reason"], ทะเบียน=r["rec"]["asset_code"],
                                            วันที่แจ้งซ่อม=r["rec"]["report_date"], แถวต้นทาง=r["rec"]["source_row"])
                                       for r in plan["review"]])
                note_df = _th(_df([r for r in plan["new"] + [u["rec"] for u in plan["update"]] if r.get("data_note")],
                                  ["doc_no", "asset_type", "report_date", "report_time_raw", "report_time", "data_note", "source_row"]))
                kp = pd.DataFrame(plan["new"])
                kpi_new = (kp.groupby(["report_month", "asset_type"])["asset_code"].nunique().unstack(fill_value=0)
                           .reindex(columns=TYPES, fill_value=0).reset_index().rename(columns={"report_month": "เดือน"})
                           if len(kp) else pd.DataFrame())
                a1, a2, a3, a4, a5 = st.tabs([f"ใหม่ ({len(plan['new']):,})", f"อัปเดต ({len(plan['update']):,})",
                                              f"รอตรวจ ({len(plan['review']):,})", f"หมายเหตุข้อมูล ({len(note_df):,})",
                                              "คัน/เดือน (เฉพาะรายการใหม่)"])
                with a1:
                    st.dataframe(new_df, hide_index=True, height=360, **full)
                with a2:
                    st.dataframe(upd_df, hide_index=True, **full)
                with a3:
                    st.caption("รายการเหล่านี้จะไม่เข้าตารางหลัก — ไปอยู่แท็บ 'รอตรวจ' ให้ DATA_OWNER กดนำเข้า/ไม่นำเข้า")
                    st.dataframe(rev_df, hide_index=True, **full)
                with a4:
                    st.caption("นำเข้าได้ แต่ค่าเวลาเก็บตามต้นฉบับ (ไม่เดา) หรือแปลงจากรูปแบบวันที่ของ Excel")
                    st.dataframe(note_df, hide_index=True, **full)
                with a5:
                    st.dataframe(kpi_new, hide_index=True, **full)
                st.download_button("Export ผลการตรวจ (Excel)", _excel_bytes({"ใหม่": new_df, "อัปเดต": upd_df, "รอตรวจ": rev_df,
                                                                            "หมายเหตุข้อมูล": note_df, "คันต่อเดือน": kpi_new}),
                                   file_name=f"ตรวจไฟล์ใบแจ้งซ่อม_{datetime.now():%Y%m%d_%H%M}.xlsx", icon="⬇️")
                st.divider()
                n_act = len(plan["new"]) + len(plan["update"]) + len(plan["review"])
                label = "3) สร้างฐานข้อมูลหลัก (ชุดแรก)" if empty else "3) ยืนยันนำเข้า"
                ok = st.checkbox("ตรวจผลด้านบนแล้ว ยืนยันนำเข้า", key="rr_ok")
                if not is_local:
                    st.caption("โหมด Cloud: บันทึกลงฐานข้อมูลชั่วคราวของ session นี้เท่านั้น")
                if st.button(label, type="primary", icon="💾", disabled=not ok or n_act == 0 or bool(wb["errors"])):
                    res = rr.commit(db, wb["records"], an["name"], an["sha"], user, an["blank"])
                    ss.pop("rr_an", None)
                    ss["rr_last"] = res
                    st.rerun()
                if n_act == 0:
                    st.info("ไม่มีรายการใหม่/อัปเดต/รอตรวจ — ไฟล์นี้อยู่ในฐานข้อมูลครบแล้ว")
        last = ss.get("rr_last")
        if last:
            st.success(f"นำเข้าแล้ว (batch {last['batch_id']}{' · ฐานหลักชุดแรก' if last['is_base'] else ''}): "
                       f"เพิ่ม {last['rows_inserted']:,} · อัปเดต {last['rows_updated']:,} · รอตรวจ {last['rows_review']:,} · "
                       f"ไม่เปลี่ยน {last['rows_unchanged']:,} · ซ้ำในไฟล์ {last['rows_dup_in_file']:,}"
                       + (f"\n\nสำรองฐานข้อมูลก่อนนำเข้า: `{last['backup']}`" if last.get("backup") else ""))

    if empty:
        for t in (t_data, t_kpi, t_rev, t_log):
            with t:
                st.info("ยังไม่มีข้อมูลในฐาน — นำเข้าไฟล์ชุดแรกที่แท็บ 'นำเข้าไฟล์'")
        return

    # ------------------------------------------------------------------ data
    with t_data:
        with st.container(border=True):
            st.markdown("**Export สำหรับอัปเดตไฟล์หลัก** — ชีต `M1_ซ่อมรถ / M2_ซ่อมหาง / M3_ซ่อมตู้` รูปแบบเดียวกับไฟล์เดิม "
                        "(คอลัมน์ B–H ข้อมูล, I–L สูตรเดิม, คอลัมน์ M `ISO = Yes`) — ไม่รวมรายการรอตรวจ และไม่รวม `ISO = No`")
            if st.button("สร้างไฟล์ M1/M2/M3", icon="📤", key="rr_mexp"):
                import monthly_report as mr
                blob, cnt = rr.export_m_sheets(db)
                fn_ = f"อัปเดต_M1_M2_M3_{datetime.now():%Y%m%d_%H%M}.xlsx"
                sp_ = mr.save_local_export(app_dir, fn_, blob, is_local=is_local)
                ss["rr_mexp_blob"] = (fn_, blob, cnt, sp_)
            if ss.get("rr_mexp_blob"):
                _m = ss["rr_mexp_blob"]
                n, blob, cnt = _m[0], _m[1], _m[2]
                sp_ = _m[3] if len(_m) > 3 else None
                st.caption(" · ".join(f"{k} {v:,} แถว" for k, v in cnt.items()))
                if sp_:
                    st.success(f"💾 บันทึกไฟล์แล้วที่: `{sp_}`")
                st.download_button(f"ดาวน์โหลด {n}", blob, file_name=n, icon="⬇️", type="primary",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        c1, c2, c3, c4, c5 = st.columns([1.1, 1.1, 1.5, 1.1, 0.9])
        ft = c1.multiselect("ประเภท", TYPES, key="rr_ft", placeholder="ทุกประเภท")
        months = [r["m"] for r in rr.fetch(db, "SELECT DISTINCT report_month m FROM repair_request ORDER BY 1")]
        fm = c2.multiselect("เดือน", months, key="rr_fm", placeholder="ทุกเดือน")
        fq = c3.text_input("ค้นหา", key="rr_fq", placeholder="ทะเบียน / เลขเอกสาร / รหัสพนักงาน")
        fs = c4.multiselect("สถานะ", ["Approved", "Open", "Cancel"], key="rr_fs", placeholder="ทุกสถานะ")
        fiso = c5.selectbox("ใน ISO", ["ทั้งหมด", "Yes", "No"], key="rr_fiso")
        sql, args = "SELECT * FROM repair_request WHERE 1=1", []
        if ft:
            sql += f" AND asset_type IN ({','.join('?' * len(ft))})"; args += ft
        if fm:
            sql += f" AND report_month IN ({','.join('?' * len(fm))})"; args += fm
        if fs:
            sql += f" AND status IN ({','.join('?' * len(fs))})"; args += fs
        if fiso in ("Yes", "No"):
            sql += " AND in_iso = ?"; args.append(fiso)
        if fq:
            sql += " AND (doc_no LIKE ? OR asset_code LIKE ? OR IFNULL(employee_code,'') LIKE ?)"; args += [f"%{fq}%"] * 3
        rows = rr.fetch(db, sql + " ORDER BY report_date DESC, doc_no", *args)
        df = _th(_df(rows, SHOW + ["report_month", "report_time_raw", "flag_reason"]))
        st.caption(f"แสดง {len(df):,} รายการ")
        st.dataframe(df, hide_index=True, height=480, **full)
        df_exp = df[df["ใน ISO"] != "No"] if "ใน ISO" in df.columns else df
        st.download_button(f"Export ข้อมูลที่กรอง (Excel · เฉพาะ ISO = Yes · {len(df_exp):,} แถว)",
                           _excel_bytes({"ใบแจ้งซ่อม": df_exp}),
                           file_name=f"ใบแจ้งซ่อม_{datetime.now():%Y%m%d_%H%M}.xlsx", icon="⬇️")

    # ------------------------------------------------------------------ KPI
    with t_kpi:
        k = pd.DataFrame(rr.kpi_table(db, iso_only=True))
        _iso_no = rr.fetch(db, "SELECT COUNT(*) n, COUNT(DISTINCT asset_code) p FROM repair_request WHERE asset_type='หัวลาก' AND in_iso='No'")
        _no_n = _iso_no[0]["n"] if _iso_no else 0
        _no_p = _iso_no[0]["p"] if _iso_no else 0
        st.markdown("**จำนวนคัน/ตู้ที่มีใบแจ้งซ่อมในเดือน (คิดเฉพาะ `ISO = Yes`)** — นับทะเบียนไม่ซ้ำ ตามวันที่แจ้งซ่อม ทุกสถานะ")
        if _no_n:
            st.caption(f"ℹ️ หัวลาก `ISO = Yes` ตรงกับรายการอุปกรณ์ (Active) · ตัดหัวลาก `ISO = No` ออก {_no_n:,} ใบ ({_no_p} ทะเบียน)")
        if len(k):
            pv = k.pivot_table(index="report_month", columns="asset_type", values="vehicles", aggfunc="sum").reindex(columns=TYPES).fillna(0).astype(int)
            pr = k.pivot_table(index="report_month", columns="asset_type", values="requests", aggfunc="sum").reindex(columns=TYPES).fillna(0).astype(int)
            c1, c2 = st.columns(2)
            c1.caption("คัน/ตู้ (ไม่ซ้ำ)"); c1.dataframe(pv.rename_axis("เดือน").reset_index(), hide_index=True, **full)
            c2.caption("จำนวนใบแจ้งซ่อม"); c2.dataframe(pr.rename_axis("เดือน").reset_index(), hide_index=True, **full)
            st.bar_chart(pv, height=300, stack=False)
            import monthly_report_page as mrp
            mrp.render_iso_kpi_export_box(db, app_dir, full, key_prefix="rr_iso")
            with st.container(border=True):
                st.markdown("**Export KPI พร้อมรายละเอียด (Excel)** — ชีต คันต่อเดือน · ใบแจ้งซ่อมต่อเดือน + รายละเอียดการซ่อม **M1_ซ่อมรถ (เฉพาะ ISO = Yes) · M2_ซ่อมหาง · M3_ซ่อมตู้** แยกชีต")
                if st.button("สร้างไฟล์ Export KPI", icon="📊", key="rr_kpi_exp"):
                    import monthly_report as mr
                    with st.spinner("กำลังสร้างไฟล์…"):
                        blob, cnt = _kpi_export_bytes(db, pv, pr)
                    fn_ = f"KPI_ใบแจ้งซ่อม_{datetime.now():%Y%m%d_%H%M}.xlsx"
                    sp_ = mr.save_local_export(app_dir, fn_, blob, is_local=is_local)
                    ss["rr_kpi_blob"] = (fn_, blob, cnt, sp_)
                if ss.get("rr_kpi_blob"):
                    _k = ss["rr_kpi_blob"]
                    n, blob, cnt = _k[0], _k[1], _k[2]
                    sp_ = _k[3] if len(_k) > 3 else None
                    st.caption(" · ".join(f"{k_} {v:,} แถว" for k_, v in cnt.items()))
                    if sp_:
                        st.success(f"💾 บันทึกไฟล์แล้วที่: `{sp_}`")
                    st.download_button(f"ดาวน์โหลด {n}", blob, file_name=n, icon="⬇️", type="primary", key="rr_kpi_dl",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ------------------------------------------------------------------ review
    with t_rev:
        import json
        open_rows = rr.fetch(db, "SELECT * FROM repair_request_review WHERE decision IS NULL ORDER BY review_id")
        st.caption("รายการที่ไม่เข้าตารางหลักเพราะข้อมูลไม่สอดคล้อง — เลือก 'นำเข้า' (ใช้ค่าตามไฟล์) หรือ 'ไม่นำเข้า' ; บันทึกชื่อผู้ตัดสิน")
        if not open_rows:
            st.success("ไม่มีรายการรอตรวจ")
        for r in open_rows:
            data = json.loads(r["incoming_json"])
            variants = data if isinstance(data, list) else [data]
            with st.container(border=True):
                st.markdown(f"**{r['doc_no']}** · {r['kind']} · batch {r['batch_id']}  \n{r['reason']}")
                st.dataframe(_th(_df(variants, ["asset_type", "report_date", "report_time_raw", "asset_raw", "asset_code",
                                                "status", "employee_code", "source_sheet", "source_row"])), hide_index=True, **full)
                c1, c2, c3 = st.columns([2, 1, 1])
                idx = 0
                if len(variants) > 1:
                    idx = c1.selectbox("ใช้แถวที่", list(range(len(variants))), format_func=lambda i: f"แถวต้นทาง {variants[i]['source_row']}",
                                       key=f"rv_idx_{r['review_id']}")
                if c2.button("นำเข้า", key=f"rv_y_{r['review_id']}", type="primary", **full):
                    ok, msg = rr.decide_review(db, r["review_id"], "นำเข้า", user or ("local" if is_local else "user"), idx)
                    st.toast(msg); st.rerun()
                if c3.button("ไม่นำเข้า", key=f"rv_n_{r['review_id']}", **full):
                    ok, msg = rr.decide_review(db, r["review_id"], "ไม่นำเข้า", user or ("local" if is_local else "user"))
                    st.toast(msg); st.rerun()
        done = rr.fetch(db, "SELECT review_id, doc_no, kind, reason, decision, decided_by, decided_at FROM repair_request_review WHERE decision IS NOT NULL ORDER BY decided_at DESC")
        if done:
            with st.expander(f"ตัดสินแล้ว ({len(done)})"):
                st.dataframe(pd.DataFrame(done), hide_index=True, **full)

    # ------------------------------------------------------------------ log
    with t_log:
        b = pd.DataFrame(s["batches"])
        if len(b):
            b = b.rename(columns={"batch_id": "batch", "file_name": "ไฟล์", "imported_by": "ผู้นำเข้า", "imported_at": "เวลา",
                                  "is_base_load": "ฐานหลักชุดแรก", "rows_read": "อ่าน", "rows_blank": "ว่าง", "rows_dup_in_file": "ซ้ำในไฟล์",
                                  "rows_inserted": "เพิ่ม", "rows_unchanged": "ไม่เปลี่ยน", "rows_updated": "อัปเดต", "rows_review": "รอตรวจ",
                                  "backup_path": "ไฟล์สำรอง", "note": "หมายเหตุ"}).drop(columns=["file_sha256"], errors="ignore")
            st.markdown("**ประวัติการนำเข้า**")
            st.dataframe(b, hide_index=True, **full)
        h = rr.fetch(db, "SELECT doc_no, field, old_value, new_value, changed_by, changed_at, batch_id FROM repair_request_history ORDER BY hist_id DESC LIMIT 2000")
        st.markdown("**ประวัติการเปลี่ยนแปลงข้อมูล**")
        st.dataframe(pd.DataFrame(h), hide_index=True, **full)
        if is_local and os.path.exists(db):
            with open(db, "rb") as fh:
                st.download_button("ดาวน์โหลดไฟล์ฐานข้อมูล (.sqlite)", fh.read(), file_name=os.path.basename(db), icon="🗄️")


# ============================================================================= หน้าเลขไมล์
def render_mileage(is_local, user, full, app_dir, kpi_card):
    import json
    ss = st.session_state
    db = db_selector(is_local, app_dir, full)
    st.markdown('<div class="app-hero"><h1>เลขไมล์</h1><p>นำเข้าเลขไมล์รายคัน (ทะเบียน · วันที่ · กม.) — '
                'กันซ้ำด้วย ทะเบียน+วันที่ · ค่าผิดปกติรอ DATA_OWNER ตัดสิน</p></div>', unsafe_allow_html=True)
    stat = rr.mileage_fetch(db, "SELECT COUNT(*) n, COUNT(DISTINCT asset_code) a, MAX(reading_date) d FROM mileage_reading")[0]
    nrev = rr.mileage_fetch(db, "SELECT COUNT(*) n FROM mileage_review WHERE decision IS NULL")[0]["n"]
    k = st.columns(4)
    kpi_card(k[0], "บันทึกเลขไมล์", f"{stat['n']:,}", "รายการ")
    kpi_card(k[1], "จำนวนคัน", f"{stat['a']:,}", "")
    kpi_card(k[2], "วันที่ล่าสุด", stat["d"] or "-", "")
    kpi_card(k[3], "รอตรวจ", f"{nrev:,}", "", "warn" if nrev else "ok")
    st.write("")
    t1, t2, t3, t4 = st.tabs(["📥 นำเข้าไฟล์", "📏 เลขไมล์ล่าสุด", f"⚠️ รอตรวจ ({nrev})", "🕘 ประวัติ"])
    with t1:
        up = st.file_uploader("ไฟล์เลขไมล์ (.xlsx) — คอลัมน์ Registration/ทะเบียน, Date/วันที่, km./เลขไมล์", type=["xlsx", "xlsm"], key="mi_up")
        if st.button("ตรวจไฟล์ (ยังไม่บันทึก)", icon="🔍", disabled=up is None, key="mi_chk") and up is not None:
            data = up.getvalue()
            wb = rr.read_mileage_workbook(io.BytesIO(data))
            con = rr.connect_mileage(db) if os.path.exists(db) else None
            if con is None:
                con = sqlite3.connect(":memory:"); con.row_factory = sqlite3.Row
                con.executescript(rr.SCHEMA_SQL); con.executescript(rr.MILEAGE_SQL)
            plan = rr.analyze_mileage(con, wb["records"]) if wb["records"] else None
            con.close()
            ss["mi_an"] = dict(name=up.name, sha=rr.file_sha256(data), wb=wb, plan=plan, db=db)
        an = ss.get("mi_an")
        if an and an["db"] == db:
            for e in an["wb"]["errors"]:
                st.error(e)
            st.caption("ชีตที่อ่าน: " + " · ".join(f"{x['sheet']} ({x['rows']} แถว{'' if x['header'] else ', ไม่มีหัวตาราง'})" for x in an["wb"]["sheets"]))
            plan = an["plan"]
            if plan:
                m = st.columns(5)
                kpi_card(m[0], "ใหม่", f"{len(plan['new']):,}", "", "ok")
                kpi_card(m[1], "มีแล้ว ไม่เปลี่ยน", f"{len(plan['unchanged']):,}", "")
                kpi_card(m[2], "รอตรวจ", f"{len(plan['review']):,}", "", "warn" if plan["review"] else "ok")
                kpi_card(m[3], "ซ้ำในไฟล์", f"{len(plan['dup_in_file']):,}", "ข้าม")
                kpi_card(m[4], "เคยส่งตรวจแล้ว", f"{len(plan['review_known']):,}", "")
                st.write("")
                a1, a2 = st.tabs([f"ใหม่ ({len(plan['new'])})", f"รอตรวจ ({len(plan['review'])})"])
                with a1:
                    st.dataframe(pd.DataFrame([dict(ทะเบียน=r["asset_code"], วันที่=r["reading_date"], กม=r["km"], หมายเหตุ=r["data_note"] or "",
                                                    ชีต=r["source_sheet"], แถว=r["source_row"]) for r in plan["new"]]), hide_index=True, **full)
                with a2:
                    st.dataframe(pd.DataFrame([dict(ทะเบียน=x["rec"]["asset_code"], วันที่=x["rec"]["reading_date"], ค่าในไฟล์=x["rec"]["km_raw"],
                                                    เหตุผล=x["reason"]) for x in plan["review"]]), hide_index=True, **full)
                ok = st.checkbox("ตรวจผลแล้ว ยืนยันนำเข้า", key="mi_ok")
                n_act = len(plan["new"]) + len(plan["review"])
                if st.button("ยืนยันนำเข้าเลขไมล์", type="primary", icon="💾", disabled=not ok or n_act == 0, key="mi_go"):
                    ss["mi_last"] = rr.commit_mileage(db, an["wb"]["records"], an["name"], an["sha"], user)
                    ss.pop("mi_an", None)
                    st.rerun()
                if n_act == 0:
                    st.info("ไฟล์นี้อยู่ในฐานข้อมูลครบแล้ว")
        if ss.get("mi_last"):
            l = ss["mi_last"]
            st.success(f"นำเข้าแล้ว (batch {l['batch_id']}): เพิ่ม {l['rows_inserted']:,} · รอตรวจ {l['rows_review']:,} · "
                       f"ไม่เปลี่ยน {l['rows_unchanged']:,} · ซ้ำในไฟล์ {l['rows_dup_in_file']:,}" + (f"\n\nสำรอง: `{l['backup']}`" if l.get("backup") else ""))
    with t2:
        df = pd.DataFrame(rr.mileage_fetch(db, "SELECT asset_code ทะเบียน, reading_date วันที่, km กม, data_note หมายเหตุ FROM v_mileage_latest ORDER BY asset_code"))
        q = st.text_input("ค้นหาทะเบียน", key="mi_q")
        if q and len(df):
            df = df[df["ทะเบียน"].str.contains(q, regex=False)]
        if len(df):
            df["วันที่"] = pd.to_datetime(df["วันที่"]).dt.strftime("%d/%m/%Y")
            df = df.astype(object).where(df.notna(), "")
        st.dataframe(df, hide_index=True, height=480, column_config={"กม": st.column_config.NumberColumn(format="%,d")}, **full)
        if len(df):
            st.download_button("Export เลขไมล์ล่าสุด (Excel)", _excel_bytes({"เลขไมล์ล่าสุด": df}),
                               file_name=f"เลขไมล์ล่าสุด_{datetime.now():%Y%m%d_%H%M}.xlsx", icon="⬇️")
    with t3:
        rows = rr.mileage_fetch(db, "SELECT * FROM mileage_review WHERE decision IS NULL ORDER BY review_id")
        st.caption("ค่าที่ใช้ไม่ได้/ขัดกับข้อมูลเดิม — กรอกเลขไมล์ที่ถูกต้องจากเอกสารจริงแล้วกดนำเข้า หรือกดไม่นำเข้า (ระบบไม่เดาค่าแทน)")
        if not rows:
            st.success("ไม่มีรายการรอตรวจ")
        for r in rows:
            with st.container(border=True):
                st.markdown(f"**{r['asset_code']}** · วันที่ {r['reading_date']} · ค่าในไฟล์ `{r['km_raw']}`  \n{r['reason']}")
                c1, c2, c3 = st.columns([2, 1, 1])
                kmv = c1.number_input("เลขไมล์ที่ถูกต้อง (ถ้าต้องการแก้)", min_value=0, step=1, value=None, key=f"mi_km_{r['review_id']}")
                if c2.button("นำเข้า", key=f"mi_y_{r['review_id']}", type="primary", **full):
                    ok, msg = rr.decide_mileage_review(db, r["review_id"], "นำเข้า", user or ("local" if is_local else "user"),
                                                       km_override=int(kmv) if kmv is not None else None)
                    (st.toast if ok else st.error)(msg)
                    if ok:
                        st.rerun()
                if c3.button("ไม่นำเข้า", key=f"mi_n_{r['review_id']}", **full):
                    rr.decide_mileage_review(db, r["review_id"], "ไม่นำเข้า", user or ("local" if is_local else "user"))
                    st.rerun()
    with t4:
        st.dataframe(pd.DataFrame(rr.mileage_fetch(db, "SELECT batch_id, file_name ไฟล์, imported_by ผู้นำเข้า, imported_at เวลา, rows_read อ่าน, rows_dup_in_file ซ้ำในไฟล์, rows_inserted เพิ่ม, rows_unchanged ไม่เปลี่ยน, rows_review รอตรวจ, backup_path ไฟล์สำรอง FROM mileage_batch ORDER BY batch_id DESC")), hide_index=True, **full)
        q2 = st.text_input("ดูประวัติเลขไมล์ทะเบียน", key="mi_h")
        if q2:
            st.dataframe(pd.DataFrame(rr.mileage_fetch(db, "SELECT reading_date วันที่, km กม, data_note หมายเหตุ, source_file ไฟล์, approved_by ผู้อนุมัติ FROM mileage_reading WHERE asset_code=? ORDER BY reading_date", q2.strip())), hide_index=True, **full)
