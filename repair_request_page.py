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
           "status": "สถานะซ่อม", "report_month": "เดือน", "source_sheet": "ชีตต้นทาง", "source_row": "แถวต้นทาง",
           "data_note": "หมายเหตุข้อมูล", "flag_reason": "เหตุผล", "confirm_flag": "Confirm Flag"}
SHOW = ["doc_no", "asset_type", "report_date", "report_time", "asset_code", "brand", "status", "employee_code",
        "open_date", "data_note", "source_sheet", "source_row"]


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


def render(is_local, user, full, app_dir, kpi_card):
    ss = st.session_state
    default_db = os.path.join(app_dir, "data", "Repair_Request_v2.sqlite")
    if is_local:
        ss.setdefault("rr_db", default_db)
    else:
        if "rr_db_tmp" not in ss:
            ss["rr_db_tmp"] = os.path.join(tempfile.mkdtemp(prefix="rrdb_"), "Repair_Request_v2.sqlite")
        ss["rr_db"] = ss["rr_db_tmp"]

    with st.sidebar:
        st.markdown("**ฐานข้อมูลใบแจ้งซ่อม**")
        if is_local:
            st.text_input("ไฟล์ฐานข้อมูล (.sqlite)", key="rr_db")
            st.caption("สำรองไฟล์อัตโนมัติไว้ที่โฟลเดอร์ backup ข้างไฟล์ฐานข้อมูล ก่อนนำเข้าทุกครั้ง")
        else:
            st.warning("โหมด Cloud: ฐานข้อมูลเป็นแบบชั่วคราวเฉพาะ session นี้ (หายเมื่อปิด/รีสตาร์ท) "
                       "ใช้สำหรับตรวจไฟล์และ Export — บันทึกฐานข้อมูลหลักให้เปิดแอปบนเครื่อง")
    db = ss["rr_db"]

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
        c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.6, 1.2])
        ft = c1.multiselect("ประเภท", TYPES, key="rr_ft", placeholder="ทุกประเภท")
        months = [r["m"] for r in rr.fetch(db, "SELECT DISTINCT report_month m FROM repair_request ORDER BY 1")]
        fm = c2.multiselect("เดือน", months, key="rr_fm", placeholder="ทุกเดือน")
        fq = c3.text_input("ค้นหา", key="rr_fq", placeholder="ทะเบียน / เลขเอกสาร / รหัสพนักงาน")
        fs = c4.multiselect("สถานะ", ["Approved", "Open", "Cancel"], key="rr_fs", placeholder="ทุกสถานะ")
        sql, args = "SELECT * FROM repair_request WHERE 1=1", []
        if ft:
            sql += f" AND asset_type IN ({','.join('?' * len(ft))})"; args += ft
        if fm:
            sql += f" AND report_month IN ({','.join('?' * len(fm))})"; args += fm
        if fs:
            sql += f" AND status IN ({','.join('?' * len(fs))})"; args += fs
        if fq:
            sql += " AND (doc_no LIKE ? OR asset_code LIKE ? OR IFNULL(employee_code,'') LIKE ?)"; args += [f"%{fq}%"] * 3
        rows = rr.fetch(db, sql + " ORDER BY report_date DESC, doc_no", *args)
        df = _th(_df(rows, SHOW + ["report_month", "report_time_raw", "flag_reason"]))
        st.caption(f"แสดง {len(df):,} รายการ")
        st.dataframe(df, hide_index=True, height=480, **full)
        st.download_button("Export ข้อมูลที่กรอง (Excel)", _excel_bytes({"ใบแจ้งซ่อม": df}),
                           file_name=f"ใบแจ้งซ่อม_{datetime.now():%Y%m%d_%H%M}.xlsx", icon="⬇️")

    # ------------------------------------------------------------------ KPI
    with t_kpi:
        k = pd.DataFrame(rr.kpi_table(db))
        st.markdown("**จำนวนคัน/ตู้ที่มีใบแจ้งซ่อมในเดือน** (นับทะเบียนไม่ซ้ำ ตามวันที่แจ้งซ่อม ทุกสถานะ — เทียบเท่าคอลัมน์ Index ในไฟล์ KPI)")
        if len(k):
            pv = k.pivot_table(index="report_month", columns="asset_type", values="vehicles", aggfunc="sum").reindex(columns=TYPES).fillna(0).astype(int)
            pr = k.pivot_table(index="report_month", columns="asset_type", values="requests", aggfunc="sum").reindex(columns=TYPES).fillna(0).astype(int)
            c1, c2 = st.columns(2)
            c1.caption("คัน/ตู้ (ไม่ซ้ำ)"); c1.dataframe(pv.rename_axis("เดือน").reset_index(), hide_index=True, **full)
            c2.caption("จำนวนใบแจ้งซ่อม"); c2.dataframe(pr.rename_axis("เดือน").reset_index(), hide_index=True, **full)
            st.bar_chart(pv, height=300, stack=False)
            st.download_button("Export KPI (Excel)", _excel_bytes({"คันต่อเดือน": pv.reset_index(), "ใบแจ้งซ่อมต่อเดือน": pr.reset_index()}),
                               file_name=f"KPI_ใบแจ้งซ่อม_{datetime.now():%Y%m%d_%H%M}.xlsx", icon="⬇️")

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
