# -*- coding: utf-8 -*-
"""หน้า 'พิมพ์รายงานประจำเดือน' (KPI ประจำเดือน + 1_หัวลาก / 2_หางพ่วง / 3_ซ่อมตู้ + ทะเบียนรถหลัก Master)"""
import io
import os
from datetime import datetime

import pandas as pd
import streamlit as st

import monthly_report as mr
import repair_request_page as rrp

TYPES = mr.TYPES


def _render_html_preview(html_str, height=760):
    if hasattr(st, "iframe"):
        st.iframe(html_str, height=height)
    else:
        import streamlit.components.v1 as components
        components.html(html_str, height=height, scrolling=True)


def render_iso_kpi_export_box(db, app_dir, full, key_prefix="iso", default_year=None, is_local=None):
    """กล่อง Export ลง Template KPI ต้นฉบับ ISO (FP-MR-00-01 Rev.00: MTD-หัวลาก / MTD-หางพ่วง / MTD-ตู้เย็น + ทะเบียนทั้งหมด)
    ใช้ร่วมกันทั้งในหน้า 'พิมพ์รายงานประจำเดือน' และหน้าเดิม 'ฐานข้อมูลใบแจ้งซ่อม → KPI คัน/เดือน'"""
    ss = st.session_state
    if is_local is None:
        is_local = (os.name == "nt" or os.environ.get("VR_LOCAL") == "1")
    months = mr.available_months(db)
    years = sorted({m[:4] for m in months}, reverse=True) or [str(datetime.now().year)]
    tpl_path = mr.find_kpi_template_path(app_dir)

    with st.container(border=True):
        st.markdown(
            "**📋 Export ลง Template KPI บริษัท (`FP-MR-00-01 Rev.00`)** — "
            "คงรูปแบบต้นฉบับครบ 3 ชีต (`MTD-หัวลาก`, `MTD-หางพ่วง`, `MTD-ตู้เย็น`) พร้อมกราฟและสูตร + ชีต `ทะเบียนทั้งหมด` (ล้าง Pivot Cache บวมแล้ว)"
        )
        c1, c2, c3, c4 = st.columns([1.0, 2.2, 1.0, 1.3])
        idx_y = years.index(default_year) if default_year in years else 0
        sel_year = c1.selectbox("ปี (ค.ศ.)", years, index=idx_y, key=f"{key_prefix}_y")
        src_mode = c2.radio(
            "แหล่งตัวเลขแถว 11 (รายการที่บำรุงรักษาได้จริง)",
            ["เฉพาะทะเบียนใน Master (แนะนำ)", "ทั้งหมดในฐานข้อมูล DB"],
            horizontal=True,
            key=f"{key_prefix}_src",
        )
        target_pct = c3.number_input("Target (%)", min_value=1, max_value=100, value=80, step=1, key=f"{key_prefix}_tgt")
        inc_matrix = c4.checkbox("รวมชีต ทะเบียนทั้งหมด", value=True, key=f"{key_prefix}_mat")

        with st.expander("⚙️ ตั้งค่าแผนบำรุงรักษาตาม Job งาน (แถว 12) / ชื่อผู้อนุมัติ / อัปโหลด Template อื่น", expanded=False):
            m_cols = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
            plan_df_init = pd.DataFrame(
                [dict(ประเภท=at, **{m_cols[i]: mr.DEFAULT_KPI_PLANS[at][i] for i in range(12)}) for at in TYPES]
            )
            edited_plans = st.data_editor(
                plan_df_init,
                hide_index=True,
                disabled=["ประเภท"],
                key=f"{key_prefix}_plans",
                **full,
            )
            s1, s2, s3, s4 = st.columns(4)
            issued_by = s1.text_input("Issued by", value="", placeholder="คงตาม Template", key=f"{key_prefix}_iss")
            reviewed_by = s2.text_input("Reviewed by", value="", placeholder="คงตาม Template", key=f"{key_prefix}_rev")
            approved_mgr = s3.text_input("Approved (Dept. Mgr.)", value="", placeholder="คงตาม Template", key=f"{key_prefix}_mgr")
            approved_qmr = s4.text_input("Approved (QMR)", value="", placeholder="คงตาม Template", key=f"{key_prefix}_qmr")

            up_tpl = st.file_uploader(
                "อัปโหลดไฟล์ Template KPI บริษัท (.xlsx) — ไม่ต้องอัปโหลดหากใช้ Template มาตรฐานที่ติดตั้งไว้แล้ว",
                type=["xlsx", "xlsm"],
                key=f"{key_prefix}_up_tpl",
            )
            if tpl_path:
                st.caption(f"Template มาตรฐานพร้อมใช้งาน: `{os.path.basename(tpl_path)}`")
            elif up_tpl is None:
                st.warning("ไม่พบไฟล์ Template KPI มาตรฐานในโฟลเดอร์ data/ — กรุณาอัปโหลดไฟล์ Template KPI ด้านบน (โหมด Cloud)")

        tpl_src = up_tpl.getvalue() if up_tpl is not None else tpl_path
        btn_disabled = tpl_src is None
        if st.button("📊 สร้างไฟล์ Template KPI (FP-MR-00-01 Rev.00)", type="primary", disabled=btn_disabled, key=f"{key_prefix}_go"):
            plans_dict = {}
            for _, row in edited_plans.iterrows():
                at_k = row["ประเภท"]
                plans_dict[at_k] = [int(row[mc]) if pd.notna(row[mc]) else 0 for mc in m_cols]
            meta_dict = {
                k: v.strip()
                for k, v in dict(
                    issued_by=issued_by,
                    reviewed_by=reviewed_by,
                    approved_mgr=approved_mgr,
                    approved_qmr=approved_qmr,
                ).items()
                if v and v.strip()
            }
            with st.spinner("กำลังเติมข้อมูลลง Template KPI (FP-MR-00-01 Rev.00)…"):
                blob, smry = mr.export_kpi_iso_template(
                    db,
                    year=sel_year,
                    template_source=tpl_src,
                    use_master_only=src_mode.startswith("เฉพาะ"),
                    plans_by_type=plans_dict,
                    target_pct=int(target_pct),
                    meta=meta_dict,
                    include_plate_matrix=inc_matrix,
                )
            fn_iso = f"KPI-MTD-{sel_year}_{datetime.now():%Y%m%d_%H%M}.xlsx"
            saved_p = mr.save_local_export(app_dir, fn_iso, blob, is_local=is_local)
            ss[f"{key_prefix}_blob"] = (fn_iso, blob, smry, saved_p)

        if ss.get(f"{key_prefix}_blob"):
            _b = ss[f"{key_prefix}_blob"]
            fn, blob, smry = _b[0], _b[1], _b[2]
            saved_p = _b[3] if len(_b) > 3 else None
            st.caption(
                " · ".join(
                    f"**{at}**: {v['months']} เดือน, จริง {v['actual']:,} / แผน {v['plan']:,} ({v['achieve_pct']:.1f}%)"
                    for at, v in smry.items()
                )
            )
            if saved_p:
                st.success(f"💾 บันทึกไฟล์แล้วที่: `{saved_p}`")
            st.download_button(
                f"⬇️ ดาวน์โหลด {fn}",
                blob,
                file_name=fn,
                type="primary",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"{key_prefix}_dl",
            )


def render_monthly(is_local, user, full, app_dir, kpi_card):
    ss = st.session_state
    db = rrp.db_selector(is_local, app_dir, full)
    logo_path = mr.find_logo_path(app_dir)

    st.markdown(
        '<div class="app-hero"><h1>🖨️ พิมพ์รายงาน KPI และรายงานซ่อมประจำเดือน</h1>'
        '<p>สรุป KPI ประจำเดือน · แบบฟอร์มตาม Template บริษัท (1_หัวลาก · 2_หางพ่วง · 3_ซ่อมตู้) · '
        'รองรับทั้ง Excel พร้อมพิมพ์ (.xlsx) และหน้าพิมพ์ / Save PDF (.html) · แยกรถนอก Master ให้ DATA_OWNER ตรวจสอบ</p></div>',
        unsafe_allow_html=True,
    )

    if not os.path.exists(db):
        st.warning("ยังไม่มีไฟล์ฐานข้อมูล — กรุณานำเข้าใบแจ้งซ่อมที่หน้า 'ฐานข้อมูลใบแจ้งซ่อม' หรืออัปโหลดไฟล์ .sqlite ที่แถบซ้ายก่อน")
        return

    months = mr.available_months(db)
    if not months:
        st.info("ยังไม่มีข้อมูลใบแจ้งซ่อมในฐานข้อมูล")
        return

    master_sum = mr.get_master_summary(db)
    total_master_all = sum(v["active"] for v in master_sum.values())

    c_mo, c_info = st.columns([1.4, 3.6])
    sel_month = c_mo.selectbox(
        "เลือกเดือนที่ต้องการพิมพ์รายงาน",
        months,
        format_func=mr.format_month_th,
        key="mr_month",
    )
    if total_master_all == 0:
        c_info.warning("⚠️ ยังไม่มีรายชื่อทะเบียนรถหลัก (Master) ในฐานข้อมูล — ไปที่แท็บ '🚛 ทะเบียนรถหลัก (Master)' เพื่อนำเข้าจากไฟล์บริษัท")
    else:
        c_info.caption(
            f"ทะเบียนรถหลัก (Active Master): หัวลาก **{master_sum['หัวลาก']['active']:,}** · "
            f"หางพ่วง **{master_sum['หางพ่วง']['active']:,}** · "
            f"ตู้เย็น **{master_sum['ตู้เย็น']['active']:,}** คัน/ตู้"
            + (" · โลโก้บริษัทพร้อมใช้งาน" if logo_path else " · ไม่พบไฟล์โลโก้ (พิมพ์แบบไม่มีโลโก้)")
        )

    rep = mr.compute_monthly_report(db, sel_month)

    # Top KPI cards
    kcols = st.columns(4)
    total_unm = 0
    for i, at in enumerate(TYPES):
        k = rep["by_type"][at]["kpi"]
        total_unm += k["unmatched_assets"]
        sub = f"เข้าซ่อม {k['repaired_in_master']:,} / {k['total_master']:,} คัน" + (
            f" · นอก Master {k['unmatched_assets']}" if k["unmatched_assets"] else ""
        )
        kpi_card(kcols[i], f"{at} ({sel_month})", f"{k['pct_repaired']*100:.1f}%", sub)
    kpi_card(
        kcols[3],
        "นอก Master (รอคุณแขกตรวจ)",
        f"{total_unm:,}",
        "รถ/ตู้ที่มีใบซ่อมแต่ไม่อยู่ใน Master",
        "warn" if total_unm else "ok",
    )
    st.write("")

    t_pkg, t_kpi, t_forms = st.tabs([
        "📋 พิมพ์ KPI ประจำเดือน",
        "📊 สรุป KPI และหน้าพิมพ์ PDF",
        f"🖨️ รายงานซ่อม (1_หัวลาก / 2_หางพ่วง / 3_ซ่อมตู้){f' · ⚠️ นอก Master {total_unm}' if total_unm else ''}",
    ])

    # ========================================================================= TAB 1: พิมพ์ KPI ประจำเดือน (Export ตาม Template 13 ชีต)
    with t_pkg:
        tpl_path = mr.find_kpi_template_path(app_dir)
        st.markdown(
            "**📋 พิมพ์ KPI ประจำเดือน (Export ตาม Template บริษัทครบ 13 ชีต)**  \n"
            "ประกอบด้วย: `Index` (สารบัญและคำอธิบายทุกชีต) · `MTD-หัวลาก` · `MTD-หางพ่วง` · `MTD-ตู้เย็น` · `ทะเบียนทั้งหมด` · "
            "`สรุป_KPI_ประจำเดือน` · `1_หัวลาก` · `2_หางพ่วง` · `3_ซ่อมตู้` · `M1_ซ่อมรถ` · `M2_ซ่อมหาง` · `M3_ซ่อมตู้` · `Last PM` "
            "(ตัด `ISO = No` ออกทั้งหมดอัตโนมัติ และเติมข้อมูลรายเดือนถึงเดือนที่เลือก)"
        )

        with st.container(border=True):
            pc1, pc2, pc3 = st.columns([1.4, 2.2, 1.0])
            pkg_month = pc1.selectbox(
                "เลือกเดือนที่ต้องการพิมพ์รายงาน",
                months,
                index=months.index(sel_month) if sel_month in months else 0,
                format_func=mr.format_month_th,
                key="pkg_month",
            )
            pkg_src_mode = pc2.radio(
                "แหล่งตัวเลขแถว 11 (รายการที่บำรุงรักษาได้จริง)",
                ["เฉพาะทะเบียนใน Master (แนะนำ)", "ทั้งหมดในฐานข้อมูล DB"],
                horizontal=True,
                key="pkg_src",
            )
            pkg_target = pc3.number_input("Target (%)", min_value=1, max_value=100, value=80, step=1, key="pkg_tgt")

            with st.expander("⚙️ ตั้งค่าแผนบำรุงรักษาตาม Job งาน (แถว 12) / ชื่อผู้อนุมัติ / อัปโหลด Template อื่น", expanded=False):
                m_cols = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
                pkg_plan_init = pd.DataFrame(
                    [dict(ประเภท=at, **{m_cols[i]: mr.DEFAULT_KPI_PLANS[at][i] for i in range(12)}) for at in TYPES]
                )
                pkg_edited_plans = st.data_editor(
                    pkg_plan_init,
                    hide_index=True,
                    disabled=["ประเภท"],
                    key="pkg_plans",
                    **full,
                )
                ps1, ps2, ps3, ps4 = st.columns(4)
                pkg_iss = ps1.text_input("Issued by", value="", placeholder="คงตาม Template", key="pkg_iss")
                pkg_rev = ps2.text_input("Reviewed by", value="", placeholder="คงตาม Template", key="pkg_rev")
                pkg_mgr = ps3.text_input("Approved (Dept. Mgr.)", value="", placeholder="คงตาม Template", key="pkg_mgr")
                pkg_qmr = ps4.text_input("Approved (QMR)", value="", placeholder="คงตาม Template", key="pkg_qmr")

                pkg_up_tpl = st.file_uploader(
                    "อัปโหลดไฟล์ Template KPI บริษัท (.xlsx) — ไม่ต้องอัปโหลดหากใช้ Template มาตรฐานที่ติดตั้งไว้แล้ว",
                    type=["xlsx", "xlsm"],
                    key="pkg_up_tpl",
                )
                if tpl_path:
                    st.caption(f"Template มาตรฐานพร้อมใช้งาน: `{os.path.basename(tpl_path)}`")
                elif pkg_up_tpl is None:
                    st.warning("ไม่พบไฟล์ Template KPI มาตรฐานในโฟลเดอร์ data/ — กรุณาอัปโหลดไฟล์ Template KPI ด้านบน")

            pkg_tpl_src = pkg_up_tpl.getvalue() if pkg_up_tpl is not None else tpl_path
            if st.button(
                f"📥 Export KPI ประจำเดือน {mr.format_month_th(pkg_month)} (ตาม Template 13 ชีต)",
                type="primary",
                disabled=(pkg_tpl_src is None),
                key="pkg_go",
                **full,
            ):
                pkg_plans_dict = {}
                for _, row in pkg_edited_plans.iterrows():
                    at_k = row["ประเภท"]
                    pkg_plans_dict[at_k] = [int(row[mc]) if pd.notna(row[mc]) else 0 for mc in m_cols]
                pkg_meta = {
                    k: v.strip()
                    for k, v in dict(
                        issued_by=pkg_iss,
                        reviewed_by=pkg_rev,
                        approved_mgr=pkg_mgr,
                        approved_qmr=pkg_qmr,
                    ).items()
                    if v and v.strip()
                }
                with st.spinner(f"กำลังสร้างไฟล์ KPI ประจำเดือน {pkg_month} ครบ 13 ชีตตาม Template…"):
                    pkg_blob, pkg_info = mr.build_kpi_monthly_package(
                        db,
                        report_month=pkg_month,
                        template_source=pkg_tpl_src,
                        logo_path=logo_path,
                        use_master_only=pkg_src_mode.startswith("เฉพาะ"),
                        plans_by_type=pkg_plans_dict,
                        target_pct=int(pkg_target),
                        meta=pkg_meta,
                    )
                fn_pkg = f"KPI-MTD_{pkg_month}_{datetime.now():%Y%m%d_%H%M}.xlsx"
                sp_pkg = mr.save_local_export(app_dir, fn_pkg, pkg_blob, is_local=is_local)
                ss["mr_pkg_blob"] = (fn_pkg, pkg_blob, pkg_info, sp_pkg)

            if ss.get("mr_pkg_blob"):
                fn_p, blob_p, info_p, sp_p = ss["mr_pkg_blob"]
                st.caption(
                    f"**เดือน:** {info_p['month_th']} ({info_p['report_month']}) · "
                    f"**รวม {len(info_p['sheets'])} ชีต:** {', '.join(info_p['sheets'])}"
                )
                st.caption(
                    " · ".join(
                        f"**{at}**: {v['months']} เดือน, จริง {v['actual']:,} / แผน {v['plan']:,} ({v['achieve_pct']:.1f}%)"
                        for at, v in info_p["iso_summary"].items()
                    )
                    + "  |  "
                    + " · ".join(f"**{k}**: {v:,} แถว" for k, v in info_p["m_counts"].items())
                )
                if sp_p:
                    st.success(f"💾 บันทึกไฟล์แล้วที่: `{sp_p}`")
                st.download_button(
                    f"⬇️ ดาวน์โหลด {fn_p}",
                    blob_p,
                    file_name=fn_p,
                    type="primary",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="pkg_dl",
                )

        # ตารางสรุป KPI ของเดือนที่เลือกในแถบนี้
        pkg_rep = rep if pkg_month == sel_month else mr.compute_monthly_report(db, pkg_month)
        st.markdown(f"**ตัวอย่างตัวเลข KPI ประจำเดือน {pkg_rep['month_th']}** ({pkg_rep['date_range_text']})")
        pkg_rows = []
        for at in TYPES:
            k = pkg_rep["by_type"][at]["kpi"]
            pkg_rows.append({
                "ประเภท": at,
                "Total (ใน Master)": k["total_master"],
                "เข้าซ่อม (ใน Master)": k["repaired_in_master"],
                "% การเข้าซ่อม": round(k["pct_repaired"] * 100, 1),
                "ใบแจ้งซ่อม (ใน Master)": k["docs_in_master"],
                "นอก Master (รอตรวจ)": k["unmatched_assets"],
            })
        st.dataframe(
            pd.DataFrame(pkg_rows),
            hide_index=True,
            column_config={"% การเข้าซ่อม": st.column_config.NumberColumn(format="%.1f%%")},
            **full,
        )

    # ========================================================================= TAB 2: สรุป KPI และหน้าพิมพ์ PDF
    with t_kpi:
        st.markdown(f"**สรุป KPI ประจำเดือน {rep['month_th']}** ({rep['date_range_text']})")
        kpi_rows = []
        for at in TYPES:
            k = rep["by_type"][at]["kpi"]
            kpi_rows.append({
                "ประเภท": at,
                "Total (ใน Master)": k["total_master"],
                "เข้าซ่อม (ใน Master)": k["repaired_in_master"],
                "% การเข้าซ่อม": round(k["pct_repaired"] * 100, 1),
                "ใบแจ้งซ่อม (ใน Master)": k["docs_in_master"],
                "นอก Master (รอตรวจ)": k["unmatched_assets"],
                "รวมในฐาน DB (คัน/ตู้)": k["total_db_assets"],
                "รวมใบแจ้งซ่อม (ใบ)": k["total_db_docs"],
            })
        st.dataframe(
            pd.DataFrame(kpi_rows),
            hide_index=True,
            column_config={"% การเข้าซ่อม": st.column_config.NumberColumn(format="%.1f%%")},
            **full,
        )

        yearly = mr.compute_yearly_kpi(db, year=sel_month[:4])
        if yearly:
            st.markdown(f"**เปรียบเทียบ KPI รายเดือน (ปี {sel_month[:4]})**")
            ydf = pd.DataFrame([
                {
                    "เดือน": y["report_month"],
                    "หัวลาก (เข้าซ่อม)": y["หัวลาก_in_master"],
                    "หัวลาก (%)": round(y["หัวลาก_pct"] * 100, 1),
                    "หัวลาก (นอก Master)": y["หัวลาก_unmatched"],
                    "หางพ่วง (เข้าซ่อม)": y["หางพ่วง_in_master"],
                    "หางพ่วง (%)": round(y["หางพ่วง_pct"] * 100, 1),
                    "หางพ่วง (นอก Master)": y["หางพ่วง_unmatched"],
                    "ตู้เย็น (เข้าซ่อม)": y["ตู้เย็น_in_master"],
                    "ตู้เย็น (%)": round(y["ตู้เย็น_pct"] * 100, 1),
                    "ตู้เย็น (นอก Master)": y["ตู้เย็น_unmatched"],
                }
                for y in yearly
            ])
            st.dataframe(ydf, hide_index=True, **full)

        # ---- กล่อง Export ลง Template KPI ต้นฉบับ ISO (FP-MR-00-01 Rev.00)
        render_iso_kpi_export_box(db, app_dir, full, key_prefix="mr_iso", default_year=sel_month[:4])

        with st.container(border=True):
            st.markdown("**สั่งพิมพ์ / ดาวน์โหลดรายงานสรุป KPI + แบบฟอร์มซ่อมประจำเดือน**")
            b1, b2, b3 = st.columns(3)
            if b1.button("📦 สร้างชุดรายงานเต็ม (KPI + 3 แบบฟอร์ม Excel)", type="primary", key="mr_btn_all_xl", **full):
                with st.spinner("กำลังสร้างไฟล์ Excel พร้อมพิมพ์…"):
                    blob, _ = mr.build_monthly_excel(db, sel_month, logo_path=logo_path, include_kpi=True, include_types=TYPES)
                fn_ = f"รายงานซ่อมและKPI_{sel_month}_{datetime.now():%Y%m%d_%H%M}.xlsx"
                sp_ = mr.save_local_export(app_dir, fn_, blob, is_local=is_local)
                ss["mr_dl_all"] = (fn_, blob, sp_)
            if b2.button("📊 สร้างเฉพาะ Excel KPI", key="mr_btn_kpi_xl", **full):
                with st.spinner("กำลังสร้างไฟล์ Excel KPI…"):
                    blob, _ = mr.build_monthly_excel(db, sel_month, logo_path=logo_path, include_kpi=True, include_types=[])
                fn_ = f"สรุป_KPI_{sel_month}_{datetime.now():%Y%m%d_%H%M}.xlsx"
                sp_ = mr.save_local_export(app_dir, fn_, blob, is_local=is_local)
                ss["mr_dl_kpi"] = (fn_, blob, sp_)
            kpi_html = mr.build_print_html(db, sel_month, section="kpi", logo_path=logo_path)
            b3.download_button(
                "🖨️ ดาวน์โหลดหน้าพิมพ์ KPI (.html → PDF)",
                kpi_html.encode("utf-8"),
                file_name=f"พิมพ์_KPI_{sel_month}.html",
                mime="text/html",
                key="mr_dl_kpi_html",
                **full,
            )
            if ss.get("mr_dl_all"):
                _a = ss["mr_dl_all"]
                fn, blob, sp_ = _a[0], _a[1], (_a[2] if len(_a) > 2 else None)
                if sp_:
                    st.success(f"💾 บันทึกไฟล์แล้วที่: `{sp_}`")
                st.download_button(
                    f"⬇️ ดาวน์โหลด {fn}",
                    blob,
                    file_name=fn,
                    type="primary",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="mr_save_all_xl",
                )
            if ss.get("mr_dl_kpi"):
                _k = ss["mr_dl_kpi"]
                fn, blob, sp_ = _k[0], _k[1], (_k[2] if len(_k) > 2 else None)
                if sp_:
                    st.success(f"💾 บันทึกไฟล์แล้วที่: `{sp_}`")
                st.download_button(
                    f"⬇️ ดาวน์โหลด {fn}",
                    blob,
                    file_name=fn,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="mr_save_kpi_xl",
                )

        with st.expander("👁️ ตัวอย่างหน้าพิมพ์ KPI (กดปุ่ม 'พิมพ์ / บันทึก PDF' ด้านในได้ทันที)", expanded=True):
            _render_html_preview(kpi_html, height=620)

    # ========================================================================= TAB 2: FORMS (1_หัวลาก / 2_หางพ่วง / 3_ซ่อมตู้)
    with t_forms:
        choice_labels = {
            "forms": "ทั้งหมด (1_หัวลาก + 2_หางพ่วง + 3_ซ่อมตู้)",
            "หัวลาก": "1_หัวลาก",
            "หางพ่วง": "2_หางพ่วง",
            "ตู้เย็น": "3_ซ่อมตู้",
        }
        sel_form = st.radio(
            "เลือกแบบฟอร์มที่ต้องการพิมพ์",
            list(choice_labels.keys()),
            format_func=lambda k: choice_labels[k],
            horizontal=True,
            key="mr_form_sel",
        )
        sel_types = TYPES if sel_form == "forms" else [sel_form]

        with st.container(border=True):
            st.markdown(
                f"**พิมพ์แบบฟอร์มตาม Template บริษัท — {choice_labels[sel_form]} ({rep['month_th']})**  \n"
                "จัดหน้า A4 แนวตั้ง · หัวกระดาษและโลโก้ซ้ำทุกหน้า · กรณีรถ 1 คันซ่อมหลายใบในเดือนแสดงทุกเลขที่เอกสารคั่นด้วย `, `"
            )
            c1, c2 = st.columns(2)
            if c1.button("📥 สร้างไฟล์ Excel พร้อมพิมพ์ (.xlsx)", type="primary", key="mr_btn_form_xl", **full):
                with st.spinner("กำลังสร้างไฟล์ Excel พร้อมพิมพ์…"):
                    blob, _ = mr.build_monthly_excel(
                        db, sel_month, logo_path=logo_path, include_kpi=(sel_form == "forms"), include_types=sel_types
                    )
                tag = "1_2_3_ทั้งหมด" if sel_form == "forms" else mr.FORM_SPECS[sel_form]["sheet_name"]
                fn_ = f"รายงานซ่อม_{tag}_{sel_month}_{datetime.now():%Y%m%d_%H%M}.xlsx"
                sp_ = mr.save_local_export(app_dir, fn_, blob, is_local=is_local)
                ss["mr_dl_form"] = (fn_, blob, sp_)
            form_html = mr.build_print_html(db, sel_month, section=sel_form, logo_path=logo_path)
            tag_h = "1_2_3_ทั้งหมด" if sel_form == "forms" else mr.FORM_SPECS[sel_form]["sheet_name"]
            c2.download_button(
                "🖨️ ดาวน์โหลดหน้าพิมพ์ / Save PDF (.html)",
                form_html.encode("utf-8"),
                file_name=f"พิมพ์_{tag_h}_{sel_month}.html",
                mime="text/html",
                key="mr_dl_form_html",
                **full,
            )
            if ss.get("mr_dl_form"):
                _f = ss["mr_dl_form"]
                fn, blob, sp_ = _f[0], _f[1], (_f[2] if len(_f) > 2 else None)
                if sp_:
                    st.success(f"💾 บันทึกไฟล์แล้วที่: `{sp_}`")
                st.download_button(
                    f"⬇️ ดาวน์โหลด {fn}",
                    blob,
                    file_name=fn,
                    type="primary",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="mr_save_form_xl",
                )

        # ---- Unmatched review section for DATA_OWNER (คุณแขก)
        unm_list = []
        for at in sel_types:
            unm_list.extend(rep["by_type"][at]["unmatched"])
        if unm_list:
            with st.expander(
                f"⚠️ รายการนอก Master ในเดือน {sel_month} ({len(unm_list)} คัน/ตู้) — แยกไว้ให้ DATA_OWNER (คุณแขก) ตัดสินใจ",
                expanded=True,
            ):
                st.caption(
                    "ทะเบียน/เบอร์ตู้ด้านล่างมีใบแจ้งซ่อมในเดือนนี้ แต่ไม่ได้อยู่ในทะเบียนรถหลัก (Master) "
                    "จึง **ไม่ถูกรวมในแบบฟอร์มพิมพ์** (อาจเป็นรถประเภทอื่นปนมา หรือรถเข้าใหม่) — "
                    "หากคุณแขกยืนยันว่าเป็นรถในกลุ่มจริง กดปุ่ม 'เพิ่มเข้า Master' ได้ทันที"
                )
                udf = pd.DataFrame([
                    {
                        "ประเภท": x["asset_type"],
                        "ทะเบียน/เบอร์ตู้": x["asset_code"],
                        "ทะเบียนต้นฉบับ": x["asset_raw"],
                        "จำนวนใบซ่อม": x["repair_count"],
                        "เลขที่ใบแจ้งซ่อม": x["doc_text"],
                        "ยี่ห้อ": x["brands"],
                        "สถานะใน Master": "ปิดใช้งานอยู่" if x["in_master_inactive"] else "ยังไม่อยู่ใน Master",
                    }
                    for x in unm_list
                ])
                st.dataframe(udf, hide_index=True, **full)

                st.markdown("**ตัดสินใจเพิ่มเข้า Master (เฉพาะรายการที่ยืนยันว่าเป็นรถในกลุ่มจริง):**")
                for idx_u, x in enumerate(unm_list):
                    u1, u2, u3 = st.columns([2.2, 3.8, 1.4])
                    u1.markdown(f"**{x['asset_type']}** · `{x['asset_code']}` ({x['repair_count']} ใบ)")
                    u2.caption(f"{x['doc_text']}  ·  {x['asset_raw'] or ''}")
                    if u3.button("➕ เพิ่มเข้า Master", key=f"mr_add_unm_{sel_month}_{x['asset_type']}_{x['asset_code']}_{idx_u}", **full):
                        ok, msg = mr.upsert_asset_master(
                            db, x["asset_type"], x["asset_code"], is_active=1,
                            note=f"อนุมัติเพิ่มจากรายการนอก Master เดือน {sel_month}",
                            user=user or ("local" if is_local else "user"),
                        )
                        (st.toast if ok else st.error)(msg)
                        if ok:
                            st.rerun()
        else:
            st.success(f"ไม่มีรายการนอก Master ตกหล่นในเดือน {sel_month} สำหรับกลุ่มที่เลือก")

        with st.expander("👁️ ตัวอย่างหน้าพิมพ์แบบฟอร์ม (กดปุ่ม 'พิมพ์ / บันทึก PDF' ด้านในได้ทันที)", expanded=True):
            _render_html_preview(form_html, height=780)


# ============================================================================= หน้า 📋 รายการอุปกรณ์ (แถบแรกของแถบซ้าย)
def render_asset_master(is_local, user, full, app_dir, kpi_card):
    db = rrp.db_selector(is_local, app_dir, full)
    actor = user or ("local" if is_local else "user")

    st.markdown(
        '<div class="app-hero"><h1>📋 รายการอุปกรณ์ (Asset Master)</h1>'
        '<p>ทะเบียนรถและอุปกรณ์หลักทั้งหมด (หัวลาก · หางพ่วง · ตู้เย็น) — ใช้เป็นฐานอ้างอิงสำหรับแบบฟอร์มซ่อมประจำเดือนและ KPI · '
        'เพิ่ม / ปรับสถานะ / วางรายชื่อชุดใหม่ / นำเข้าจาก Excel และ Export</p></div>',
        unsafe_allow_html=True,
    )

    if not os.path.exists(db):
        st.warning("ยังไม่มีไฟล์ฐานข้อมูล — กรุณาเลือกไฟล์ .sqlite ที่แถบซ้าย หรือนำเข้าใบแจ้งซ่อมชุดแรกก่อน")
        return

    master_sum = mr.get_master_summary(db)
    total_active = sum(v["active"] for v in master_sum.values())
    total_inactive = sum(v["inactive"] for v in master_sum.values())

    kcols = st.columns(4)
    kpi_card(kcols[0], "อุปกรณ์ใช้งานทั้งหมด", f"{total_active:,}", f"ปิดใช้งาน/ปลดระวาง {total_inactive:,}", "ok" if total_active else "warn")
    for i, at in enumerate(TYPES, start=1):
        ms = master_sum[at]
        kpi_card(kcols[i], at, f"{ms['active']:,}", f"ปิดใช้งาน {ms['inactive']:,}")
    st.write("")

    t_list, t_edit, t_imp = st.tabs([
        f"📋 รายชื่ออุปกรณ์ทั้งหมด ({total_active:,})",
        "➕ เพิ่ม / แก้ไข / วางรายชื่อชุดใหม่",
        "📥 นำเข้าจากไฟล์ Excel",
    ])

    # ---- Tab 1: Browse & Export
    with t_list:
        fc1, fc2, fc3, fc4 = st.columns([1.3, 2.0, 1.3, 1.4])
        ft_m = fc1.selectbox("ประเภท", ["ทั้งหมด"] + TYPES, key="am_ft")
        fq_m = fc2.text_input("ค้นหาทะเบียน/เบอร์ตู้", key="am_fq", placeholder="เช่น 60-0000 หรือ ABCU")
        show_inact = fc3.checkbox("แสดงที่ปิดใช้งานด้วย", value=True, key="am_inact")
        m_rows = mr.list_asset_master(db, asset_type=None if ft_m == "ทั้งหมด" else ft_m, include_inactive=show_inact)
        mdf = pd.DataFrame(m_rows)
        if len(mdf):
            if fq_m:
                mdf = mdf[mdf["asset_code"].str.contains(fq_m.strip(), case=False, regex=False)]
            mdf["สถานะ"] = mdf["is_active"].map(lambda v: "ใช้งาน" if v else "ปิดใช้งาน")
            mdf = mdf.rename(columns={
                "asset_type": "ประเภท", "asset_code": "ทะเบียน/เบอร์ตู้", "sort_order": "ลำดับ",
                "note": "หมายเหตุ", "updated_at": "อัปเดตล่าสุด", "updated_by": "โดย",
            })[["ลำดับ", "ประเภท", "ทะเบียน/เบอร์ตู้", "สถานะ", "หมายเหตุ", "อัปเดตล่าสุด", "โดย"]]
        xl_data = mr.export_asset_master_excel(db)
        fc4.download_button(
            "⬇️ Export Excel",
            xl_data,
            file_name=f"รายการอุปกรณ์_Master_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="am_dl_xl",
            **full,
        )
        st.caption(f"แสดง {len(mdf):,} รายการ")
        st.dataframe(mdf, hide_index=True, height=480, **full)

    # ---- Tab 2: Single Add/Edit + Bulk Paste
    with t_edit:
        with st.container(border=True):
            st.markdown("**1) เพิ่มทะเบียนใหม่ หรือปรับสถานะ (ใช้งาน / ปลดระวาง) รายคัน**")
            with st.form("am_add_single"):
                f1, f2, f3, f4 = st.columns([1.2, 1.6, 1.3, 2.0])
                at_in = f1.selectbox("ประเภท", TYPES)
                code_in = f2.text_input("ทะเบียน / เบอร์ตู้", placeholder="เช่น 60-0000 หรือ ABCU1234567")
                act_in = f3.selectbox("สถานะ", [1, 0], format_func=lambda v: "ใช้งาน (Active)" if v == 1 else "ปิดใช้งาน (ปลดระวาง)")
                note_in = f4.text_input("หมายเหตุ (ถ้ามี)")
                if st.form_submit_button("💾 บันทึกรายการ", type="primary"):
                    ok, msg = mr.upsert_asset_master(db, at_in, code_in, is_active=act_in, note=note_in or None, user=actor)
                    (st.success if ok else st.error)(msg)
                    if ok:
                        st.rerun()

        with st.container(border=True):
            st.markdown("**2) วางรายชื่อหลายคันพร้อมกัน (1 บรรทัด = 1 ทะเบียน/เบอร์ตู้)**")
            with st.form("am_bulk_form"):
                b1, b2 = st.columns([1.4, 2.6])
                at_b = b1.selectbox("ประเภทอุปกรณ์", TYPES, key="am_b_type")
                mode_b = b2.radio(
                    "โหมดการบันทึก",
                    ["เพิ่ม/เปิดใช้งานเฉพาะรายการที่วาง", "แทนที่ทั้งหมดตามรายชื่อนี้เท่านั้น (ลบรายการเดิมของประเภทนี้)"],
                    horizontal=True,
                    key="am_b_mode",
                )
                txt_b = st.text_area("วางรายชื่อทะเบียน/เบอร์ตู้ (คัดลอกจาก Excel มาวางได้เลย)", height=180, placeholder="60-0001\n60-0002\n60-0003")
                note_b = st.text_input("หมายเหตุการอัปเดต", value="อัปเดตรายชื่อจากหน้าจัดการอุปกรณ์")
                confirm_b = st.checkbox("ยืนยันบันทึก (ระบบจะสำรองฐานข้อมูลอัตโนมัติก่อนบันทึก)", key="am_b_ok")
                if st.form_submit_button("💾 บันทึกรายชื่อชุดนี้", type="primary"):
                    if not confirm_b:
                        st.warning("กรุณาติ๊กยืนยันก่อนบันทึก")
                    else:
                        replace_all = mode_b.startswith("แทนที่")
                        res = mr.bulk_update_asset_master(db, at_b, txt_b.splitlines(), replace_all=replace_all, note=note_b, user=actor)
                        if res["ok"]:
                            st.success(f"บันทึก {at_b} สำเร็จ: ทั้งหมด {res['total']:,} รายการ (เพิ่ม {res['inserted']:,}, อัปเดต {res['updated']:,})")
                            st.rerun()
                        else:
                            st.error(res["msg"])

    # ---- Tab 3: Import from Excel Template
    with t_imp:
        st.markdown("**นำเข้า / อัปเดตรายชื่ออุปกรณ์จากไฟล์ Excel บริษัท** (รองรับชีต `F1_หัวลาก / F2_หางพ่วง / F3_ตู้เย็น` หรือ `1_หัวลาก / 2_ห่างพ่วง / 3_ซ่อมตู้`)")
        up_m = st.file_uploader("เลือกไฟล์ Template บริษัท (.xlsx)", type=["xlsx", "xlsm"], key="am_m_up")
        rep_mode = st.checkbox(
            "ปิดใช้งาน (ปลดระวาง) ทะเบียนเดิมที่ไม่มีในไฟล์นี้ด้วย (ไม่ติ๊ก = เพิ่ม/เปิดใช้งานเฉพาะรายการในไฟล์)",
            value=False,
            key="am_m_rep",
        )
        if st.button("📥 นำเข้าทะเบียนอุปกรณ์", type="primary", disabled=(up_m is None), key="am_m_go") and up_m is not None:
            with st.spinner("กำลังอ่านและบันทึกลง asset_master…"):
                res = mr.import_asset_master_from_excel(db, io.BytesIO(up_m.getvalue()), user=actor, replace_mode=rep_mode)
            if res["ok"]:
                summary_txt = " · ".join(f"{k}: ทั้งหมด {v['total']:,} (ใหม่ {v['inserted']:,}, อัปเดต {v['updated']:,})" for k, v in res["counts"].items())
                st.success(f"นำเข้าทะเบียนอุปกรณ์สำเร็จ: {summary_txt}")
                st.rerun()
            else:
                for e in res["errors"]:
                    st.error(e)
