# Vehicle Report Importer (Speed Fleet Maintenance)

Streamlit web app สำหรับอ่านรายงาน Crystal Reports "Finish Goods Takeoff and Work Order Report by Vehicle"
(`rpt_SumWO_Date.rpt`) รายคัน (.xls / .xlsx) → กระทบยอด → ติด Confirm Flag → Export Excel

> **Repo นี้เป็น Public — เก็บเฉพาะโค้ด ห้าม commit ข้อมูลจริง** (ไฟล์รายงาน, Excel ที่ export, ฐานข้อมูล, รหัสผ่าน) — `.gitignore` กันไฟล์ .xls/.xlsx/.zip/.csv/.sqlite และ secrets ไว้แล้ว
> ถ้า deploy บน Streamlit Community Cloud ให้ตั้งแอปเป็น **Private** (เชิญเฉพาะอีเมลที่อนุญาต) เพราะผู้ใช้จะอัปโหลดข้อมูลจริงเข้าแอป

## หน้าในแอป (7 หน้า)
1. **📋 รายการอุปกรณ์** — จัดการทะเบียนอุปกรณ์หลัก (Master: หัวลาก / หางพ่วง / ตู้เย็น) เพิ่ม/แก้ไข/วางชุดใหม่/นำเข้า Excel และ Export (`monthly_report_page.py`, `monthly_report.py`)
2. **🚛 รายงานซ่อมรายคัน** — อ่าน Crystal Report รายคัน → กระทบยอด → Export Excel (`vehicle_report_web.py`, `vehicle_report_app.py`)
3. **🗂️ ฐานข้อมูลใบแจ้งซ่อม** — นำเข้าใบแจ้งซ่อม (ชีต M1/M2/M3 หรือ หัวลาก/หางพ่วง/ตู้เย็น) เข้า SQLite กันซ้ำด้วยเลขที่เอกสาร, ตรวจก่อนบันทึก, backup อัตโนมัติ, แท็บ "รอตรวจ" ให้ DATA_OWNER ตัดสิน, คอลัมน์ `ISO (Yes/No)`, KPI คัน/เดือน และ Export M1/M2/M3 (เฉพาะ `ISO = Yes`) (`repair_request_page.py`, `repair_request_db.py`)
4. **🖨️ พิมพ์รายงานประจำเดือน** — แถบ **พิมพ์ KPI ประจำเดือน** (Export Excel 13 ชีตตาม Template: `Index`, `MTD-*`, `ทะเบียนทั้งหมด`, `สรุป_KPI_ประจำเดือน`, `1_หัวลาก..3_ซ่อมตู้`, `M1..M3`, `Last PM`), หน้าสรุป KPI/PDF และแบบฟอร์มรายงานซ่อมประจำเดือน (`monthly_report_page.py`, `monthly_report.py`)
5. **📏 เลขไมล์** — นำเข้าและตรวจสอบประวัติเลขไมล์รายคัน (ใช้ฐานข้อมูล SQLite เดียวกัน)
6. **🛠️ PM Plan รายเดือน** — ประวัติบริการจากไฟล์ Export รายงานซ่อมรายคัน (W/O ศูนย์บริการ 20,000 กม. / WX ลาน 30,000 กม. หรือ 6 เดือน อย่างใดถึงก่อน) และคำนวณกำหนด PM (`pm_plan.py`, `pm_plan_page.py`)
7. **📖 คู่มือการใช้งาน** — แสดงคู่มือในหน้าแอป (`static/manual.html`) หรือเปิดแท็บใหม่ที่ `?manual=1`

## ไฟล์
| ไฟล์ | หน้าที่ |
|---|---|
| `vehicle_report_web.py` | หน้าเว็บหลัก (Main file สำหรับ Streamlit Cloud — บนเครื่องอยู่ใต้ `Scrips/`) |
| `vehicle_report_app.py` | ตัวอ่านรายงาน Crystal, กระทบยอด, เลขไมล์, Export Excel (+ GUI Tkinter บน Windows) |
| `repair_request_db.py` | engine ฐานข้อมูลใบแจ้งซ่อม & เลขไมล์ (กันซ้ำ, review, backup, `in_iso`, Export M1/M2/M3) |
| `repair_request_page.py` | UI หน้า "ฐานข้อมูลใบแจ้งซ่อม" และ "เลขไมล์" |
| `monthly_report.py` | engine รายการอุปกรณ์ Master, รายงานประจำเดือน, KPI ISO 13 ชีต (`Index`, `MTD-*`, `Last PM`) — อ่านค่าเฉพาะองค์กรจาก `data/local_settings.json` |
| `monthly_report_page.py` | UI หน้า "รายการอุปกรณ์" และ "พิมพ์รายงานประจำเดือน" |
| `pm_plan.py` / `pm_plan_page.py` | engine และ UI หน้า "PM Plan รายเดือน" |
| `requirements.txt` | ไลบรารีที่ใช้ |
| `.streamlit/config.toml` | ปิดการส่งสถิติการใช้งาน, จำกัดขนาดอัปโหลด 200 MB |

## โหมดการทำงาน
- **Cloud (Linux):** อัปโหลดไฟล์ .xls/.xlsx หลายไฟล์ หรือ .zip ทั้งโฟลเดอร์ → ไฟล์ถูกอ่านในโฟลเดอร์ชั่วคราวแล้วลบทันที →
  Export เป็นไฟล์ดาวน์โหลด (ไม่บันทึกบนเซิร์ฟเวอร์)
- **เครื่อง Windows:** เลือกโฟลเดอร์ในเครื่อง (เช่น `D:\Vehicle_Reports`) และบันทึก Excel ลงโฟลเดอร์ (เช่น `D:\MT_Speed`) หรือใช้โหมดอัปโหลดก็ได้

## Deploy บน Streamlit Community Cloud
1. ไปที่ https://share.streamlit.io → Sign in with GitHub
2. **Create app** → เลือก repo นี้, branch `main`, Main file `vehicle_report_web.py`
3. Deploy แล้วเปิดแอป จะเจอหน้า **ตั้งค่าผู้ใช้ครั้งแรก** → กรอกชื่อผู้ใช้/รหัสผ่าน → กดสร้าง → คัดลอกข้อความที่ได้
   แล้ววางที่ *Manage app → Settings → Secrets* (รูปแบบด้านล่าง) → Save → รีเฟรช
   ```toml
   [users]
   oo = "pbkdf2_sha256$200000$...."
   staff1 = "pbkdf2_sha256$200000$...."
   ```
   รหัสผ่านเก็บเป็น hash (PBKDF2-SHA256) เท่านั้น — ห้ามใส่ Secrets ลงใน repo
4. Deploy → เมื่อแอปขึ้นแล้ว ไปที่ **Settings → Sharing** ตรวจว่าเป็น *Only specific people can view this app* แล้วเพิ่มอีเมลผู้ใช้

## Login
- บน Cloud ต้อง login ทุกครั้ง (ถ้ายังไม่ตั้ง `[users]` แอปจะเปิดใช้งานไม่ได้)
- ใส่รหัสผิด 5 ครั้ง → รอ 60 วินาที
- รันบนเครื่อง Windows ไม่ต้อง login ยกเว้นสร้าง `.streamlit/secrets.toml` ที่มี `[users]`

## รันในเครื่อง
```
py -m pip install -r requirements.txt
py -m streamlit run vehicle_report_web.py --server.address localhost
```
