# Vehicle Report Importer (Speed Fleet Maintenance)

Streamlit web app สำหรับอ่านรายงาน Crystal Reports "Finish Goods Takeoff and Work Order Report by Vehicle"
(`rpt_SumWO_Date.rpt`) รายคัน (.xls / .xlsx) → กระทบยอด → ติด Confirm Flag → Export Excel

> **Repo นี้เป็น Public — เก็บเฉพาะโค้ด ห้าม commit ข้อมูลจริง** (ไฟล์รายงาน, Excel ที่ export, ฐานข้อมูล, รหัสผ่าน) — `.gitignore` กันไฟล์ .xls/.xlsx/.zip/.csv/.sqlite และ secrets ไว้แล้ว
> ถ้า deploy บน Streamlit Community Cloud ให้ตั้งแอปเป็น **Private** (เชิญเฉพาะอีเมลที่อนุญาต) เพราะผู้ใช้จะอัปโหลดข้อมูลจริงเข้าแอป

## ไฟล์
| ไฟล์ | หน้าที่ |
|---|---|
| `vehicle_report_web.py` | หน้าเว็บ (Main file สำหรับ Streamlit Cloud) |
| `vehicle_report_app.py` | ตัวอ่านรายงาน, กระทบยอด, เลขไมล์, Export Excel (+ GUI Tkinter เมื่อรันบน Windows) |
| `requirements.txt` | ไลบรารีที่ใช้ |
| `.streamlit/config.toml` | ปิดการส่งสถิติการใช้งาน, จำกัดขนาดอัปโหลด 200 MB |

## โหมดการทำงาน
- **Cloud (Linux):** อัปโหลดไฟล์ .xls/.xlsx หลายไฟล์ หรือ .zip ทั้งโฟลเดอร์ → ไฟล์ถูกอ่านในโฟลเดอร์ชั่วคราวแล้วลบทันที →
  Export เป็นไฟล์ดาวน์โหลด (ไม่บันทึกบนเซิร์ฟเวอร์)
- **เครื่อง Windows:** เลือกโฟลเดอร์ในเครื่อง (เช่น `D:\Vehicle_Reports`) และบันทึก Excel ลงโฟลเดอร์ (เช่น `D:\MT_Speet`) หรือใช้โหมดอัปโหลดก็ได้

## Deploy บน Streamlit Community Cloud
1. ไปที่ https://share.streamlit.io → Sign in with GitHub
2. **Create app** → เลือก repo นี้, branch `main`, Main file `vehicle_report_web.py`
3. (แนะนำ) Advanced settings → Secrets ใส่รหัสผ่านชั้นที่ 2:
   ```toml
   app_password = "ตั้งรหัสผ่านเอง"
   ```
4. Deploy → เมื่อแอปขึ้นแล้ว ไปที่ **Settings → Sharing** ตรวจว่าเป็น *Only specific people can view this app* แล้วเพิ่มอีเมลผู้ใช้

## รันในเครื่อง
```
py -m pip install -r requirements.txt
py -m streamlit run vehicle_report_web.py --server.address localhost
```
