# บันทึกรายจ่ายคู่รัก

ระบบบันทึกรายจ่ายแยกคน + หารครึ่ง ผ่าน LINE OA โดยมี Google Sheet + Apps Script เป็นหลังบ้าน

## โครงสร้าง repo

| ไฟล์/โฟลเดอร์ | หน้าที่ |
|---|---|
| `index.html`, `manifest.json`, `sw.js`, `icon-*.png` | หน้าแอปสำหรับผู้ดูแล (เปิดผ่าน GitHub Pages ติดตั้งลงมือถือได้) |
| `apps-script/Code.gs` | โค้ดหลังบ้าน คัดลอกไปวางใน Apps Script ของ Google Sheet |
| `docs/คู่มือติดตั้ง.md` | คู่มือติดตั้งและวิธีใช้งานทั้งหมด |

## เปิดใช้ GitHub Pages

Settings → Pages → Source: **Deploy from a branch** → Branch: `main` / `(root)` → Save
เว็บจะอยู่ที่ `https://<ชื่อผู้ใช้>.github.io/<ชื่อ repo>/`

## ห้าม commit ข้อมูลเหล่านี้เด็ดขาด

ค่าทุกตัวใน Script Properties (`LINE_CHANNEL_ACCESS_TOKEN`, `CLAUDE_API_KEY`, `ADMIN_KEY`, LINE ID ต่างๆ) และ Web app URL — โค้ดใน repo นี้ไม่มีความลับอยู่ข้างใน ทุกอย่างอ่านจาก Script Properties
