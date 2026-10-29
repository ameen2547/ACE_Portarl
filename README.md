# ACE Content Portal (Google Apps Script)

ระบบ Web App สำหรับจัดการและติดตามข้อมูล ACE Content พัฒนาด้วย Google Apps Script, HTML, Tailwind CSS และ Vanilla JavaScript

## ฟังก์ชันหลัก

- แดชบอร์ด KPI สำหรับผู้ดูแลระบบและ VJ/MC
- เข้าสู่ระบบ VJ/MC และ Admin พร้อมตรวจสิทธิ์การใช้งาน
- จัดการ User Mapping และประวัติรายวัน
- ดูข้อมูลเวลาเข้า–ออกงานจากชีต Attendance ภายนอก และบันทึกข้อมูลแก้ไขในชีตของระบบ
- ประกาศข่าวสาร แบนเนอร์ รูปภาพ และการยืนยันรับทราบ
- จัดการการล็อกผู้ใช้/IP และบันทึกเหตุการณ์
- เมนู Sidebar แบบ Responsive และ dropdown ค้นหารายชื่อ

## โครงสร้างไฟล์

```text
.
├── code.gs                 # Backend API และการเชื่อมต่อ Google Sheets/Drive
├── index.html              # Frontend ของ Web App
├── README.md               # ภาพรวมและวิธีเริ่มต้น
└── docs/
    └── PROJECT_GUIDE.md    # คู่มือ Deploy, การตั้งค่า และการดูแลระบบ
```

## เริ่มต้นใช้งาน

1. เปิดหรือสร้างโปรเจกต์ที่ [script.google.com](https://script.google.com/).
2. วาง `code.gs` เป็นไฟล์สคริปต์ และสร้างไฟล์ HTML ชื่อ `index` แล้ววางเนื้อหาจาก `index.html`.
3. ตรวจสอบสิทธิ์ของบัญชีที่เป็นเจ้าของสคริปต์ต่อ Spreadsheet และ Drive ที่ระบบใช้งาน
4. บันทึกโปรเจกต์ แล้วเปิด **Deploy → Test deployments → Web app** เพื่อทดสอบ URL `/dev`.
5. เปิด URL `/dev` ด้วยบัญชีที่มีสิทธิ์ Editor ใน Apps Script project.

รายละเอียดการติดตั้งและข้อควรตรวจสอบอยู่ใน [คู่มือระบบ](docs/PROJECT_GUIDE.md).

## หมายเหตุด้านข้อมูลและความปลอดภัย

- ระบบออกแบบให้ใช้ Spreadsheet เดิมและคง Schema ที่มีอยู่ โดยเฉพาะคอลัมน์ A–J ของ `User_Mapping`.
- Backend มีการอ้างอิง Spreadsheet ID และมีค่า fallback PIN สำหรับการเข้าสู่ระบบในโค้ดปัจจุบัน
- Repository นี้เป็น **public**: ก่อนใช้งานจริง ให้ผู้ดูแลตรวจสอบ/เปลี่ยนค่าเริ่มต้นด้านรหัสผ่านและทบทวนการเปิดเผย Spreadsheet ID รวมถึงการตั้งค่าสิทธิ์ Deploy
- ห้ามใส่ OAuth token, API key, รหัสผ่านจริง หรือข้อมูลส่วนบุคคลใน Git
- URL `/dev` ใช้ทดสอบและจำกัดสิทธิ์เฉพาะผู้แก้ไขโปรเจกต์; สำหรับผู้ใช้ทั่วไป ให้เลือก deployment และ audience ให้ตรงนโยบายองค์กร

## การปรับปรุงโค้ด

ไม่มีขั้นตอน build สำหรับไฟล์ Frontend ปัจจุบัน หากแก้ไขใน repository ให้นำไฟล์ที่อัปเดตไปซิงก์กับโปรเจกต์ Apps Script แล้วบันทึกก่อนทดสอบ `/dev`.
