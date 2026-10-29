# ACE Content Portal (Google Apps Script)

เว็บแอปสำหรับจัดการและติดตามข้อมูล ACE Content พัฒนาด้วย Google Apps Script, HTML, Tailwind CSS และ Vanilla JavaScript

## ความสามารถหลัก

- แสดงแดชบอร์ด KPI สำหรับผู้ดูแลระบบและ VJ/MC
- รองรับการเข้าสู่ระบบของ VJ/MC และผู้ดูแลระบบ พร้อมตรวจสอบสิทธิ์
- จัดการ User Mapping และดูประวัติแยกตามวัน
- อ่านข้อมูลเวลาเข้า–ออกงานจาก Google Sheets ภายนอก และบันทึกรายการแก้ไขในชีตของระบบ
- จัดการประกาศข่าวสาร แบนเนอร์ รูปภาพ และการยืนยันรับทราบ
- จัดการการล็อกบัญชีผู้ใช้และ IP Address พร้อมบันทึกเหตุการณ์การใช้งาน
- รองรับเมนูด้านข้างแบบปรับตามขนาดหน้าจอ และช่องค้นหาในรายการแบบดรอปดาวน์

## โครงสร้างไฟล์

    .
    ├── code.gs
    ├── index.html
    ├── appsscript.json
    ├── scripts/
    │   ├── setup-gas.ps1
    │   ├── upload-dev.ps1
    │   └── release-prod.ps1
    ├── docs/
    │   ├── DEPLOYMENT.md
    │   └── PROJECT_GUIDE.md
    └── README.md

## อัปโหลดไปยัง Google Apps Script

ต้องติดตั้ง Node.js และ clasp ก่อน โดยเปิด PowerShell ในโฟลเดอร์โปรเจกต์:

    npm install --global @google/clasp
    clasp login

1. ดู Script ID ได้จาก Apps Script → Project Settings → IDs แล้วตั้งค่าปลายทางครั้งแรก:

    .\scripts\setup-gas.ps1 -ScriptId "SCRIPT_ID"

2. ตรวจสอบไฟล์ที่จะส่งและอัปโหลดเพื่อทดสอบ /dev:

    .\scripts\upload-dev.ps1

   สคริปต์จะแสดงปลายทางและรายการไฟล์ แล้วรอให้พิมพ์ PUSH ก่อนเขียนทับไฟล์ใน Apps Script การ push จะทำให้ /dev ใช้ซอร์สล่าสุด แต่ไม่เปลี่ยนเวอร์ชันของ /exec

3. หลังทดสอบ /dev แล้ว หากต้องการอัปเดตเว็บแอป /exec ให้สร้างเวอร์ชันใน Apps Script หรือใช้ clasp version "คำอธิบายการปล่อยเวอร์ชัน" จากนั้นระบุ Deployment ID และเลขเวอร์ชัน:

    .\scripts\release-prod.ps1 -DeploymentId "DEPLOYMENT_ID" -VersionNumber "VERSION_NUMBER"

   สคริปต์จะขอให้พิมพ์ DEPLOY ก่อนเปลี่ยน Deployment ที่ระบุ

ดูขั้นตอน สิทธิ์ และแนวทางแก้ปัญหาเพิ่มเติมที่ คู่มืออัปโหลดและปล่อยเวอร์ชัน (docs/DEPLOYMENT.md)

## ข้อมูลและความปลอดภัย

- ระบบใช้ Google Sheets เดิมและต้องคงโครงสร้างข้อมูลที่ใช้อยู่ โดยเฉพาะคอลัมน์ A–J ใน User_Mapping
- ห้ามบันทึก OAuth Token, API Key, รหัสผ่านจริง หรือไฟล์ .clasp.json ลงใน Git
- .claspignore จำกัดไฟล์ที่ส่งไป Apps Script ให้มีเพียง code.gs, index.html และ appsscript.json
- URL /dev ใช้ทดสอบและจำกัดการเข้าถึงไว้สำหรับผู้แก้ไขโปรเจกต์ Apps Script
- ตรวจสอบการตั้งค่า Execute as และ Who has access ของ Web App ให้ตรงกับนโยบาย Google Workspace

## การพัฒนา

แก้ไข code.gs และ index.html ในคลังนี้ แล้วใช้สคริปต์อัปโหลดตามขั้นตอนข้างต้น การ push ไม่ได้สร้างเวอร์ชันใหม่ของ /exec โดยอัตโนมัติ

รันทดสอบ Unit/Integration แบบจำลองของ User Mapping ด้วย Node.js:

    npm test

ผลและข้อจำกัดการทดสอบล่าสุดอยู่ใน `docs/TEST_REPORT_TH.md` การทดสอบแบบจำลองไม่แทนการทดสอบหน้า Web App และ Google Sheets จริง
