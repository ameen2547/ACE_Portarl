# คู่มืออัปโหลดและปล่อย ACE Content Portal

## สิ่งที่ต้องเตรียม

- Node.js และ npm
- สิทธิ์ Editor ในโปรเจกต์ Google Apps Script เป้าหมาย
- Apps Script API เปิดใช้งานสำหรับบัญชี Google ที่ใช้กับ clasp
- สิทธิ์เข้าถึง Google Sheets และ Google Drive ที่ระบบ ACE ใช้งาน
- Script ID จากหน้า Apps Script → Project Settings → IDs

URL /dev หรือ /exec ที่เปิดเว็บแอปได้มี Deployment ID อยู่ใน URL ซึ่งไม่ใช่ Script ID สำหรับ .clasp.json อย่านำ Deployment ID มาใส่เป็น Script ID

## ติดตั้งและผูกโปรเจกต์

ติดตั้งเครื่องมือและเข้าสู่ระบบครั้งแรก:

    npm install --global @google/clasp
    clasp login

จากโฟลเดอร์คลังซอร์ส ตั้งค่า Script ID:

    .\scripts\setup-gas.ps1 -ScriptId "SCRIPT_ID"

ไฟล์ .clasp.json ถูกยกเว้นจาก Git เพื่อไม่ให้การตั้งค่าปลายทางเฉพาะเครื่องถูก commit หากต้องเปลี่ยนปลายทาง สคริปต์จะขอพิมพ์ REPLACE ก่อนเขียนทับ

หากโปรเจกต์ Apps Script มีไฟล์เดิมที่ไม่ได้อยู่ในคลังนี้ ให้ตรวจสอบก่อนอัปโหลด สคริปต์นี้ส่งเฉพาะ code.gs, index.html และ appsscript.json โดยตั้งใจ ไม่ลบไฟล์อื่นบน Apps Script

## อัปโหลดเพื่อทดสอบ /dev

    .\scripts\upload-dev.ps1

ตรวจสอบ Script ID และรายการไฟล์ที่ clasp status แสดง จากนั้นพิมพ์ PUSH เพื่อยืนยัน การอัปโหลดใช้ clasp push โดยไม่ใช้ --force เพื่อไม่เขียนทับ manifest ฝั่งเซิร์ฟเวอร์เมื่อมีความขัดแย้งที่ต้องตรวจสอบ

หลังอัปโหลด เปิด URL /dev และทดสอบด้วยบัญชีผู้แก้ไขโปรเจกต์ เนื่องจากการแก้ไข manifest หรือบริการที่เรียกใช้อาจทำให้ Google ขออนุญาตใหม่ ให้ตรวจสอบรายการสิทธิ์ก่อนอนุญาตด้วยบัญชีเจ้าของระบบ

## ปล่อยหรืออัปเดต /exec

/exec ใช้เวอร์ชันที่เผยแพร่ จึงไม่เปลี่ยนตาม clasp push เพียงอย่างเดียว

1. ทดสอบโค้ดบน /dev ให้เรียบร้อย
2. สร้างเวอร์ชันใน Apps Script หรือใช้:

    clasp version "คำอธิบายเวอร์ชัน"

3. ดูหมายเลขเวอร์ชันและ Deployment ID ด้วยคำสั่ง:

    clasp versions
    clasp deployments

4. ระบุ Deployment ID ที่ต้องการอัปเดตและหมายเลขเวอร์ชัน:

    .\scripts\release-prod.ps1 -DeploymentId "DEPLOYMENT_ID" -VersionNumber "VERSION_NUMBER"

5. พิมพ์ DEPLOY เพื่อยืนยันการเปลี่ยน /exec ของ Deployment นั้น

การอัปเดต Deployment เดิมจะคง URL เดิมไว้ หากต้องการปล่อยเป็น URL ใหม่ ให้จัดการ deployment ผ่าน Apps Script ตามกระบวนการของผู้ดูแล

## คำสั่ง npm แบบย่อ

    npm run gas:setup -- -ScriptId "SCRIPT_ID"
    npm run gas:push:dev
    npm run gas:release:prod -- -DeploymentId "DEPLOYMENT_ID" -VersionNumber "VERSION_NUMBER"

## ปัญหาที่พบบ่อย

- ยังไม่ได้เข้าสู่ระบบ: รัน clasp login ด้วยบัญชีที่มีสิทธิ์ Editor ในโปรเจกต์ปลายทาง
- ไม่พบ Script ID หรือสิทธิ์ไม่พอ: ตรวจว่า .clasp.json ใช้ Script ID จาก Project Settings และบัญชีมีสิทธิ์เข้าถึงโปรเจกต์นั้น
- Apps Script API ปิดอยู่: เปิด Apps Script API ในการตั้งค่าบัญชี Google ตามข้อความแจ้งจาก clasp
- manifest conflict: ตรวจ manifest ฝั่ง Apps Script ก่อน แล้วค่อยตัดสินใจว่าจะปรับไฟล์ในคลังหรือรวมค่าที่จำเป็น หลีกเลี่ยง clasp push --force หากยังไม่ทราบผลกระทบ
- /dev ไม่เปลี่ยน: ตรวจว่ากำลังเปิด URL /dev ของ Script ID เดียวกับที่ตั้งค่าไว้ แล้ว reload แบบไม่ใช้ cache
- /exec ยังเป็นโค้ดเดิม: สร้างหมายเลขเวอร์ชันและอัปเดต Deployment ที่ถูกต้องตามขั้นตอนด้านบน
