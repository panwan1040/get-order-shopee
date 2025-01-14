# 📋 สคริปต์ดึงข้อมูลคำสั่งซื้อ Shopee

---

## 🚀 **คุณสมบัติ**
- ดึงข้อมูลคำสั่งซื้อ เช่น ชื่อสินค้า จำนวน ราคา และสถานะคำสั่งซื้อ
- เรียกข้อมูลจาก API ของ Shopee โดยตรง
- สร้างไฟล์ Excel ชื่อ `shopee_orders.xlsx` ซึ่งมีข้อมูลคำสั่งซื้อในรูปแบบที่อ่านง่าย

---

## 🛠️ **วิธีใช้งาน**

1. **เปิดเว็บไซต์ Shopee**  
   ไปที่ [Shopee](https://shopee.co.th/) และเข้าสู่ระบบด้วยบัญชีของคุณ

2. **เปิดเครื่องมือสำหรับนักพัฒนา**  
   กดปุ่ม `F12` (หรือ `Ctrl + Shift + I` / `Cmd + Option + I` สำหรับ Mac) เพื่อเปิด Developer Tools

3. **เลือกแท็บ Console**  
   คลิกที่แท็บ **Console** ในหน้าต่าง Developer Tools

4. **รันสคริปต์**  
   คัดลอกและวางโค้ดด้านล่างลงในหน้าต่าง Console แล้วกด `Enter`:

   ```javascript
   (async function() {
       // (โค้ดทั้งหมดจากด้านบนนี้)
   })();
   ```

5. **ดาวน์โหลดไฟล์ Excel**  
   เมื่อสคริปต์ทำงานเสร็จ ไฟล์ Excel ชื่อ `shopee_orders.xlsx` จะถูกดาวน์โหลดลงในเครื่องของคุณ

---

## 📦 **สิ่งที่ต้องใช้**

สคริปต์นี้โหลดไลบรารี `xlsx` แบบไดนามิกเพื่อใช้สร้างไฟล์ Excel  
โปรดตรวจสอบว่าอุปกรณ์ของคุณเชื่อมต่ออินเทอร์เน็ตระหว่างการใช้งาน

---

## ⚠️ **ข้อควรระวัง**
- คุณต้องเข้าสู่ระบบในบัญชี Shopee ก่อนใช้งานสคริปต์
- สคริปต์นี้อาศัยโครงสร้าง API ของ Shopee ซึ่งอาจมีการเปลี่ยนแปลงในอนาคต
- เพื่อความปลอดภัย โปรดใช้งานสคริปต์นี้ด้วยความระมัดระวัง และหลีกเลี่ยงการแชร์ข้อมูลการเข้าสู่ระบบหรือข้อมูลเบราว์เซอร์ของคุณกับผู้อื่น

---

## 🛡️ **คำปฏิเสธความรับผิดชอบ**

สคริปต์นี้ถูกจัดทำขึ้นเพื่อการส่วนตัวและผู้เขียนไม่เกี่ยวข้องกับ Shopee และไม่รับผิดชอบหาก API ของ Shopee มีการเปลี่ยนแปลงจนทำให้สคริปต์ไม่สามารถทำงานได้

---

นำไปใช้และปรับแต่งได้ตามความต้องการ! 😊
