# === 🛠️ Import Libraries ===
# นำเข้าไลบรารีที่จำเป็นสำหรับโปรแกรม
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import time
import os
import shutil
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# === ⚙️ ส่วนที่สามารถแก้ไขได้ (Configuration) ===
# ส่วนนี้คือตัวแปรหลักที่สามารถปรับเปลี่ยนได้ง่าย เพื่อให้สะดวกต่อการใช้งาน
# ทั้งหมดนี้จะถูกใช้ในส่วนต่างๆ ของโปรแกรม

# 🧾 ข้อมูลล็อกอิน
# กำหนดชื่อผู้ใช้, รหัสผ่าน, และ URL สำหรับการเข้าสู่ระบบและหน้าดาวน์โหลดรายงาน
USERNAME = "your_username"
PASSWORD = "your_password"
LOGIN_URL = "http://nmsgov.ntcsoc.net/Orion/Login.aspx"
REPORT_URL = "http://nmsgov.ntcsoc.net/Orion/reports/viewreports.aspx"

# 📥 โฟลเดอร์สำหรับบันทึกรายงาน
# กำหนดที่อยู่ของโฟลเดอร์สำหรับไฟล์ที่ดาวน์โหลด
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads", "temp_downloads")
TARGET_DIR =r"C:\path\to\your\target\folder"

# ✅ Path สำหรับ Edge Driver ที่ดาวน์โหลดด้วยตัวเอง
# หากการดาวน์โหลดอัตโนมัติล้มเหลว โปรแกรมจะใช้ driver จาก path นี้
edge_driver_path = r"C:\path\to\your\target\folder\msedgedriver.exe"

# 📧 ข้อมูลสำหรับส่ง Email
# ตั้งค่าข้อมูลอีเมลผู้ส่งและผู้รับ
# หมายเหตุ: ใช้ App Password แทนรหัสผ่านปกติเพื่อความปลอดภัย
EMAIL_ADDRESS = "your_email@gmail.com"
EMAIL_PASSWORD = "your_app_password"
RECIPIENT_EMAIL = "recipient_email@example.com"


# 🆕 เพิ่มการตั้งค่าเซิร์ฟเวอร์ SMTP และ Port
SMTP_SERVER = "smtp.gmail.com" 
SMTP_PORT = 587 

# 📝 เนื้อหาอีเมล
# กำหนดหัวข้อและเนื้อหาของอีเมล
# {report_name} และ {date} เป็นตัวแปรที่จะถูกแทนที่อัตโนมัติ
EMAIL_SUBJECT = "รายงาน Availability Last Month ประจำเดือน ({month}/{year})"
EMAIL_BODY = "เรียนผู้รับที่เกี่ยวข้อง\nรายงาน Availability Last Month \nแนบไฟล์ไว้ข้างล่างนี้\nขอบคุณค่ะ"


# === 📜 ฟังก์ชันแสดง Log ใน Console ===
# ฟังก์ชันนี้ใช้สำหรับแสดงข้อความสถานะในหน้าต่าง Terminal/Command Prompt
def log(text):
    print(text)
    sys.stdout.flush()

# === 📧 ฟังก์ชันสำหรับส่งอีเมลพร้อมไฟล์แนบ ===
# ฟังก์ชันนี้จะทำการเชื่อมต่อกับเซิร์ฟเวอร์ SMTP และส่งอีเมลพร้อมแนบไฟล์
def send_email_with_attachments(sender_email, sender_password, recipient_email, subject, body, attachment_paths):
    log("📧 กำลังเตรียมส่งอีเมล...")
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        # วนลูปเพื่อแนบไฟล์ทั้งหมดในรายการ
        for file_path in attachment_paths:
            part = MIMEBase('application', 'octet-stream')
            with open(file_path, 'rb') as file:
                part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file_path)}"')
            msg.attach(part)
            log(f"📎 แนบไฟล์: {os.path.basename(file_path)}")

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        text = msg.as_string()
        server.sendmail(sender_email, recipient_email, text)
        server.quit()
        log("✅ ส่งอีเมลสำเร็จแล้ว")
    except Exception as e:
        log(f"❌ เกิดข้อผิดพลาดในการส่งอีเมล: {e}")

# === ▶️ เริ่มดาวน์โหลด ===
# ฟังก์ชันหลักของโปรแกรมที่จัดการขั้นตอนการทำงานทั้งหมด
def start_download():
    log("▶️ เริ่มต้นโปรแกรมดาวน์โหลดรายงาน")

    start_time = time.time()
    success_count = 0
    fail_count = 0
    failed_reports = []
    driver = None
    daily_target_dir = ""
    downloaded_files = [] 

    try:
        # ลบและสร้างโฟลเดอร์ดาวน์โหลดชั่วคราวใหม่ เพื่อให้แน่ใจว่าไม่มีไฟล์เก่าค้างอยู่
        if os.path.exists(DOWNLOAD_DIR):
            shutil.rmtree(DOWNLOAD_DIR)
        os.makedirs(DOWNLOAD_DIR)
        
        # ตรวจสอบว่ามีโฟลเดอร์ปลายทางหรือไม่ ถ้าไม่มีให้สร้าง
        if not os.path.exists(TARGET_DIR):
            os.makedirs(TARGET_DIR)

        # ตั้งค่าสำหรับ Edge Browser
        edge_options = EdgeOptions()
        edge_options.add_argument("--start-maximized")
        edge_options.add_argument("--no-sandbox")
        edge_options.add_argument("--disable-web-security")
        edge_options.add_argument("--ignore-certificate-errors")
        edge_options.add_argument("--disable-infobars")
        edge_options.add_argument("--headless=new") 
        
        # ตั้งค่า Preference สำหรับการดาวน์โหลด
        prefs = {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "safebrowsing.enabled": False,
            "safebrowsing.disable_download_protection": True,
            "download.directory_upgrade": True,
        }
        edge_options.add_experimental_option("prefs", prefs)

        # ลองดาวน์โหลด Edge Driver อัตโนมัติด้วย webdriver_manager
        log("🚀 กำลังพยายามดาวน์โหลด Edge Driver อัตโนมัติ...")
        try:
            service = EdgeService(EdgeChromiumDriverManager().install())
            driver = webdriver.Edge(service=service, options=edge_options)
            log("✅ เปิด Edge Browser สำเร็จ (อัตโนมัติ)")
        except Exception as e:
            # หากดาวน์โหลดอัตโนมัติล้มเหลว จะใช้ driver ที่ระบุ path ไว้ด้านบน
            service = EdgeService(edge_driver_path)
            driver = webdriver.Edge(service=service, options=edge_options)
        
        driver.implicitly_wait(10) # ตั้งค่าการรอคอยแบบ Implicit

        # เข้าสู่ระบบด้วย Selenium
        log("🔐 กำลังเข้าสู่ระบบ...")
        driver.get(LOGIN_URL)
        WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.ID, "ctl00_BodyContent_Username")))
        driver.find_element(By.ID, "ctl00_BodyContent_Username").send_keys(USERNAME)
        driver.find_element(By.ID, "ctl00_BodyContent_Password").send_keys(PASSWORD)
        driver.find_element(By.ID, "ctl00_BodyContent_LoginButton").click()
        log("✅ เข้าสู่ระบบสำเร็จ")

        # ไปที่หน้า URL ของรายงาน
        driver.get(REPORT_URL)

        # 🆕 สร้างโฟลเดอร์ใหม่สำหรับเดือนปัจจุบันในโฟลเดอร์ปลายทาง
        daily_folder_name = datetime.now().strftime("%m%Y") 
        daily_target_dir = os.path.join(TARGET_DIR, daily_folder_name)
        if not os.path.exists(daily_target_dir):
            os.makedirs(daily_target_dir)
            log(f"📁 สร้างโฟลเดอร์สำหรับเดือนนี้: {daily_folder_name}")

        # เริ่มต้น Loop เพื่อดาวน์โหลดรายงานทีละหน้า
        while True:
            report_links = WebDriverWait(driver, 300).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, 'Report.aspx?ReportID=')]"))
            )
            if not report_links:
                break
            
            all_links = [(elem.get_attribute("href"), elem.text.strip()) for elem in report_links]
            log(f"📄 พบ {len(all_links)} รายงานในหน้านี้")

            # Loop เพื่อดาวน์โหลดแต่ละรายงาน
            for href, name in all_links:
                retry = 0
                success = False
                while retry < 2 and not success:
                    try:
                        driver.get(href)
                        log(f"🟢 กำลังเปิดรายงาน: {name} (รอบที่ {retry+1})")
                        
                        WebDriverWait(driver, 300).until_not(
                            EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Your requested report')]"))
                        )
                        log("✅ รายงานพร้อมแล้ว")

                        export_btn = WebDriverWait(driver, 300).until(
                            EC.element_to_be_clickable((By.LINK_TEXT, "Export to Excel"))
                        )
                        export_btn.click()
                        log("📥 คลิก Export เรียบร้อยแล้ว")
                        
                        # รอให้ไฟล์ดาวน์โหลดเสร็จสมบูรณ์
                        download_complete = False
                        start_download_time = time.time()
                        last_size = -1
                        stagnant_time = 0

                        while not download_complete and (time.time() - start_download_time < 900):
                            files = [f for f in os.listdir(DOWNLOAD_DIR) if not f.endswith('.tmp') and not f.endswith('.crdownload')]
                            if files:
                                latest_file = max(
                                    [os.path.join(DOWNLOAD_DIR, f) for f in files],
                                    key=os.path.getctime
                                )
                                current_size = os.path.getsize(latest_file)
                                if current_size == last_size:
                                    stagnant_time += 1
                                else:
                                    stagnant_time = 0
                                    last_size = current_size

                                if stagnant_time >= 3:
                                    log("❗ ไฟล์ดาวน์โหลดหยุดนิ่ง")
                                    break
                                
                                if not latest_file.endswith('.crdownload') and current_size > 0:
                                    download_complete = True
                                    break
                            time.sleep(2)
                        
                        if download_complete:
                            # เปลี่ยนชื่อไฟล์ที่ดาวน์โหลดแล้วย้ายไปยังโฟลเดอร์ปลายทาง
                            file_ext = os.path.splitext(latest_file)[1]
                            timestamp = datetime.now().strftime("%d%m%Y")
                            new_name = f"{timestamp}_{name.replace(' ', '_').replace('/', '-')}{file_ext}"
                            
                            destination_path = os.path.join(daily_target_dir, new_name)
                            shutil.move(latest_file, destination_path)
                            log(f"✅ บันทึกสำเร็จ: {new_name}")
                            success_count += 1
                            success = True
                            downloaded_files.append(destination_path) 

                        else:
                            raise Exception("การดาวน์โหลดล้มเหลวหรือใช้เวลานานเกินไป")
                            
                    except Exception as e:
                        # หากเกิดข้อผิดพลาด ให้ลองดาวน์โหลดใหม่อีกครั้ง
                        retry += 1
                        log(f"❌ ล้มเหลวรอบที่ {retry}: {name} -> {e}")
                        if retry >= 2:
                            fail_count += 1
                            failed_reports.append(name)
                    finally:
                        # กลับไปหน้าหลักของรายงานทุกครั้งหลังจบการดาวน์โหลด
                        driver.get(REPORT_URL)

            # ตรวจสอบว่ามีปุ่ม "Next" เพื่อไปยังหน้าถัดไปหรือไม่
            try:
                next_btn = driver.find_element(By.LINK_TEXT, "Next")
                next_btn.click()
                time.sleep(2)
            except:
                # ถ้าไม่พบปุ่ม "Next" แสดงว่าดาวน์โหลดครบทุกหน้าแล้ว ให้หยุด Loop
                break

    except Exception as e:
        log(f"❌ เกิดข้อผิดพลาดที่ไม่คาดคิด: {e}")

    finally:
        # ปิดเบราว์เซอร์เมื่อโปรแกรมทำงานเสร็จสิ้นหรือเกิดข้อผิดพลาด
        if driver:
            driver.quit()
        log("🚪 ปิดเบราว์เซอร์")
        log(f"\n📊 สรุปผล:")
        log(f"✅ ดาวน์โหลดสำเร็จ: {success_count} รายการ")
        log(f"❌ ดาวน์โหลดล้มเหลว: {fail_count} รายการ")
        if failed_reports:
            log("🛑 รายงานที่ล้มเหลว:")
            for r in failed_reports:
                log(f" - {r}")
        log(f"⏰ ใช้เวลาทั้งหมด: {round(time.time() - start_time)} วินาที")
        
        # ส่งอีเมลฉบับเดียวพร้อมแนบไฟล์ทั้งหมด
        if downloaded_files:
            subject = EMAIL_SUBJECT.format(month=datetime.now().strftime('%m'), year=datetime.now().strftime('%Y'))
            body = EMAIL_BODY
            send_email_with_attachments(EMAIL_ADDRESS, EMAIL_PASSWORD, RECIPIENT_EMAIL, subject, body, downloaded_files)

start_download()