# === 🛠️ Import Libraries ===
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

# === 🧾 Credentials และ URLs ===
USERNAME = "csoc_reports"
PASSWORD = "csoc@reports"
LOGIN_URL = "http://nmsgov.ntcsoc.net/Orion/Login.aspx"
REPORT_URL = "http://nmsgov.ntcsoc.net/Orion/reports/viewreports.aspx"

# === 📥 กำหนดโฟลเดอร์สำหรับดาวน์โหลด ===
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads", "temp_downloads")
TARGET_DIR = r"C:\Users\1705w\OneDrive\Desktop\Autoreport\Report"

# ✅ กำหนด Path สำหรับ Edge Driver ที่ดาวน์โหลดด้วยตัวเอง
edge_driver_path = r"C:\Users\1705w\Downloads\edgedriver_win32\msedgedriver.exe"

# === 📜 ฟังก์ชันแสดง Log ใน Console ===
def log(text):
    print(text)
    sys.stdout.flush()

# === ▶️ เริ่มดาวน์โหลด ===
def start_download():
    log("▶️ เริ่มต้นโปรแกรมดาวน์โหลดรายงาน")

    start_time = time.time()
    success_count = 0
    fail_count = 0
    failed_reports = []
    driver = None

    try:
        # ✅ ลบและสร้างโฟลเดอร์สำหรับดาวน์โหลดใหม่ทุกครั้งที่รัน
        if os.path.exists(DOWNLOAD_DIR):
            shutil.rmtree(DOWNLOAD_DIR)
            #log(f"📁 ลบโฟลเดอร์เก่า: {DOWNLOAD_DIR}")
        os.makedirs(DOWNLOAD_DIR)
        #log(f"📁 สร้างโฟลเดอร์ชั่วคราว: {DOWNLOAD_DIR}")
        
        if not os.path.exists(TARGET_DIR):
            os.makedirs(TARGET_DIR)
            #log(f"📂 สร้างโฟลเดอร์บันทึก: {TARGET_DIR}")

        # === ⚙️ ตั้งค่า Edge Options ===
        edge_options = EdgeOptions()
        edge_options.add_argument("--start-maximized")
        edge_options.add_argument("--no-sandbox")
        edge_options.add_argument("--disable-web-security")
        edge_options.add_argument("--ignore-certificate-errors")
        edge_options.add_argument("--disable-infobars")
        edge_options.add_argument("--headless=new")
        
        prefs = {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "safebrowsing.enabled": False,
            "safebrowsing.disable_download_protection": True,
            "download.directory_upgrade": True,
        }
        edge_options.add_experimental_option("prefs", prefs)

        # ✅ ลองดาวน์โหลด Driver อัตโนมัติก่อน
        log("🚀 กำลังพยายามดาวน์โหลด Edge Driver อัตโนมัติ...")
        try:
            service = EdgeService(EdgeChromiumDriverManager().install())
            driver = webdriver.Edge(service=service, options=edge_options)
            log("✅ เปิด Edge Browser สำเร็จ (อัตโนมัติ)")
        except Exception as e:
            #log(f"❌ การดาวน์โหลดอัตโนมัติล้มเหลว: {e}")
            # ✅ ถ้าล้มเหลว ให้ใช้ Driver ที่กำหนดเองแทน
            #log(f"🔄 กำลังใช้ Driver จาก Path: {edge_driver_path}")
            service = EdgeService(edge_driver_path)
            driver = webdriver.Edge(service=service, options=edge_options)
            #log("✅ เปิด Edge Browser สำเร็จ (ด้วยตนเอง)")

        driver.implicitly_wait(10)

        # === 🔐 Login ด้วย Selenium ===
        log("🔐 กำลังเข้าสู่ระบบ...")
        driver.get(LOGIN_URL)
        WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.ID, "ctl00_BodyContent_Username")))

        driver.find_element(By.ID, "ctl00_BodyContent_Username").send_keys(USERNAME)
        driver.find_element(By.ID, "ctl00_BodyContent_Password").send_keys(PASSWORD)
        driver.find_element(By.ID, "ctl00_BodyContent_LoginButton").click()
        log("✅ เข้าสู่ระบบสำเร็จ")

        # === 🔄 Loop หน้า Report ===
        driver.get(REPORT_URL)

        while True:
            # ✅ ดึงลิงก์ทั้งหมดในหน้าปัจจุบันแค่ครั้งเดียว
            report_links = WebDriverWait(driver, 300).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, 'Report.aspx?ReportID=')]"))
            )
            if not report_links:
                break
            
            all_links = [(elem.get_attribute("href"), elem.text.strip()) for elem in report_links]
            log(f"📄 พบ {len(all_links)} รายงานในหน้านี้")

            for href, name in all_links:
                retry = 0
                success = False
                while retry < 2 and not success:
                    try:
                        driver.get(href)
                        log(f"🟢 กำลังเปิดรายงาน: {name} (รอบที่ {retry+1})")
                        
                        # รอให้หน้าโหลดรายงานเสร็จ
                        WebDriverWait(driver, 300).until_not(
                            EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Your requested report')]"))
                        )
                        log("✅ รายงานพร้อมแล้ว")

                        # รอปุ่ม Export to Excel พร้อมใช้งาน
                        export_btn = WebDriverWait(driver, 300).until(
                            EC.element_to_be_clickable((By.LINK_TEXT, "Export to Excel"))
                        )
                        export_btn.click()
                        log("📥 คลิก Export เรียบร้อยแล้ว")
                        
                        # === รอไฟล์ดาวน์โหลดเสร็จ ===
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
                            file_ext = os.path.splitext(latest_file)[1]
                            timestamp = datetime.now().strftime("%Y-%m-%d")
                            new_name = f"{timestamp}_{name.replace(' ', '_').replace('/', '-')}{file_ext}"
                            shutil.move(latest_file, os.path.join(TARGET_DIR, new_name))
                            log(f"✅ บันทึกสำเร็จ: {new_name}")
                            success_count += 1
                            success = True
                        else:
                            raise Exception("การดาวน์โหลดล้มเหลวหรือใช้เวลานานเกินไป")
                            
                    except Exception as e:
                        retry += 1
                        log(f"❌ ล้มเหลวรอบที่ {retry}: {name} -> {e}")
                        if retry >= 2:
                            fail_count += 1
                            failed_reports.append(name)
                    finally:
                        # ✅ แก้ไข: กลับมาหน้าหลักของรายงานทุกครั้งหลังจบการดาวน์โหลด
                        driver.get(REPORT_URL)

            # ตรวจสอบปุ่ม Next
            try:
                next_btn = driver.find_element(By.LINK_TEXT, "Next")
                next_btn.click()
                time.sleep(2)
            except:
                break

    except Exception as e:
        log(f"❌ เกิดข้อผิดพลาดที่ไม่คาดคิด: {e}")

    finally:
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

start_download()