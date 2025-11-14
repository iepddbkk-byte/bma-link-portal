import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests 
import datetime

# --- 1. ตั้งค่าการเชื่อมต่อ Google Sheets ---
SCOPE = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
]
CREDS_FILE = 'my-project-12345.json' 
SHEET_KEY = "1z3-cjGsP8EHoVa85rn_O_F9NAkKz0ZCW4L0ybCnmcZM" 

db_sheet = None 

try:
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, SCOPE)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SHEET_KEY)
    db_sheet = spreadsheet.worksheet("Database")
    print("✅ (CHECKER) เชื่อมต่อ Google Sheet 'Database' สำเร็จ!")
except Exception as e:
    print(f"❌ (CHECKER) เกิดข้อผิดพลาดในการเชื่อมต่อ Sheet: {e}")
    exit() 

# ปลอม User-Agent ให้เหมือนเบราว์เซอร์ Chrome ปกติ
REQUEST_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

def check_all_links():
    if db_sheet is None:
        print("❌ (CHECKER) ไม่พบ 'Database' sheet")
        return

    print("🚀 (CHECKER) เริ่มกระบวนการตรวจสอบลิงค์...")
    
    try:
        records = db_sheet.get_all_records()
        if not records:
            print("ℹ️ (CHECKER) ไม่พบข้อมูลลิงค์ใน Sheet")
            return
            
        updates = [] 
        
        for i, record in enumerate(records, start=2):
            url = record.get('URL')
            current_row = i
            
            if not url:
                print(f"⚠️ (CHECKER) แถวที่ {current_row}: ข้ามเนื่องจาก URL ว่างเปล่า")
                continue

            if not url.startswith('http://') and not url.startswith('https://'):
                url = 'http://' + url

            status_message = ""
            try:
                response = requests.get(url, 
                                        headers=REQUEST_HEADERS, 
                                        timeout=5, 
                                        allow_redirects=True)
                
                if 200 <= response.status_code < 300:
                    status_message = "OK"
                
                # (ใหม่!) เพิ่มการดักจับ 403 (Forbidden)
                elif response.status_code == 403:
                    status_message = "403 Blocked" # (แปลว่าเว็บนี้ห้ามบอท)
                
                else:
                    status_message = f"{response.status_code} Error" # เช่น 404, 500

            except requests.exceptions.Timeout:
                status_message = "Timeout" 
            except requests.exceptions.ConnectionError:
                status_message = "Connection Error"
            except requests.exceptions.RequestException:
                status_message = "URL Error"

            print(f"  -> แถวที่ {current_row}: {url} | ผลลัพธ์: {status_message}")
            updates.append({
                'range': f'L{current_row}', # คอลัมน์ L (LinkStatus)
                'values': [[status_message]]
            })

        if updates:
            print("\n...กำลังบันทึกผลลัพธ์ลง Google Sheet...")
            db_sheet.batch_update(updates, value_input_option='RAW')
            
        print(f"\n🏁 (CHECKER) ตรวจสอบลิงค์ทั้งหมด {len(records)} รายการ สำเร็จ!")
        print(f"🕘 (CHECKER) เวลาสิ้นสุด: {datetime.datetime.now()}")

    except Exception as e:
        print(f"❌ (CHECKER) เกิดข้อผิดพลาดร้ายแรงระหว่างทำงาน: {e}")

# --- รันฟังก์ชันหลัก ---
if __name__ == "__main__":
    check_all_links()