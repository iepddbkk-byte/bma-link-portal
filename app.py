import gspread
import qrcode
from datetime import timedelta
from io import BytesIO
from flask import (
    Flask, jsonify, request, render_template, 
    session, redirect, url_for, flash, send_file
)
from oauth2client.service_account import ServiceAccountCredentials
from werkzeug.security import check_password_hash, generate_password_hash
import datetime
import random
import string
import os
import json
import requests
import uuid 
from collections import Counter
from threading import Thread 
from PIL import Image
from flask_mail import Mail, Message
from itsdangerous import URLSafeTimedSerializer, SignatureExpired, BadTimeSignature

# --- 1. ตั้งค่า Flask App ---
app = Flask(__name__)
app.config['JSON_AS_ASCII'] = False
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', os.urandom(24))

# --- 2. ตั้งค่า FLASK-MAIL ---
app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USERNAME'] = os.environ.get('MAIL_USERNAME', 'iepdd.bkk@gmail.com') 
app.config['MAIL_PASSWORD'] = os.environ.get('MAIL_PASSWORD', 'gemq daaw qbbq xfts')  
app.config['MAIL_DEFAULT_SENDER'] = ('BMA Link Registry', app.config['MAIL_USERNAME'])

mail = Mail(app)
s = URLSafeTimedSerializer(app.config['SECRET_KEY'])

# --- 3. ตั้งค่าการเชื่อมต่อ Google Sheets ---
SCOPE = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
]
CREDS_FILE = 'my-project-12345.json' 
SHEET_KEY = "1z3-cjGsP8EHoVa85rn_O_F9NAkKz0ZCW4L0ybCnmcZM" 

db_sheet = None 
staff_sheet = None 
invite_sheet = None 
feedback_sheet = None
stats_sheet = None 
favorites_sheet = None # [NEW] เพิ่มตัวแปรสำหรับ Favorites Sheet

try:
    json_creds = os.environ.get('GOOGLE_CREDENTIALS')
    if json_creds:
        creds_dict = json.loads(json_creds)
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, SCOPE)
    else:
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, SCOPE)

    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SHEET_KEY)
    
    db_sheet = spreadsheet.worksheet("Database")
    staff_sheet = spreadsheet.worksheet("StaffList")
    invite_sheet = spreadsheet.worksheet("InviteCodes")
    feedback_sheet = spreadsheet.worksheet("Feedback")
    
    try:
        stats_sheet = spreadsheet.worksheet("SiteStats")
    except:
        print("⚠️ Warning: SiteStats sheet not found.")

    # [NEW] เชื่อมต่อกับ UserFavorites Sheet
    try:
        favorites_sheet = spreadsheet.worksheet("UserFavorites")
    except:
        print("⚠️ Warning: UserFavorites sheet not found. Please create a new tab named 'UserFavorites'.")

    print("✅ เชื่อมต่อ Google Sheet ครบทุกแท็บสำเร็จ!")

except Exception as e:
    print(f"❌ เกิดข้อผิดพลาดในการเชื่อมต่อ Sheet: {e}")

# --- 3.5. ข้อมูลคงที่ ---
bureaus_list = [
    "สำนักงานเลขานุการสภากรุงเทพมหานคร", "สำนักงานเลขานุการผู้ว่าราชการกรุงเทพมหานคร",
    "สำนักงานคณะกรรมการข้าราชการกรุงเทพมหานคร", "สำนักปลัดกรุงเทพมหานคร", "สำนักการแพทย์",
    "สำนักอนามัย", "สำนักการศึกษา", "สำนักการโยธา", "สำนักการระบายน้ำ", "สำนักการคลัง",
    "สำนักเทศกิจ", "สำนักการจราจรและขนส่ง", "สำนักการวางผังเมืองและพัฒนาเมือง",
    "สำนักป้องกันและบรรเทาสาธารณภัย", "สำนักงบประมาณกรุงเทพมหานคร", "สำนักยุทธศาสตร์และประเมินผล",
    "สำนักสิ่งแวดล้อม", "สำนักวัฒนธรรม กีฬา และการท่องเที่ยว", "สำนักพัฒนาสังคม"
]

districts_list = [
    "สำนักงานเขตคลองเตย", "สำนักงานเขตคลองสาน", "สำนักงานเขตคลองสามวา", "สำนักงานเขตคันนายาว",
    "สำนักงานเขตจตุจักร", "สำนักงานเขตจอมทอง", "สำนักงานเขตดอนเมือง", "สำนักงานเขตดินแดง",
    "สำนักงานเขตดุสิต", "สำนักงานเขตตลิ่งชัน", "สำนักงานเขตทวีวัฒนา", "สำนักงานเขตทุ่งครุ",
    "สำนักงานเขตธนบุรี", "สำนักงานเขตบางกอกน้อย", "สำนักงานเขตบางกอกใหญ่", "สำนักงานเขตบางกะปิ",
    "สำนักงานเขตบางขุนเทียน", "สำนักงานเขตบางเขน", "สำนักงานเขตบางคอแหลม", "สำนักงานเขตบางแค",
    "สำนักงานเขตบางซื่อ", "สำนักงานเขตบางนา", "สำนักงานเขตบางบอน", "สำนักงานเขตบางพลัด",
    "สำนักงานเขตบางรัก", "สำนักงานเขตบึงกุ่ม", "สำนักงานเขตปทุมวัน", "สำนักงานเขตประเวศ",
    "สำนักงานเขตป้อมปราบศัตรูพ่าย", "สำนักงานเขตพญาไท", "สำนักงานเขตพระโขนง", "สำนักงานเขตพระนคร",
    "สำนักงานเขตภาษีเจริญ", "สำนักงานเขตมีนบุรี", "สำนักงานเขตยานนาวา", "สำนักงานเขตราชเทวี",
    "สำนักงานเขตราษฎร์บูรณะ", "สำนักงานเขตลาดกระบัง", "สำนักงานเขตลาดพร้าว", "สำนักงานเขตวังทองหลาง",
    "สำนักงานเขตวัฒนา", "สำนักงานเขตสวนหลวง", "สำนักงานเขตสะพานสูง", "สำนักงานเขตสัมพันธวงศ์",
    "สำนักงานเขตสาทร", "สำนักงานเขตสายไหม", "สำนักงานเขตหนองแขม", "สำนักงานเขตหนองจอก",
    "สำนักงานเขตหลักสี่", "สำนักงานเขตห้วยขวาง"
]

# [UPDATE] เพิ่ม "ความเป็นส่วนตัว" เข้าไปใน Headers (Index 13)
DB_HEADERS = [
    "ID", "ประเภท", "หน่วยงาน", "ส่วนราชการ", "อีเมลผู้รับผิดชอบ", "เบอร์โทรติดต่อ",
    "ชื่อลิงก์", "URL", "สถานะ", "รายละเอียด", "วันที่อัปเดต", "CreatorUsername", "LinkStatus", "ความเป็นส่วนตัว", "Clicks", "EditCount"
]

STAFF_HEADERS = [
    "Username", "PasswordHash", "Level", "ชื่อ", "ตำแหน่ง", 
    "หน่วยงาน", "ส่วนราชการ", "เบอร์โทร", "Email", "CreatedAt", "UpdatedAt"
]

# ==========================================
# ส่วนระบบ Caching
# ==========================================
CACHE_TIMEOUT = 15 
_app_cache = {
    'db_records': {'data': None, 'time': None},
    'staff_records': {'data': None, 'time': None},
    'invite_codes': {'data': None, 'time': None} # [เพิ่ม] เพิ่ม Cache สำหรับ Invite Code
}

def clear_db_cache():
    global _app_cache
    _app_cache['db_records'] = {'data': None, 'time': None}
    print("🧹 DB Cache Cleared")

def clear_staff_cache():
    global _app_cache
    _app_cache['staff_records'] = {'data': None, 'time': None}
    print("🧹 Staff Cache Cleared")

# --- 4. ฟังก์ชันตัวช่วย ---
def generate_new_id():
    code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
    return f"BMA-{code}"

def get_current_timestamp():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def generate_invite_code():
    code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
    return f"INVITE-{code}"

def records_to_dict(records_list, headers):
    records_dict = []
    if not records_list or len(records_list) < 2:
        return []
        
    for row in records_list[1:]:
        if row and len(row) > 0 and row[0].strip(): 
            record = {}
            for i, header in enumerate(headers):
                if i < len(row):
                    value = row[i]
                    if isinstance(value, str):
                        record[header] = value.strip()
                    else:
                        record[header] = value
                else:
                    record[header] = "" 
            records_dict.append(record)
    return records_dict

def get_db_records(force_refresh=False):
    global _app_cache
    if db_sheet is None: return []

    now = datetime.datetime.now()
    cache = _app_cache['db_records']

    if not force_refresh and cache['data'] is not None and cache['time'] is not None:
        if (now - cache['time']).total_seconds() < CACHE_TIMEOUT:
            return cache['data']

    try:
        print("zzz Fetching DB from Google Sheets... (API Call)")
        all_values = db_sheet.get_all_values()
        data = records_to_dict(all_values, DB_HEADERS)
        
        cache['data'] = data
        cache['time'] = now
        return data
    except Exception as e:
        print(f"❌ Error fetching DB: {e}")
        return []

def get_staff_records(force_refresh=False):
    global _app_cache
    if staff_sheet is None: return []

    now = datetime.datetime.now()
    cache = _app_cache['staff_records']

    if not force_refresh and cache['data'] is not None and cache['time'] is not None:
        if (now - cache['time']).total_seconds() < CACHE_TIMEOUT:
            return cache['data']

    try:
        print("zzz Fetching Staff from Google Sheets... (API Call)")
        all_values = staff_sheet.get_all_values()
        data = records_to_dict(all_values, STAFF_HEADERS)
        
        cache['data'] = data
        cache['time'] = now
        return data
    except Exception as e:
        print(f"❌ Error fetching Staff: {e}")
        return []
        
def get_invite_codes(force_refresh=False):
    global _app_cache
    if invite_sheet is None: return []

    now = datetime.datetime.now()
    cache = _app_cache['invite_codes']

    # ถ้ามีข้อมูลใน Cache และยังไม่หมดเวลา ให้ใช้ของเดิม
    if not force_refresh and cache['data'] is not None and cache['time'] is not None:
        if (now - cache['time']).total_seconds() < CACHE_TIMEOUT:
            return cache['data']

    try:
        print("zzz Fetching Invite Codes... (API Call)")
        # ใช้ get_all_records เพื่อให้ได้ list of dicts เลย
        data = invite_sheet.get_all_records()
        
        cache['data'] = data
        cache['time'] = now
        return data
    except Exception as e:
        print(f"❌ Error fetching Invite Codes: {e}")
        return []

def clear_invite_cache():
    global _app_cache
    _app_cache['invite_codes'] = {'data': None, 'time': None}
    print("🧹 Invite Cache Cleared")
    
def send_reset_email(username, recipient_email):
    try:
        token = s.dumps(username, salt='password-reset-salt')
        reset_url = url_for('reset_password_page', token=token, _external=True)
        msg = Message("คำขอรีเซ็ตรหัสผ่าน - BMA SMART LINKAGE", recipients=[recipient_email])
        msg.body = f"""สวัสดีครับ,\nเราได้รับคำขอรีเซ็ตรหัสผ่านสำหรับ Username: {username}\nคลิกลิงค์เพื่อตั้งรหัสผ่านใหม่: {reset_url}\n(ลิงค์หมดอายุใน 1 ชั่วโมง)"""
        mail.send(msg)
        return True
    except Exception as e:
        print(f"Mail Error: {e}")
        return False

def increment_site_views():
    if stats_sheet:
        try:
            current_views = int(stats_sheet.cell(1, 2).value or 0)
            stats_sheet.update_cell(1, 2, current_views + 1)
        except Exception as e:
            print(f"Error incrementing views: {e}")

# [NEW] ฟังก์ชันตรวจสอบสิทธิ์การเข้าถึงลิงค์
def check_link_permission(link, user_session):
    privacy = link.get('ความเป็นส่วนตัว', 'สาธารณะ') # Default เป็นสาธารณะ
    
    # 1. ลิงค์สาธารณะ (ใครก็เห็นได้)
    if privacy == 'สาธารณะ' or not privacy:
        return True
        
    # กรณีอื่นๆ ต้อง Login ก่อน
    if not user_session.get('logged_in'):
        return False
        
    user_role = user_session.get('level', '')
    user_agency = user_session.get('main_agency', '')
    username = user_session.get('username', '')
    
    # Admin เห็นทุกอย่าง
    if user_role == 'Admin':
        return True
        
    # 2. เฉพาะเจ้าหน้าที่ในหน่วยงานเดียวกัน
    if privacy == 'เฉพาะหน่วยงาน':
        return link.get('ส่วนราชการ') == user_agency
        
    # 3. เห็นเฉพาะตนเอง (ส่วนตัว)
    if privacy == 'ส่วนตัว':
        return link.get('CreatorUsername') == username
        
    return False

# --- 5. API ---
@app.route('/check_username', methods=['POST'])
def check_username():
    if staff_sheet is None: return jsonify({'available': False, 'message': 'DB Error'})
    try:
        data = request.get_json()
        username = data.get('username')
        if not username: return jsonify({'available': False})
        all_staff = get_staff_records() 
        all_usernames = [u['Username'] for u in all_staff]
        is_available = username.lower() not in [u.lower() for u in all_usernames]
        return jsonify({'available': is_available})
    except: return jsonify({'available': False})

@app.route('/check_invite_code', methods=['POST'])
def check_invite_code():
    if invite_sheet is None: return jsonify({'available': False, 'message': 'DB Error'})
    try:
        data = request.get_json()
        code = data.get('invite_code')
        cell = invite_sheet.find(code)
        if not cell: return jsonify({'available': False, 'message': 'รหัสไม่ถูกต้อง'})
        status = invite_sheet.cell(cell.row, 2).value 
        if status == 'Available': return jsonify({'available': True, 'message': 'รหัสถูกต้อง'})
        else: return jsonify({'available': False, 'message': 'ถูกใช้งานแล้ว'})
    except: return jsonify({'available': False, 'message': 'Server Error'})

@app.route('/run_link_checker')
def run_link_checker():
    key = request.args.get('key')
    secret_key = os.environ.get('CHECKER_SECRET', 'my_super_secret_checker_key')
    if key != secret_key: return jsonify({'status': 'error', 'message': 'Unauthorized'}), 401
    if db_sheet is None: return jsonify({'status': 'error', 'message': 'Database not connected'}), 500
    try:
        records = get_db_records()
        updates = []
        headers = {'User-Agent': 'Mozilla/5.0'}
        for i, record in enumerate(records[:30]): 
            row_index = i + 2 
            url = record.get('URL')
            status_msg = "Unknown"
            if url and url.startswith('http'):
                try:
                    resp = requests.get(url, headers=headers, timeout=3)
                    status_msg = "OK" if 200 <= resp.status_code < 300 or resp.status_code == 403 else f"{resp.status_code} Error"
                except:
                    status_msg = "Error/Timeout"
            # อัปเดตที่คอลัมน์ M (LinkStatus)
            updates.append({'range': f'M{row_index}', 'values': [[status_msg]]})
        if updates:
            db_sheet.batch_update(updates, value_input_option='RAW')
            clear_db_cache()
            return jsonify({'status': 'success', 'message': f'Checked {len(updates)} links'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500
    return jsonify({'status': 'success', 'message': 'No links checked'})

# --- QR CODE ---
@app.route('/generate_qr')
def generate_qr_code():
    url_to_encode = request.args.get('url')
    if not url_to_encode: return "No URL provided", 400
    try:
        qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=10, border=4)
        qr.add_data(url_to_encode)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white").convert('RGBA')
        
        logo_path = os.path.join(app.root_path, 'pictures', 'qr_logo.png')
        if os.path.exists(logo_path):
            logo = Image.open(logo_path).convert('RGBA')
            qr_width, qr_height = img.size
            logo_size = int(qr_width * 0.25) 
            logo = logo.resize((logo_size, logo_size), Image.LANCZOS)
            pos = ((qr_width - logo_size) // 2, (qr_height - logo_size) // 2)
            img.paste(logo, pos, logo)
        
        img_io = BytesIO()
        img.save(img_io, 'PNG')
        img_io.seek(0)
        return send_file(img_io, mimetype='image/png', as_attachment=False)
    except Exception as e:
        return f"QR Generation Error: {e}", 500

@app.route('/go/<link_id>')
def go_to_link(link_id):
    try:
        cell = db_sheet.find(link_id)
        if cell:
            row_values = db_sheet.row_values(cell.row)
            target_url = row_values[7] 
            
            def update_click_count(row_num):
                try:
                    # [UPDATE] Clicks ขยับไปเป็นคอลัมน์ที่ 15 (O) เพราะเพิ่ม Privacy เข้ามาแทรก
                    current_clicks = db_sheet.cell(row_num, 15).value
                    if not current_clicks: current_clicks = 0
                    db_sheet.update_cell(row_num, 15, int(current_clicks) + 1)
                except: pass
            Thread(target=update_click_count, args=(cell.row,)).start()

            if target_url: return redirect(target_url)
        return render_template('error.html', message="ไม่พบลิงค์ที่คุณค้นหา หรือลิงค์ถูกลบไปแล้ว") 
    except Exception as e:
        print(f"[ERROR] Go-to-link error: {e}")
        return render_template('error.html', message="เกิดข้อผิดพลาดในการเชื่อมต่อข้อมูล") 

# --- 6. Routes (General) ---
@app.route('/')
def home():
    if db_sheet is None: 
        return render_template('index.html', session=session, bureaus=bureaus_list, districts=districts_list, links=[], error="Sheet Error")
    try:
        if not session.get('visited_home'):
            Thread(target=increment_site_views).start()
            session['visited_home'] = True
        
        if not session.get('visitor_id'):
            session['visitor_id'] = str(uuid.uuid4())

        total_views = 0
        if stats_sheet:
            try: total_views = stats_sheet.cell(1, 2).value or 0
            except: pass

        all_records = get_db_records()
        
        # [UPDATE] กรองลิงค์ตามสิทธิ์การเข้าถึง (Privacy)
        visible_links = [link for link in all_records if link.get('สถานะ') == 'ใช้งาน' and check_link_permission(link, session)]
        
        for l in visible_links:
            try: l['Clicks'] = int(l.get('Clicks', 0))
            except: l['Clicks'] = 0
        
        top_5_links = sorted(visible_links, key=lambda x: x['Clicks'], reverse=True)[:5]

        for link in visible_links:
            date_str = link.get('วันที่อัปเดต')
            link['วันที่อัปเดต_สั้น'] = date_str.split(' ')[0] if date_str else ''

        return render_template('index.html', 
                               session=session,
                               bureaus=bureaus_list,
                               districts=districts_list,
                               links=visible_links, 
                               top_links=top_5_links, 
                               total_views=total_views, 
                               error=None)
    except Exception as e: 
        return render_template('index.html', session=session, bureaus=bureaus_list, districts=districts_list, links=[], error=str(e))

@app.route('/links')
def links_page():
    if db_sheet is None: 
        return render_template('links_page.html', links=[], error="Sheet Error", session=session, agency_name="Error")
    try:
        if not session.get('visitor_id'):
            session['visitor_id'] = str(uuid.uuid4())

        agency_filter = request.args.get('agency') 
        category_filter = request.args.get('category')
        all_records = get_db_records()
        
        # [UPDATE] กรองลิงค์ตามสิทธิ์การเข้าถึง (Privacy) ก่อนกรองหมวดหมู่
        valid_links = [l for l in all_records if l.get('สถานะ') == 'ใช้งาน' and check_link_permission(l, session)]

        if agency_filter:
            links_to_display = [l for l in valid_links if l.get('ส่วนราชการ') == agency_filter]
            page_title = agency_filter
        elif category_filter:
            links_to_display = [l for l in valid_links if l.get('ประเภท') == category_filter]
            page_title = category_filter
        else:
            links_to_display = valid_links
            page_title = "ลิงค์ทั้งหมด"
        
        for link in links_to_display:
            date_str = link.get('วันที่อัปเดต')
            link['วันที่อัปเดต_สั้น'] = date_str.split(' ')[0] if date_str else ''

        # [NEW] ดึงรายการ ID ที่ผู้ใช้กด Favorite ไว้ (เพื่อไปแสดงสถานะปุ่มดาว)
        my_fav_ids = []
        if session.get('logged_in'):
            # เรียกฟังก์ชัน get_user_favorite_ids ที่เราเพิ่มไปก่อนหน้านี้
            my_fav_ids = get_user_favorite_ids(session.get('username'))

        return render_template('links_page.html', 
                               links=links_to_display, 
                               error=None, 
                               session=session, 
                               agency_name=page_title,
                               fav_ids=my_fav_ids, # [NEW] ส่งรายการ ID ที่ชอบไปที่ Template
                               bureaus=bureaus_list, 
                               districts=districts_list)
    except Exception as e: 
        return render_template('links_page.html', links=[], error=str(e), session=session, agency_name="Error")
        
@app.route('/my_favorites')
def my_favorites_page():
    # 1. บังคับ Login
    if not session.get('logged_in'): 
        return redirect(url_for('login_page'))
    
    if db_sheet is None: 
        return render_template('favorites.html', links=[], error="Database Error", session=session)

    try:
        # 2. ดึงข้อมูลทั้งหมด
        all_links = get_db_records()
        
        # 3. ดึง ID รายการโปรดของผู้ใช้
        my_fav_ids = get_user_favorite_ids(session.get('username'))
        
        # 4. กรองเฉพาะลิงก์ที่เป็น Favorite และสถานะใช้งานได้ (และ User มีสิทธิ์เห็น)
        fav_links = []
        for l in all_links:
            if l.get('ID') in my_fav_ids and l.get('สถานะ') == 'ใช้งาน' and check_link_permission(l, session):
                # จัดการวันที่ให้สวยงาม
                date_str = l.get('วันที่อัปเดต')
                l['วันที่อัปเดต_สั้น'] = date_str.split(' ')[0] if date_str else ''
                fav_links.append(l)

        # ส่งข้อมูลไปที่หน้า favorites.html
        return render_template('favorites.html', 
                               links=fav_links, 
                               session=session,
                               fav_ids=my_fav_ids, # ส่งไปเพื่อแสดงปุ่มดาวสีเหลือง
                               bureaus=bureaus_list, 
                               districts=districts_list)
                               
    except Exception as e:
        print(f"Favorites Page Error: {e}")
        return redirect(url_for('dashboard'))
        

# --- 7. Routes (Auth) ---
@app.route('/login')
def login_page():
    if session.get('logged_in'): return redirect(url_for('dashboard'))
    return render_template('login.html') 

@app.route('/login_action', methods=['POST'])
def login_action():
    if staff_sheet is None: return redirect(url_for('login_page'))
    try:
        username = request.form.get('username')
        password = request.form.get('password')
        staff_list = get_staff_records()
        user_found = next((u for u in staff_list if u['Username'].lower() == username.lower()), None)
        
        if user_found and check_password_hash(user_found['PasswordHash'], password):
            session['logged_in'] = True
            session['username'] = user_found['Username']
            session['level'] = user_found['Level']
            session['name'] = user_found['ชื่อ']
            session['email'] = user_found['Email']
            session['main_agency'] = user_found.get('ส่วนราชการ', 'N/A') 
            flash('เข้าสู่ระบบสำเร็จ!', 'success') 
            return redirect(url_for('dashboard'))
        else:
            flash('Username หรือ รหัสผ่าน ไม่ถูกต้อง', 'error')
            return redirect(url_for('login_page'))
    except Exception as e:
        flash(f'Error: {e}', 'error')
        return redirect(url_for('login_page'))

@app.route('/logout')
def logout():
    session.clear() 
    flash('ออกจากระบบแล้ว', 'info')
    return redirect(url_for('home'))

@app.route('/register')
def register_page():
    if session.get('logged_in'): return redirect(url_for('dashboard'))
    return render_template('register.html') 

@app.route('/register_action', methods=['POST'])
def register_action():
    if staff_sheet is None: return redirect(url_for('register_page'))
    try:
        username = request.form.get('username')
        password = request.form.get('password')
        fullname = request.form.get('fullname')
        email = request.form.get('email')
        phone = request.form.get('phone')
        if phone and not phone.startswith("'"): phone = "'" + phone
        invite_code = request.form.get('invite_code')
        position = request.form.get('position') 
        division = request.form.get('division') 
        main_agency = request.form.get('main_agency') 

        all_staff = get_staff_records()
        if username.lower() in [u['Username'].lower() for u in all_staff]:
            flash('Username ซ้ำ', 'error'); return redirect(url_for('register_page'))
        
        code_cell = invite_sheet.find(invite_code)
        if not code_cell or invite_sheet.cell(code_cell.row, 2).value != 'Available':
            flash('รหัสเชิญไม่ถูกต้อง', 'error'); return redirect(url_for('register_page'))

        hashed_password = generate_password_hash(password)
        current_time = get_current_timestamp()
        
        new_row_final = [username, hashed_password, 'Users', fullname, position, division, main_agency, phone, email, current_time, current_time]
        staff_sheet.append_row(new_row_final, value_input_option='USER_ENTERED')
        
        clear_staff_cache()
        
        invite_sheet.update_cell(code_cell.row, 2, "Used")
        invite_sheet.update_cell(code_cell.row, 3, username)
        invite_sheet.update_cell(code_cell.row, 4, current_time)
        
        flash('สมัครสมาชิกสำเร็จ!', 'success')
        return redirect(url_for('login_page')) 
    except Exception as e:
        print(f"Register Error: {e}")
        flash('เกิดข้อผิดพลาด', 'error')
        return redirect(url_for('register_page'))

@app.route('/forgot_password', methods=['GET', 'POST'])
def forgot_password():
    if request.method == 'POST':
        try:
            username_input = request.form.get('username')
            all_staff = get_staff_records() 
            user_found = next((u for u in all_staff if u['Username'].lower() == username_input.lower()), None)
            if user_found: send_reset_email(user_found['Username'], user_found['Email'])
            flash('หากพบข้อมูล ระบบได้ส่งอีเมลไปแล้ว', 'info')
            return redirect(url_for('login_page'))
        except Exception as e:
            flash('Error sending email', 'error')
            return redirect(url_for('forgot_password'))
    return render_template('forgot_password.html')

@app.route('/reset_password/<token>', methods=['GET', 'POST'])
def reset_password_page(token):
    try: username = s.loads(token, salt='password-reset-salt', max_age=3600)
    except: flash('ลิงค์หมดอายุหรือผิดพลาด', 'error'); return redirect(url_for('forgot_password'))
    
    if request.method == 'POST':
        password = request.form.get('password')
        confirm = request.form.get('confirm_password')
        if password != confirm:
            flash('รหัสผ่านไม่ตรงกัน', 'error'); return render_template('reset_password.html', token=token)
        try:
            new_hash = generate_password_hash(password)
            cell = staff_sheet.find(username)
            if cell:
                staff_sheet.update_cell(cell.row, 2, new_hash) 
                staff_sheet.update_cell(cell.row, 11, get_current_timestamp())
                
                clear_staff_cache()
                
                flash('เปลี่ยนรหัสผ่านสำเร็จ', 'success')
                return redirect(url_for('login_page'))
        except: flash('Error updating password', 'error')
    return render_template('reset_password.html', token=token)

# --- 8. Routes (Profile & Edit) ---
@app.route('/profile')
def profile_page():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    return redirect(url_for('view_profile', username=session.get('username')))

@app.route('/view_profile/<username>')
def view_profile(username):
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        all_staff = get_staff_records() 
        user_info = next((u for u in all_staff if u['Username'] == username), None)
        if not user_info: flash(f'ไม่พบผู้ใช้: {username}', 'error'); return redirect(url_for('dashboard'))
        all_links = get_db_records() 
        count = sum(1 for link in all_links if link.get('CreatorUsername') == username)
        is_own = (username == session.get('username'))
        return render_template('profile.html', session=session, user=user_info, links_count=count, is_own_profile=is_own)
    except: return redirect(url_for('dashboard'))

@app.route('/edit_profile')
def edit_profile_page():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        all_staff = get_staff_records() 
        user_info = next((u for u in all_staff if u['Username'] == session.get('username')), None)
        return render_template('edit_profile.html', session=session, user=user_info)
    except: return redirect(url_for('profile_page'))

@app.route('/edit_profile_action', methods=['POST'])
def edit_profile_action():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        fullname = request.form.get('fullname')
        email = request.form.get('email')
        phone = request.form.get('phone')
        if phone and not phone.startswith("'"): phone = "'" + phone
        position = request.form.get('position')
        division = request.form.get('division') 
        main_agency = request.form.get('main_agency') 
        
        cell = staff_sheet.find(session.get('username'))
        if cell:
            staff_sheet.update_cell(cell.row, 4, fullname)    
            staff_sheet.update_cell(cell.row, 5, position)    
            staff_sheet.update_cell(cell.row, 6, division)    
            staff_sheet.update_cell(cell.row, 7, main_agency) 
            staff_sheet.update_cell(cell.row, 8, phone)       
            staff_sheet.update_cell(cell.row, 9, email)       
            staff_sheet.update_cell(cell.row, 11, get_current_timestamp()) 
            
            clear_staff_cache()
            
            session['name'] = fullname
            session['email'] = email
            session['main_agency'] = main_agency
            flash('บันทึกเรียบร้อย', 'success')
            return redirect(url_for('profile_page'))
    except Exception as e:
        flash(f'Error: {e}', 'error')
        return redirect(url_for('edit_profile_page'))

# --- 9. Routes (Dashboard & Links) ---
@app.route('/dashboard')
def dashboard():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    if db_sheet is None: return render_template('dashboard.html', session=session, links=[])
    try: 
        if not session.get('visitor_id'):
            session['visitor_id'] = str(uuid.uuid4())

        all_links = get_db_records()
        
        # [NEW] ดึงรายการ ID ที่ผู้ใช้กด Favorite ไว้
        my_fav_ids = get_user_favorite_ids(session.get('username'))
        
        # Admin เห็นทั้งหมด, User เห็นเฉพาะของตัวเอง
        user_links = []
        if session.get('level') == 'Admin':
            user_links = all_links
        else:
            user_links = [l for l in all_links if l.get('CreatorUsername') == session.get('username')]
        
        # [NEW] กรองเฉพาะรายการโปรด (Link ID ต้องอยู่ใน my_fav_ids และสถานะต้องใช้งานได้)
        # ใช้ list comprehension เพื่อสร้างลิสต์ object ของลิงก์โปรด
        favorite_links_obj = [l for l in all_links if l.get('ID') in my_fav_ids and l.get('สถานะ') == 'ใช้งาน']

        for l in user_links:
            try: l['Clicks'] = int(l.get('Clicks', 0))
            except: l['Clicks'] = 0
            date_str = l.get('วันที่อัปเดต')
            l['วันที่อัปเดต_สั้น'] = date_str.split(' ')[0] if date_str else ''
            
        top_10 = sorted(user_links, key=lambda x: x['Clicks'], reverse=True)[:10]
        
        total_views = 0
        if stats_sheet:
            try: total_views = int(stats_sheet.cell(1, 2).value or 0)
            except: pass
            
        chart_data = {
            "total_views": total_views,
            "top_10_links": top_10,
            "total_links": len(user_links),
            "active_links": sum(1 for l in user_links if l.get('สถานะ') == 'ใช้งาน'),
            "total_users": len(get_staff_records()) 
        }

        return render_template('dashboard.html', 
                               session=session, 
                               links=user_links, 
                               favorite_links=favorite_links_obj, # [NEW] ส่งรายการโปรด
                               fav_ids=my_fav_ids, # [NEW] ส่ง ID เพื่อเช็คสถานะปุ่มดาว
                               bureaus=bureaus_list, 
                               districts=districts_list,
                               chart_data=chart_data,
                               total_views=total_views,
                               top_10_links=top_10)  
    except Exception as e: 
        print(f"Dashboard Error: {e}")
        return redirect(url_for('home'))

@app.route('/add')
def add_link_page():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    user_main_agency = session.get('main_agency', 'N/A')
    return render_template('add_link.html', session=session, locked_agency=user_main_agency)

@app.route('/add_action', methods=['POST'])
def add_link_action():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        # [UPDATE] รับค่า "ความเป็นส่วนตัว" จากฟอร์ม
        data = {k: request.form.get(k) for k in ['ประเภท', 'หน่วยงาน', 'อีเมลผู้รับผิดชอบ', 'เบอร์โทรติดต่อ', 'ชื่อลิงก์', 'URL', 'รายละเอียด', 'สถานะ', 'ความเป็นส่วนตัว']}
        if data['เบอร์โทรติดต่อ'] and not data['เบอร์โทรติดต่อ'].startswith("'"): data['เบอร์โทรติดต่อ'] = "'" + data['เบอร์โทรติดต่อ']
        main_agency = session.get('main_agency', 'N/A') 
        
        # [UPDATE] บันทึกลง Sheet (เพิ่ม privacy ก่อน clicks)
        new_row = [
            generate_new_id(), data['ประเภท'], data['หน่วยงาน'], main_agency, 
            data['อีเมลผู้รับผิดชอบ'], data['เบอร์โทรติดต่อ'], data['ชื่อลิงก์'], 
            data['URL'], data['สถานะ'], data['รายละเอียด'], 
            get_current_timestamp(), session.get('username'), '', 
            data.get('ความเป็นส่วนตัว', 'สาธารณะ'), 0, 
            0  # <--- [NEW] EditCount เริ่มต้นเป็น 0
        ]
        db_sheet.append_row(new_row, value_input_option='USER_ENTERED')
        
        clear_db_cache()

        flash('เพิ่มลิงค์สำเร็จ', 'success'); return redirect(url_for('links_page'))
    except Exception as e: 
        flash(f'Error: {e}', 'error'); return redirect(url_for('add_link_page'))

@app.route('/delete/<link_id>', methods=['POST'])
def delete_link_action(link_id):
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        cell = db_sheet.find(link_id)
        if not cell: return redirect(url_for('dashboard'))
        row_data = db_sheet.row_values(cell.row)
        creator = row_data[11].strip() 
        if session.get('level', '').strip() == 'Admin' or creator.strip() == session.get('username', '').strip():
            db_sheet.delete_rows(cell.row)
            clear_db_cache()
            flash('ลบสำเร็จ', 'success')
        else:
            flash('ไม่มีสิทธิ์', 'error')
        return redirect(url_for('dashboard'))
    except: return redirect(url_for('dashboard'))

@app.route('/edit/<link_id>')
def edit_link_page(link_id):
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        all_links = get_db_records()
        link = next((l for l in all_links if l.get('ID', '').strip() == link_id.strip()), None)
        if not link: 
            flash(f'ไม่พบลิงค์ ID: {link_id}', 'error')
            return redirect(url_for('dashboard'))
        creator = link.get('CreatorUsername', '').strip()
        username = session.get('username', '').strip()
        is_admin = (session.get('level', '').strip() == 'Admin')
        if not is_admin and creator != username:
            flash('คุณไม่มีสิทธิ์แก้ไขลิงก์นี้', 'error')
            return redirect(url_for('dashboard'))
        return render_template('edit_link.html', session=session, link=link, locked_agency=link.get('ส่วนราชการ', 'N/A'), is_admin=is_admin)
    except Exception as e: 
        return redirect(url_for('dashboard'))

@app.route('/update_action/<link_id>', methods=['POST'])
def update_link_action(link_id):
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        cell = db_sheet.find(link_id)
        if not cell: return redirect(url_for('dashboard'))
        row_vals = db_sheet.row_values(cell.row)
        
        # ... (โค้ดตรวจสอบสิทธิ์เหมือนเดิม) ...
        creator = row_vals[11].strip()
        if session.get('level', '').strip() != 'Admin' and creator != session.get('username', '').strip():
             flash('ไม่มีสิทธิ์', 'error'); return redirect(url_for('dashboard'))

        data = {k: request.form.get(k) for k in ['ประเภท', 'หน่วยงาน', 'อีเมลผู้รับผิดชอบ', 'เบอร์โทรติดต่อ', 'ชื่อลิงก์', 'URL', 'สถานะ', 'รายละเอียด', 'ความเป็นส่วนตัว']}
        if data['เบอร์โทรติดต่อ'] and not data['เบอร์โทรติดต่อ'].startswith("'"): data['เบอร์โทรติดต่อ'] = "'" + data['เบอร์โทรติดต่อ']
        main_agency = request.form.get('ส่วนราชการ') if session.get('level', '').strip() == 'Admin' else row_vals[3]
        
        # [NEW LOGIC] คำนวณยอด Clicks และ EditCount
        current_clicks = row_vals[14] if len(row_vals) > 14 else 0
        
        # ดึงค่า EditCount เดิมมาบวก 1
        try:
            current_edits = int(row_vals[15]) if len(row_vals) > 15 and row_vals[15] else 0
        except:
            current_edits = 0
        
        new_edit_count = current_edits + 1

        new_vals = [
            link_id, data['ประเภท'], data['หน่วยงาน'], main_agency, 
            data['อีเมลผู้รับผิดชอบ'], data['เบอร์โทรติดต่อ'], data['ชื่อลิงก์'], 
            data['URL'], data['สถานะ'], data['รายละเอียด'], 
            get_current_timestamp(), creator, row_vals[12] if len(row_vals) > 12 else '',
            data.get('ความเป็นส่วนตัว', 'สาธารณะ'),
            current_clicks,
            new_edit_count # <--- [NEW] บันทึกค่าใหม่ลงไป
        ]
        
        # อัปเดตช่วง A ถึง P (Column 1 ถึง 16)
        range_name = f"A{cell.row}:P{cell.row}" 
        db_sheet.update(range_name, [new_vals])
        
        clear_db_cache()

        flash(f'แก้ไขสำเร็จ (แก้ไขไปแล้ว {new_edit_count} ครั้ง)', 'success')
        return redirect(url_for('dashboard')) 
    except Exception as e: 
        print(e)
        return redirect(url_for('dashboard'))
        
@app.route('/analytics')
def analytics_page():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    if db_sheet is None: return render_template('analytics.html', session=session, chart_data={})
    try:
        links = get_db_records()
        users = get_staff_records()
        feedback = feedback_sheet.get_all_records()
        
        # แปลงค่าตัวเลขให้พร้อมคำนวณ
        for l in links:
            try: l['Clicks'] = int(l.get('Clicks', 0))
            except: l['Clicks'] = 0
            try: l['EditCount'] = int(l.get('EditCount', 0))
            except: l['EditCount'] = 0

        # [NEW] แยก Top 10 เป็น 2 ชุด
        top_10_clicks = sorted(links, key=lambda x: x['Clicks'], reverse=True)[:10]
        top_10_edits = sorted(links, key=lambda x: x['EditCount'], reverse=True)[:10]

        total_views = 0
        if stats_sheet:
            try: total_views = int(stats_sheet.cell(1, 2).value or 0)
            except: pass

        cat_counts = Counter(l['ประเภท'] for l in links if l.get('ประเภท'))
        dept_counts = Counter(l['ส่วนราชการ'] for l in links if l.get('ส่วนราชการ')).most_common(5)
        
        monthly = {}
        for l in links:
            if l.get('วันที่อัปเดต'):
                try: 
                    m = datetime.datetime.strptime(l['วันที่อัปเดต'].split()[0], '%Y-%m-%d').strftime('%Y-%m')
                    monthly[m] = monthly.get(m, 0) + 1
                except: pass
        sorted_m = sorted(monthly.items())
        
        sat, ease, comments, features = [], [], [], []
        for f in feedback:
            try: sat.append(int(f['SatisfactionScore']))
            except: pass
            try: ease.append(int(f['EaseOfUseScore']))
            except: pass
            if f.get('Comments'): comments.append({'user': f['Username'], 'text': f['Comments']})
            if f.get('FeatureRequest'): features.append({'user': f['Username'], 'text': f['FeatureRequest']})
        
        chart_data = {
            "total_views": total_views,
            "top_10_clicks": top_10_clicks, # [NEW] ส่งตัวแปร 1
            "top_10_edits": top_10_edits,   # [NEW] ส่งตัวแปร 2
            "total_links": len(links), "total_users": len(users), 
            "active_links": sum(1 for l in links if l.get('สถานะ') == 'ใช้งาน'),
            "total_responses": len(feedback),
            "category_labels": list(cat_counts.keys()), "category_data": list(cat_counts.values()),
            "dept_labels": [d[0] for d in dept_counts], "dept_data": [d[1] for d in dept_counts],
            "month_labels": [m[0] for m in sorted_m], "month_data": [m[1] for m in sorted_m],
            "avg_sat": round(sum(sat)/len(sat), 1) if sat else 0,
            "avg_ease": round(sum(ease)/len(ease), 1) if ease else 0,
            "recent_comments": comments[-5:][::-1], "recent_features": features[-5:][::-1]
        }
        return render_template('analytics.html', session=session, chart_data=chart_data)
    except Exception as e:
        print(f"Analytics Error: {e}")
        return redirect(url_for('dashboard'))

@app.route('/admin')
def admin_panel():
    if not session.get('logged_in') or session.get('level', '').strip() != 'Admin': 
        return redirect(url_for('dashboard'))
    try:
        # [แก้ไข] เปลี่ยน invite_sheet.get_all_records() เป็น get_invite_codes()
        return render_template('admin_panel.html', 
                             session=session, 
                             staff_list=get_staff_records(), 
                             invite_codes=get_invite_codes()) 
    except Exception as e:
        return f"Error loading admin panel: {e} <br><a href='/dashboard'>Back</a>"

@app.route('/admin/change_level', methods=['POST'])
def change_user_level():
    if not session.get('logged_in') or session.get('level', '').strip() != 'Admin': return redirect(url_for('login_page'))
    try:
        user, level = request.form.get('username'), request.form.get('level')
        if user == session.get('username'): flash('เปลี่ยนระดับตัวเองไม่ได้', 'error'); return redirect(url_for('admin_panel'))
        cell = staff_sheet.find(user)
        if cell: 
            staff_sheet.update_cell(cell.row, 3, level)
            clear_staff_cache()
            flash('สำเร็จ', 'success')
    except: flash('Error', 'error')
    return redirect(url_for('admin_panel'))

@app.route('/admin/delete_user', methods=['POST'])
def delete_user():
    if not session.get('logged_in') or session.get('level', '').strip() != 'Admin': return redirect(url_for('login_page'))
    try:
        user = request.form.get('username')
        if user == session.get('username'): flash('ลบตัวเองไม่ได้', 'error'); return redirect(url_for('admin_panel'))
        cell = staff_sheet.find(user)
        if cell: 
            staff_sheet.delete_rows(cell.row)
            clear_staff_cache()
            flash('สำเร็จ', 'success')
    except: flash('Error', 'error')
    return redirect(url_for('admin_panel'))

@app.route('/admin/generate_code', methods=['POST'])
def generate_code():
    # ตรวจสอบสิทธิ์ Admin
    if not session.get('logged_in') or session.get('level', '').strip() != 'Admin': 
        return redirect(url_for('login_page'))
    
    try:
        # 1. รับค่าจำนวนจากฟอร์ม (ถ้ามี error ให้เป็น 1)
        try:
            amount = int(request.form.get('amount', 1))
        except ValueError:
            amount = 1
            
        # 2. ตรวจสอบความถูกต้อง (ขั้นต่ำ 1, สูงสุด 50 เพื่อป้องกัน Server ค้าง)
        if amount < 1: amount = 1
        if amount > 50: amount = 50

        # 3. เตรียมข้อมูลสำหรับบันทึกหลายแถว
        new_rows = []
        for _ in range(amount):
            code = generate_invite_code()
            # รูปแบบข้อมูล: [Code, Status, UsedBy, UsedAt]
            new_rows.append([code, 'Available', '', ''])

        # 4. บันทึกลง Google Sheet ทีเดียว (เร็วกว่า loop append_row มาก)
        if invite_sheet:
            invite_sheet.append_rows(new_rows, value_input_option='USER_ENTERED')
            clear_invite_cache()
            flash(f'✅ สร้างรหัสเชิญใหม่จำนวน {amount} รหัส สำเร็จเรียบร้อย!', 'success')
        else:
            flash('ไม่สามารถเชื่อมต่อฐานข้อมูลได้', 'error')

    except Exception as e:
        print(f"Generate Code Error: {e}")
        flash(f'เกิดข้อผิดพลาด: {e}', 'error')
        
    return redirect(url_for('admin_panel'))
    
@app.route('/admin/delete_code', methods=['POST'])
def delete_code():
    if not session.get('logged_in') or session.get('level', '').strip() != 'Admin': return redirect(url_for('login_page'))
    try:
        cell = invite_sheet.find(request.form.get('code'))
        if cell: 
            invite_sheet.delete_rows(cell.row)
            clear_invite_cache() # [เพิ่ม] สั่งล้าง Cache ตรงนี้ด้วย
            flash('ลบสำเร็จ', 'success')
    except: flash('Error', 'error')
    return redirect(url_for('admin_panel'))

@app.route('/feedback')
def feedback_page():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    return render_template('feedback.html', session=session)

@app.route('/feedback_action', methods=['POST'])
def feedback_action():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        row = [get_current_timestamp(), session.get('username'), request.form.get('satisfaction'), request.form.get('ease_of_use'), request.form.get('comments', ''), request.form.get('features', '')]
        feedback_sheet.append_row(row, value_input_option='USER_ENTERED')
        flash('ขอบคุณสำหรับข้อเสนอแนะ', 'success'); return redirect(url_for('dashboard'))
    except: return redirect(url_for('feedback_page'))

@app.route('/get_links', methods=['GET'])
def get_all_links():
    try: return jsonify({"status": "success", "data": get_db_records()})
    except: return jsonify({"status": "error"}), 500

# --- Real-time Online Counter System ---

online_users = {}

def cleanup_online_users():
    global online_users
    now = datetime.datetime.now()
    online_users = {k: v for k, v in online_users.items() if now - v < timedelta(seconds=30)}

@app.route('/heartbeat', methods=['POST'])
def heartbeat():
    global online_users
    if session.get('logged_in'):
        user_key = f"user:{session.get('username')}"
    else:
        if not session.get('visitor_id'):
            session['visitor_id'] = str(uuid.uuid4())
        user_key = f"guest:{session.get('visitor_id')}"
        
    online_users[user_key] = datetime.datetime.now()
    cleanup_online_users()
    return jsonify({'status': 'ok', 'online_count': len(online_users)})

@app.route('/offline', methods=['POST'])
def offline():
    global online_users
    try:
        if session.get('logged_in'):
            user_key = f"user:{session.get('username')}"
        else:
            if session.get('visitor_id'):
                user_key = f"guest:{session.get('visitor_id')}"
            else:
                return '', 204
        if user_key in online_users:
            del online_users[user_key]
    except Exception as e:
        print(f"Offline signal error: {e}")
    return '', 204
    
# ==========================================
# [NEW] My Favorites System Logic
# ==========================================

def get_user_favorite_ids(username):
    """ดึงรายการ LinkID ที่ User คนนี้กด Favorite ไว้"""
    if favorites_sheet is None: return []
    try:
        # เพื่อประสิทธิภาพที่ดี ควรทำ Cache ในอนาคตหากข้อมูลเยอะ
        all_favs = favorites_sheet.get_all_records()
        user_favs = [row['LinkID'] for row in all_favs if row['Username'] == username]
        return user_favs
    except Exception as e:
        print(f"Error getting favorites: {e}")
        return []

@app.route('/toggle_favorite', methods=['POST'])
def toggle_favorite():
    """API สำหรับกดปุ่มดาว (เพิ่ม/ลบ)"""
    if not session.get('logged_in'):
        return jsonify({'status': 'error', 'message': 'Please login first'}), 401
    
    if favorites_sheet is None:
        return jsonify({'status': 'error', 'message': 'Database not connected'}), 500

    try:
        data = request.get_json()
        link_id = data.get('link_id')
        username = session.get('username')
        
        # ดึงข้อมูลทั้งหมดมาตรวจสอบ (Row Index เริ่มที่ 1)
        all_records = favorites_sheet.get_all_values() 
        found_row_index = -1
        
        for i, row in enumerate(all_records):
            if i == 0: continue # ข้าม Header row
            # Column Structure: A=Timestamp, B=Username, C=LinkID
            # Index ใน list: 0, 1, 2
            if len(row) > 2 and row[1] == username and row[2] == link_id:
                found_row_index = i + 1 # +1 สำหรับ gspread index
                break
        
        if found_row_index != -1:
            # เจอ -> ลบออก (Unfavorite)
            favorites_sheet.delete_rows(found_row_index)
            action = 'removed'
        else:
            # ไม่เจอ -> เพิ่มใหม่ (Favorite)
            favorites_sheet.append_row([get_current_timestamp(), username, link_id], value_input_option='USER_ENTERED')
            action = 'added'
            
        return jsonify({'status': 'success', 'action': action})
        
    except Exception as e:
        print(f"Toggle Favorite Error: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 500
        
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))