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
favorites_sheet = None

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
    
    try: stats_sheet = spreadsheet.worksheet("SiteStats")
    except: print("⚠️ Warning: SiteStats sheet not found.")

    try: favorites_sheet = spreadsheet.worksheet("UserFavorites")
    except: print("⚠️ Warning: UserFavorites sheet not found.")

    print("✅ เชื่อมต่อ Google Sheet ครบทุกแท็บสำเร็จ!")

except Exception as e:
    print(f"❌ เกิดข้อผิดพลาดในการเชื่อมต่อ Sheet: {e}")

# --- 3.5. ข้อมูลคงที่ ---
bureaus_list = [
    "สำนักงานเลขานุการ", 
    "สำนักงานบริหารยุทธศาสตร์",
    "กองพัฒนายุทธศาสตร์เศรษฐกิจเมือง", 
    "กองยุทธศาสตร์สภาพแวดล้อมและความปลอดภัยเมือง", 
    "กองยุทธศาสตร์พัฒนาโครงสร้างพื้นฐานเมือง", 
    "กองยุทธศาสตร์คุณภาพชีวิตเมือง"
]

districts_list = [
    "ฝ่ายบริหารงานทั่วไป (สก.)", "ฝ่ายการเจ้าหน้าที่", "ฝ่ายการคลัง",
    "กลุ่มช่วยนักบริหาร", "ฝ่ายบริหารงานทั่วไป (สบย.)", "ส่วนบริหารยุทธศาสตร์", "กลุ่มงานนโยบายและยุทธศาสตร์", "กลุ่มงานวิจัยและประเมินผล", "กลุ่มงานยุทธศาสตร์การพัฒนาพื้นที่", "กลุ่มงานวิเคราะห์ข้อมูลเมือง", "กลุ่มงานถ่ายโอนและภารกิจพิเศษ",
    "ฝ่ายบริหารงานทั่วไป (กยศ.)", "กลุ่มงานยุทธศาสตร์เศรษฐกิจ", "กลุ่มงานนโยบายเศรษฐกิจเมือง", "กลุ่มงานส่งเสริมเศรษฐกิจเมือง",
    "ฝ่ายบริหารงานทั่วไป (กยส.)", "กลุ่มงานยุทธศาสตร์และประเมินผลด้านสิ่งแวดล้อม", "กลุ่มงานยุทธศาสตร์และประเมินผลด้านความปลอดภัยและจัดระเบียบเมือง",
    "ฝ่ายบริหารงานทั่วไป (กยพ.)", "กลุ่มงานยุทธศาสตร์และประเมินผลด้านโยธาและระบายน้ำ", "กลุ่มงานยุทธศาสตร์และประเมินผลด้านผังเมือง การจราจรและขนส่ง",
    "ฝ่ายบริหารงานทั่วไป (กยม.)", "กลุ่มงานยุทธศาสตร์และประเมินผลด้านการศึกษา พัฒนาสังคม และสร้างสรรค์เมือง", "กลุ่มงานยุทธศาสตร์และประเมินผลด้านสาธารณสุข"
]

DB_HEADERS = [
    "ID", "ประเภท", "หน่วยงาน", "ส่วนราชการ", "อีเมลผู้รับผิดชอบ", "เบอร์โทรติดต่อ",
    "ชื่อลิงก์", "URL", "สถานะ", "รายละเอียด", "วันที่อัปเดต", "CreatorUsername", "LinkStatus", "ความเป็นส่วนตัว", "Clicks", "EditCount", "ปักหมุด", "วันที่สิ้นสุด"
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
    'invite_codes': {'data': None, 'time': None}
}

def clear_db_cache():
    global _app_cache
    _app_cache['db_records'] = {'data': None, 'time': None}

def clear_staff_cache():
    global _app_cache
    _app_cache['staff_records'] = {'data': None, 'time': None}

def clear_invite_cache():
    global _app_cache
    _app_cache['invite_codes'] = {'data': None, 'time': None}

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
    if not records_list or len(records_list) < 2: return []
    for row in records_list[1:]:
        if row and len(row) > 0 and row[0].strip(): 
            record = {}
            for i, header in enumerate(headers):
                record[header] = row[i].strip() if i < len(row) and isinstance(row[i], str) else row[i] if i < len(row) else ""
            records_dict.append(record)
    return records_dict

def get_db_records(force_refresh=False):
    global _app_cache
    if db_sheet is None: return []
    now = datetime.datetime.now()
    cache = _app_cache['db_records']
    if not force_refresh and cache['data'] is not None and cache['time'] is not None:
        if (now - cache['time']).total_seconds() < CACHE_TIMEOUT: return cache['data']
    try:
        data = records_to_dict(db_sheet.get_all_values(), DB_HEADERS)
        cache['data'] = data
        cache['time'] = now
        return data
    except: return []

def get_staff_records(force_refresh=False):
    global _app_cache
    if staff_sheet is None: return []
    now = datetime.datetime.now()
    cache = _app_cache['staff_records']
    if not force_refresh and cache['data'] is not None and cache['time'] is not None:
        if (now - cache['time']).total_seconds() < CACHE_TIMEOUT: return cache['data']
    try:
        data = records_to_dict(staff_sheet.get_all_values(), STAFF_HEADERS)
        cache['data'] = data
        cache['time'] = now
        return data
    except: return []
        
def get_invite_codes(force_refresh=False):
    global _app_cache
    if invite_sheet is None: return []
    now = datetime.datetime.now()
    cache = _app_cache['invite_codes']
    if not force_refresh and cache['data'] is not None and cache['time'] is not None:
        if (now - cache['time']).total_seconds() < CACHE_TIMEOUT: return cache['data']
    try:
        data = invite_sheet.get_all_records()
        cache['data'] = data
        cache['time'] = now
        return data
    except: return []

def send_reset_email(username, recipient_email):
    try:
        token = s.dumps(username, salt='password-reset-salt')
        reset_url = url_for('reset_password_page', token=token, _external=True)
        msg = Message("คำขอรีเซ็ตรหัสผ่าน - BMA SMART LINKAGE", recipients=[recipient_email])
        msg.body = f"""สวัสดีครับ,\nเราได้รับคำขอรีเซ็ตรหัสผ่านสำหรับ Username: {username}\nคลิกลิงก์เพื่อตั้งรหัสผ่านใหม่: {reset_url}\n(ลิงก์หมดอายุใน 1 ชั่วโมง)"""
        mail.send(msg)
        return True
    except: return False

def increment_site_views():
    if stats_sheet:
        try:
            current_views = int(stats_sheet.cell(1, 2).value or 0)
            stats_sheet.update_cell(1, 2, current_views + 1)
        except: pass

# =======================================================
# [UPDATED] สิทธิ์การเข้าถึงลิงก์ (รองรับ 3 Role)
# =======================================================
def check_link_permission(link, user_session):
    privacy = link.get('ความเป็นส่วนตัว', 'สาธารณะ')
    if privacy == 'สาธารณะ' or not privacy: return True
    if not user_session.get('logged_in'): return False
        
    user_role = user_session.get('level', '')
    username = user_session.get('username', '')
    user_bureau = user_session.get('main_agency', '')
    user_division = user_session.get('division', '')
    
    # 1. Super Admin เห็นทุกลิงก์ในระบบ
    if user_role == 'Super Admin': return True
        
    link_creator = link.get('CreatorUsername')
    link_bureau = link.get('ส่วนราชการ')
    link_division = link.get('หน่วยงาน')

    # 2. Admin (กอง) เห็นทุกลิงก์ในกองของตัวเอง แม้คนสร้างจะตั้งค่าส่วนตัว
    if user_role == 'Admin' and link_bureau == user_bureau:
        return True

    # 3. User ทั่วไป (และ Admin ที่ดูลิงก์กองอื่น) เช็คตาม Privacy
    if privacy == 'สยป.': return True
    if privacy == 'กอง/สำนักงาน': return link_bureau == user_bureau
    if privacy == 'กลุ่มงาน/ฝ่าย': return link_division == user_division and link_bureau == user_bureau
    if privacy == 'ส่วนตัว': return link_creator == username
        
    return False

# --- 5. API ---
@app.route('/check_username', methods=['POST'])
def check_username():
    if staff_sheet is None: return jsonify({'available': False})
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
    if invite_sheet is None: return jsonify({'available': False})
    try:
        data = request.get_json()
        code = data.get('invite_code')
        cell = invite_sheet.find(code)
        if not cell: return jsonify({'available': False, 'message': 'รหัสไม่ถูกต้อง'})
        status = invite_sheet.cell(cell.row, 2).value 
        if status == 'Available': return jsonify({'available': True, 'message': 'รหัสถูกต้อง'})
        else: return jsonify({'available': False, 'message': 'ถูกใช้งานแล้ว'})
    except: return jsonify({'available': False})

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
            updates.append({'range': f'M{row_index}', 'values': [[status_msg]]})
        if updates:
            db_sheet.batch_update(updates, value_input_option='RAW')
            clear_db_cache()
            return jsonify({'status': 'success', 'message': f'Checked {len(updates)} links'})
    except Exception as e: return jsonify({'status': 'error', 'message': str(e)}), 500
    return jsonify({'status': 'success'})

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
    except Exception as e: return f"QR Generation Error: {e}", 500

@app.route('/go/<link_id>')
def go_to_link(link_id):
    try:
        cell = db_sheet.find(link_id)
        if cell:
            row_values = db_sheet.row_values(cell.row)
            target_url = row_values[7] 
            def update_click_count(row_num):
                try:
                    current_clicks = db_sheet.cell(row_num, 15).value
                    if not current_clicks: current_clicks = 0
                    db_sheet.update_cell(row_num, 15, int(current_clicks) + 1)
                except: pass
            Thread(target=update_click_count, args=(cell.row,)).start()
            if target_url: return redirect(target_url)
        return render_template('error.html', message="ไม่พบลิงก์ที่คุณค้นหา หรือลิงก์ถูกลบไปแล้ว") 
    except: return render_template('error.html', message="เกิดข้อผิดพลาดในการเชื่อมต่อข้อมูล") 

# --- 6. Routes (General) ---
@app.route('/')
def home():
    if db_sheet is None: return render_template('index.html', session=session, bureaus=bureaus_list, districts=districts_list, links=[], error="Sheet Error")
    try:
        if not session.get('visited_home'):
            Thread(target=increment_site_views).start()
            session['visited_home'] = True
        if not session.get('visitor_id'): session['visitor_id'] = str(uuid.uuid4())

        total_views = 0
        if stats_sheet:
            try: total_views = stats_sheet.cell(1, 2).value or 0
            except: pass

        all_records = get_db_records()
        visible_links = [link for link in all_records if link.get('สถานะ') == 'ใช้งาน' and check_link_permission(link, session)]
        
        for l in visible_links:
            try: l['Clicks'] = int(l.get('Clicks', 0))
            except: l['Clicks'] = 0
        
        top_5_links = sorted(visible_links, key=lambda x: x['Clicks'], reverse=True)[:5]

        for link in visible_links:
            date_str = link.get('วันที่อัปเดต')
            link['วันที่อัปเดต_สั้น'] = date_str.split(' ')[0] if date_str else ''

        return render_template('index.html', session=session, bureaus=bureaus_list, districts=districts_list, links=visible_links, top_links=top_5_links, total_views=total_views, error=None)
    except Exception as e: return render_template('index.html', session=session, error=str(e))

@app.route('/links')
def links_page():
    if db_sheet is None: return render_template('links_page.html', links=[], error="Sheet Error", session=session)
    try:
        if not session.get('visitor_id'): session['visitor_id'] = str(uuid.uuid4())

        search_query = request.args.get('search', '').lower()
        agency_filter = request.args.get('agency') 
        category_filter = request.args.get('category')
        all_records = get_db_records()
        
        valid_links = [l for l in all_records if l.get('สถานะ') == 'ใช้งาน' and check_link_permission(l, session)]

        links_to_display = []
        for l in valid_links:
            if search_query and search_query not in l.get('ชื่อลิงก์', '').lower() and search_query not in l.get('รายละเอียด', '').lower() and search_query not in l.get('หน่วยงาน', '').lower(): continue
            if agency_filter and agency_filter not in (l.get('ส่วนราชการ', ''), l.get('หน่วยงาน', '')): continue
            if category_filter and category_filter != l.get('ประเภท', ''): continue
            links_to_display.append(l)

        if agency_filter: page_title = agency_filter
        elif category_filter: page_title = category_filter
        elif search_query: page_title = f"ผลการค้นหา: '{request.args.get('search')}'"
        else: page_title = "รายชื่อลิงก์ทั้งหมด"
            
        pinned_links, unpinned_links = [], []
        for link in links_to_display:
            date_str = link.get('วันที่อัปเดต')
            link['วันที่อัปเดต_สั้น'] = date_str.split(' ')[0] if date_str else ''
            if link.get('ปักหมุด') == 'ปักหมุด': pinned_links.append(link)
            else: unpinned_links.append(link)
                
        pinned_links.sort(key=lambda x: x.get('วันที่อัปเดต', ''), reverse=True)
        unpinned_links.sort(key=lambda x: x.get('วันที่อัปเดต', ''), reverse=True)
        final_links = pinned_links + unpinned_links

        my_fav_ids = get_user_favorite_ids(session.get('username')) if session.get('logged_in') else []

        return render_template('links_page.html', links=final_links, session=session, agency_name=page_title, fav_ids=my_fav_ids, bureaus=bureaus_list, districts=districts_list)
    except Exception as e: return render_template('links_page.html', links=[], error=str(e), session=session)
        
@app.route('/my_favorites')
def my_favorites_page():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    if db_sheet is None: return render_template('favorites.html', links=[], session=session)
    try:
        all_links = get_db_records()
        my_fav_ids = get_user_favorite_ids(session.get('username'))
        
        fav_links = []
        for l in all_links:
            if l.get('ID') in my_fav_ids and l.get('สถานะ') == 'ใช้งาน' and check_link_permission(l, session):
                date_str = l.get('วันที่อัปเดต')
                l['วันที่อัปเดต_สั้น'] = date_str.split(' ')[0] if date_str else ''
                fav_links.append(l)

        return render_template('favorites.html', links=fav_links, session=session, fav_ids=my_fav_ids, bureaus=bureaus_list, districts=districts_list)
    except: return redirect(url_for('dashboard'))

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
            session['division'] = user_found.get('หน่วยงาน', 'N/A')
            flash('เข้าสู่ระบบสำเร็จ!', 'success') 
            return redirect(url_for('dashboard'))
        else:
            flash('Username หรือ รหัสผ่าน ไม่ถูกต้อง', 'error')
            return redirect(url_for('login_page'))
    except Exception as e: return redirect(url_for('login_page'))

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
        
        new_row = [username, hashed_password, 'Users', fullname, position, division, main_agency, phone, email, current_time, current_time]
        staff_sheet.append_row(new_row, value_input_option='USER_ENTERED')
        clear_staff_cache()
        
        invite_sheet.update_cell(code_cell.row, 2, "Used")
        invite_sheet.update_cell(code_cell.row, 3, username)
        invite_sheet.update_cell(code_cell.row, 4, current_time)
        
        flash('สมัครสมาชิกสำเร็จ!', 'success')
        return redirect(url_for('login_page')) 
    except: flash('เกิดข้อผิดพลาด', 'error'); return redirect(url_for('register_page'))

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
        except: return redirect(url_for('forgot_password'))
    return render_template('forgot_password.html')

@app.route('/reset_password/<token>', methods=['GET', 'POST'])
def reset_password_page(token):
    try: username = s.loads(token, salt='password-reset-salt', max_age=3600)
    except: flash('ลิงก์หมดอายุหรือผิดพลาด', 'error'); return redirect(url_for('forgot_password'))
    
    if request.method == 'POST':
        password = request.form.get('password')
        confirm = request.form.get('confirm_password')
        if password != confirm: flash('รหัสผ่านไม่ตรงกัน', 'error'); return render_template('reset_password.html', token=token)
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
            session['division'] = division
            flash('บันทึกเรียบร้อย', 'success')
            return redirect(url_for('profile_page'))
    except Exception as e: return redirect(url_for('edit_profile_page'))

# --- 9. Routes (Dashboard & Links Management) ---

# =======================================================
# [UPDATED] การดึงข้อมูลใน Dashboard (รองรับ 3 Role)
# =======================================================
@app.route('/dashboard')
def dashboard():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    if db_sheet is None: return render_template('dashboard.html', session=session, links=[])
    try: 
        if not session.get('visitor_id'): session['visitor_id'] = str(uuid.uuid4())

        all_links = get_db_records()
        my_fav_ids = get_user_favorite_ids(session.get('username'))
        
        user_role = session.get('level', '')
        user_bureau = session.get('main_agency', '')
        username = session.get('username', '')
        
        user_links = []
        if user_role == 'Super Admin':
            user_links = all_links
        elif user_role == 'Admin':
            # Admin เห็นและจัดการลิงก์ทุกอันใน 'ส่วนราชการ' (กอง/สำนัก) ของตนเอง
            user_links = [l for l in all_links if l.get('ส่วนราชการ') == user_bureau]
        else:
            # User ทั่วไป เห็นเฉพาะที่ตัวเองสร้าง
            user_links = [l for l in all_links if l.get('CreatorUsername') == username]
        
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
            "total_views": total_views, "top_10_links": top_10, "total_links": len(user_links),
            "active_links": sum(1 for l in user_links if l.get('สถานะ') == 'ใช้งาน'),
            "total_users": len(get_staff_records()) 
        }

        return render_template('dashboard.html', session=session, links=user_links, favorite_links=favorite_links_obj, fav_ids=my_fav_ids, bureaus=bureaus_list, districts=districts_list, chart_data=chart_data, total_views=total_views, top_10_links=top_10)  
    except Exception as e: return redirect(url_for('home'))

@app.route('/add')
def add_link_page():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    user_main_agency = session.get('main_agency', 'N/A')
    user_division = session.get('division', 'N/A')
    return render_template('add_link.html', session=session, locked_agency=user_main_agency, locked_division=user_division)

@app.route('/add_action', methods=['POST'])
def add_link_action():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        data = {k: request.form.get(k) for k in ['ประเภท', 'อีเมลผู้รับผิดชอบ', 'เบอร์โทรติดต่อ', 'ชื่อลิงก์', 'URL', 'รายละเอียด', 'สถานะ', 'ความเป็นส่วนตัว']}
        if data['เบอร์โทรติดต่อ'] and not data['เบอร์โทรติดต่อ'].startswith("'"): data['เบอร์โทรติดต่อ'] = "'" + data['เบอร์โทรติดต่อ']
        
        main_agency = session.get('main_agency', 'N/A') 
        division = session.get('division', 'N/A')
        
        is_pinned = 'ไม่ปักหมุด'
        end_date = ''
        
        # [UPDATED] สิทธิ์การปักหมุดสำหรับ Super Admin และ Admin
        if session.get('level', '').strip() in ['Super Admin', 'Admin']:
            if request.form.get('is_pinned'): is_pinned = 'ปักหมุด'
            end_date = request.form.get('end_date', '')
            
        new_row = [
            generate_new_id(), data['ประเภท'], division, main_agency, 
            data['อีเมลผู้รับผิดชอบ'], data['เบอร์โทรติดต่อ'], data['ชื่อลิงก์'], 
            data['URL'], data['สถานะ'], data['รายละเอียด'], 
            get_current_timestamp(), session.get('username'), '', 
            data.get('ความเป็นส่วนตัว', 'สาธารณะ'), 0, 0, is_pinned, end_date
        ]
        db_sheet.append_row(new_row, value_input_option='USER_ENTERED')
        clear_db_cache()
        flash('เพิ่มลิงก์สำเร็จ', 'success')
        return redirect(url_for('links_page'))
    except Exception as e: return redirect(url_for('add_link_page'))

# =======================================================
# [UPDATED] สิทธิ์การลบลิงก์ (รองรับ 3 Role)
# =======================================================
@app.route('/delete/<link_id>', methods=['POST'])
def delete_link_action(link_id):
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        cell = db_sheet.find(link_id)
        if not cell: return redirect(url_for('dashboard'))
        row_data = db_sheet.row_values(cell.row)
        
        creator = row_data[11].strip() 
        link_bureau = row_data[3].strip() if len(row_data) > 3 else ''
        
        user_role = session.get('level', '').strip()
        username = session.get('username', '').strip()
        user_bureau = session.get('main_agency', '').strip()

        can_delete = False
        if user_role == 'Super Admin':
            can_delete = True
        elif user_role == 'Admin' and link_bureau == user_bureau:
            can_delete = True
        elif creator == username:
            can_delete = True

        if can_delete:
            db_sheet.delete_rows(cell.row)
            clear_db_cache()
            flash('ลบสำเร็จ', 'success')
        else:
            flash('ไม่มีสิทธิ์', 'error')
            
        return redirect(url_for('dashboard'))
    except: return redirect(url_for('dashboard'))

# =======================================================
# [UPDATED] สิทธิ์การหน้าแก้ไขลิงก์ (รองรับ 3 Role)
# =======================================================
@app.route('/edit/<link_id>')
def edit_link_page(link_id):
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        all_links = get_db_records()
        link = next((l for l in all_links if l.get('ID', '').strip() == link_id.strip()), None)
        if not link: 
            flash(f'ไม่พบลิงก์ ID: {link_id}', 'error')
            return redirect(url_for('dashboard'))
            
        creator = link.get('CreatorUsername', '').strip()
        link_bureau = link.get('ส่วนราชการ', '').strip()
        
        user_role = session.get('level', '').strip()
        username = session.get('username', '').strip()
        user_bureau = session.get('main_agency', '').strip()
        
        can_edit = False
        is_super_admin = (user_role == 'Super Admin')

        if is_super_admin:
            can_edit = True
        elif user_role == 'Admin' and link_bureau == user_bureau:
            can_edit = True
        elif creator == username:
            can_edit = True
        
        if not can_edit:
            flash('คุณไม่มีสิทธิ์แก้ไขลิงก์นี้', 'error')
            return redirect(url_for('dashboard'))
            
        # ทั้ง Super Admin และ Admin สามารถ ปักหมุด+แก้ไขเวลา ได้
        can_pin = (user_role in ['Super Admin', 'Admin'])
            
        return render_template('edit_link.html', session=session, link=link, locked_agency=link.get('ส่วนราชการ', 'N/A'), is_super_admin=is_super_admin, can_pin=can_pin)
    except Exception as e: return redirect(url_for('dashboard'))

# =======================================================
# [UPDATED] สิทธิ์การบันทึกแก้ไขลิงก์ (รองรับ 3 Role)
# =======================================================
@app.route('/update_action/<link_id>', methods=['POST'])
def update_link_action(link_id):
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        cell = db_sheet.find(link_id)
        if not cell: return redirect(url_for('dashboard'))
        row_vals = db_sheet.row_values(cell.row)
        
        creator = row_vals[11].strip()
        link_bureau = row_vals[3].strip() if len(row_vals) > 3 else ''
        
        user_role = session.get('level', '').strip()
        username = session.get('username', '').strip()
        user_bureau = session.get('main_agency', '').strip()
        
        can_edit = False
        if user_role == 'Super Admin': can_edit = True
        elif user_role == 'Admin' and link_bureau == user_bureau: can_edit = True
        elif creator == username: can_edit = True

        if not can_edit:
             flash('ไม่มีสิทธิ์', 'error')
             return redirect(url_for('dashboard'))

        data = {k: request.form.get(k) for k in ['ประเภท', 'หน่วยงาน', 'อีเมลผู้รับผิดชอบ', 'เบอร์โทรติดต่อ', 'ชื่อลิงก์', 'URL', 'สถานะ', 'รายละเอียด', 'ความเป็นส่วนตัว']}
        if data['เบอร์โทรติดต่อ'] and not data['เบอร์โทรติดต่อ'].startswith("'"): data['เบอร์โทรติดต่อ'] = "'" + data['เบอร์โทรติดต่อ']
        
        # เฉพาะ Super Admin เท่านั้นที่ย้ายลิงก์ข้ามกองได้
        main_agency = request.form.get('ส่วนราชการ') if user_role == 'Super Admin' else row_vals[3]
        
        current_clicks = row_vals[14] if len(row_vals) > 14 else 0
        try: current_edits = int(row_vals[15]) if len(row_vals) > 15 and row_vals[15] else 0
        except: current_edits = 0
        new_edit_count = current_edits + 1

        current_pinned = row_vals[16] if len(row_vals) > 16 else 'ไม่ปักหมุด'
        current_end_date = row_vals[17] if len(row_vals) > 17 else ''
        
        # Super Admin และ Admin อัปเดตสถานะปักหมุดได้
        if user_role in ['Super Admin', 'Admin']:
            is_pinned = 'ปักหมุด' if request.form.get('is_pinned') else 'ไม่ปักหมุด'
            end_date = request.form.get('end_date', '')
        else:
            is_pinned = current_pinned
            end_date = current_end_date

        new_vals = [
            link_id, data['ประเภท'], data['หน่วยงาน'], main_agency, 
            data['อีเมลผู้รับผิดชอบ'], data['เบอร์โทรติดต่อ'], data['ชื่อลิงก์'], 
            data['URL'], data['สถานะ'], data['รายละเอียด'], 
            get_current_timestamp(), creator, row_vals[12] if len(row_vals) > 12 else '',
            data.get('ความเป็นส่วนตัว', 'สาธารณะ'), current_clicks, new_edit_count, is_pinned, end_date 
        ]
        
        range_name = f"A{cell.row}:R{cell.row}" 
        db_sheet.update(range_name, [new_vals])
        clear_db_cache()
        flash(f'แก้ไขสำเร็จ (แก้ไขไปแล้ว {new_edit_count} ครั้ง)', 'success')
        return redirect(url_for('dashboard')) 
    except Exception as e: return redirect(url_for('dashboard'))
        
@app.route('/analytics')
def analytics_page():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    if db_sheet is None: return render_template('analytics.html', session=session, chart_data={})
    try:
        links = get_db_records()
        users = get_staff_records()
        feedback = feedback_sheet.get_all_records()
        
        for l in links:
            try: l['Clicks'] = int(l.get('Clicks', 0))
            except: l['Clicks'] = 0
            try: l['EditCount'] = int(l.get('EditCount', 0))
            except: l['EditCount'] = 0

        user_stats, bureau_stats, division_stats = {}, {}, {}

        for l in links:
            creator = l.get('CreatorUsername', '')
            bureau = l.get('ส่วนราชการ', 'ไม่ระบุ')
            division = l.get('หน่วยงาน', 'ไม่ระบุ')
            edits = l['EditCount']
            score = 1 + edits 

            if creator:
                if creator not in user_stats: user_stats[creator] = {'username': creator, 'score': 0, 'links': 0, 'edits': 0}
                user_stats[creator]['score'] += score
                user_stats[creator]['links'] += 1
                user_stats[creator]['edits'] += edits

            if bureau and bureau != 'ไม่ระบุ' and bureau != 'N/A':
                if bureau not in bureau_stats: bureau_stats[bureau] = {'name': bureau, 'score': 0, 'links': 0, 'edits': 0}
                bureau_stats[bureau]['score'] += score
                bureau_stats[bureau]['links'] += 1
                bureau_stats[bureau]['edits'] += edits

            if division and division != 'ไม่ระบุ' and division != 'N/A':
                if division not in division_stats: division_stats[division] = {'name': division, 'score': 0, 'links': 0, 'edits': 0}
                division_stats[division]['score'] += score
                division_stats[division]['links'] += 1
                division_stats[division]['edits'] += edits

        for creator, data in user_stats.items():
            user_info = next((u for u in users if u['Username'] == creator), None)
            if user_info:
                data['fullname'] = user_info.get('ชื่อ', creator)
                data['bureau'] = user_info.get('ส่วนราชการ', 'ไม่ระบุ')
                data['division'] = user_info.get('หน่วยงาน', 'ไม่ระบุ')
            else:
                data['fullname'] = creator; data['bureau'] = 'ไม่ระบุ'; data['division'] = 'ไม่ระบุ'

        top_10_users = sorted(user_stats.values(), key=lambda x: x['score'], reverse=True)[:10]
        top_10_bureaus = sorted(bureau_stats.values(), key=lambda x: x['score'], reverse=True)[:10]
        top_10_divisions = sorted(division_stats.values(), key=lambda x: x['score'], reverse=True)[:10]
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
            "total_views": total_views, "top_10_clicks": top_10_clicks, "top_10_edits": top_10_edits,
            "top_10_users": top_10_users, "top_10_bureaus": top_10_bureaus, "top_10_divisions": top_10_divisions,
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
    except: return redirect(url_for('dashboard'))

# =======================================================
# [UPDATED] สิทธิ์เข้าเมนูผู้ดูแลระบบ (เฉพาะ Super Admin)
# =======================================================
@app.route('/admin')
def admin_panel():
    if not session.get('logged_in') or session.get('level', '').strip() != 'Super Admin': 
        return redirect(url_for('dashboard'))
    try:
        return render_template('admin_panel.html', session=session, staff_list=get_staff_records(), invite_codes=get_invite_codes()) 
    except Exception as e: return f"Error loading admin panel: {e} <br><a href='/dashboard'>Back</a>"

@app.route('/admin/change_level', methods=['POST'])
def change_user_level():
    if not session.get('logged_in') or session.get('level', '').strip() != 'Super Admin': return redirect(url_for('login_page'))
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
    if not session.get('logged_in') or session.get('level', '').strip() != 'Super Admin': return redirect(url_for('login_page'))
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
    if not session.get('logged_in') or session.get('level', '').strip() != 'Super Admin': 
        return redirect(url_for('login_page'))
    try:
        try: amount = int(request.form.get('amount', 1))
        except ValueError: amount = 1
            
        if amount < 1: amount = 1
        if amount > 50: amount = 50

        new_rows = []
        for _ in range(amount):
            new_rows.append([generate_invite_code(), 'Available', '', ''])

        if invite_sheet:
            invite_sheet.append_rows(new_rows, value_input_option='USER_ENTERED')
            clear_invite_cache()
            flash(f'✅ สร้างรหัสเชิญใหม่จำนวน {amount} รหัส สำเร็จเรียบร้อย!', 'success')
        else: flash('ไม่สามารถเชื่อมต่อฐานข้อมูลได้', 'error')
    except Exception as e: flash(f'เกิดข้อผิดพลาด: {e}', 'error')
        
    return redirect(url_for('admin_panel'))
    
@app.route('/admin/delete_code', methods=['POST'])
def delete_code():
    if not session.get('logged_in') or session.get('level', '').strip() != 'Super Admin': return redirect(url_for('login_page'))
    try:
        cell = invite_sheet.find(request.form.get('code'))
        if cell: 
            invite_sheet.delete_rows(cell.row)
            clear_invite_cache()
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
    if session.get('logged_in'): user_key = f"user:{session.get('username')}"
    else:
        if not session.get('visitor_id'): session['visitor_id'] = str(uuid.uuid4())
        user_key = f"guest:{session.get('visitor_id')}"
        
    online_users[user_key] = datetime.datetime.now()
    cleanup_online_users()
    return jsonify({'status': 'ok', 'online_count': len(online_users)})

@app.route('/offline', methods=['POST'])
def offline():
    global online_users
    try:
        if session.get('logged_in'): user_key = f"user:{session.get('username')}"
        else:
            if session.get('visitor_id'): user_key = f"guest:{session.get('visitor_id')}"
            else: return '', 204
        if user_key in online_users: del online_users[user_key]
    except: pass
    return '', 204
    
# ==========================================
# [NEW] My Favorites System Logic
# ==========================================
def get_user_favorite_ids(username):
    if favorites_sheet is None: return []
    try:
        all_favs = favorites_sheet.get_all_records()
        return [row['LinkID'] for row in all_favs if row['Username'] == username]
    except: return []

@app.route('/toggle_favorite', methods=['POST'])
def toggle_favorite():
    if not session.get('logged_in'): return jsonify({'status': 'error', 'message': 'Please login first'}), 401
    if favorites_sheet is None: return jsonify({'status': 'error', 'message': 'Database not connected'}), 500
    try:
        data = request.get_json()
        link_id = data.get('link_id')
        username = session.get('username')
        
        all_records = favorites_sheet.get_all_values() 
        found_row_index = -1
        
        for i, row in enumerate(all_records):
            if i == 0: continue
            if len(row) > 2 and row[1] == username and row[2] == link_id:
                found_row_index = i + 1
                break
        
        if found_row_index != -1:
            favorites_sheet.delete_rows(found_row_index)
            action = 'removed'
        else:
            favorites_sheet.append_row([get_current_timestamp(), username, link_id], value_input_option='USER_ENTERED')
            action = 'added'
            
        return jsonify({'status': 'success', 'action': action})
    except Exception as e: return jsonify({'status': 'error', 'message': str(e)}), 500
        
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=True)