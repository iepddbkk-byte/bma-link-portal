import gspread
import qrcode
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
from collections import Counter
from threading import Thread 
from PIL import Image # (สำคัญ!) สำหรับจัดการ Logo

# Import Flask-Mail และ itsdangerous
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
app.config['MAIL_USERNAME'] = os.environ.get('MAIL_USERNAME', 'your-email@gmail.com') 
app.config['MAIL_PASSWORD'] = os.environ.get('MAIL_PASSWORD', 'xxxx xxxx xxxx xxxx')  
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

DB_HEADERS = [
    "ID", "ประเภท", "หน่วยงาน", "ส่วนราชการ", "อีเมลผู้รับผิดชอบ", "เบอร์โทรติดต่อ",
    "ชื่อลิงก์", "URL", "สถานะ", "รายละเอียด", "วันที่อัปเดต", "CreatorUsername", "LinkStatus"
]

STAFF_HEADERS = [
    "Username", "PasswordHash", "Level", "ชื่อ", "ตำแหน่ง", 
    "หน่วยงาน", "ส่วนราชการ", "เบอร์โทร", "Email", "CreatedAt", "UpdatedAt"
]

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
    """ แปลงข้อมูลและตัดช่องว่าง (Strip) พร้อมกรองแถวว่าง """
    records_dict = []
    for row in records_list[1:]:
        # (สำคัญ!) กรองแถวว่าง: ต้องมี ID (คอลัมน์ 0)
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

def get_db_records():
    if db_sheet is None: return []
    all_values = db_sheet.get_all_values()
    return records_to_dict(all_values, DB_HEADERS)

def get_staff_records():
    if staff_sheet is None: return []
    all_values = staff_sheet.get_all_values()
    return records_to_dict(all_values, STAFF_HEADERS)

def send_reset_email(username, recipient_email):
    try:
        token = s.dumps(username, salt='password-reset-salt')
        reset_url = url_for('reset_password_page', token=token, _external=True)
        msg = Message("คำขอรีเซ็ตรหัสผ่าน - BMA Link Registry", recipients=[recipient_email])
        msg.body = f"""สวัสดีครับ,\nเราได้รับคำขอรีเซ็ตรหัสผ่านสำหรับ Username: {username}\nคลิกลิงค์เพื่อตั้งรหัสผ่านใหม่: {reset_url}\n(ลิงค์หมดอายุใน 1 ชั่วโมง)"""
        mail.send(msg)
        return True
    except Exception as e:
        print(f"Mail Error: {e}")
        return False

# --- 5. API ---
@app.route('/check_username', methods=['POST'])
def check_username():
    if staff_sheet is None: return jsonify({'available': False, 'message': 'DB Error'})
    try:
        data = request.get_json()
        username = data.get('username')
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
            
            updates.append({'range': f'M{row_index}', 'values': [[status_msg]]})
        
        if updates:
            db_sheet.batch_update(updates, value_input_option='RAW')
            return jsonify({'status': 'success', 'message': f'Checked {len(updates)} links'})
            
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500
    return jsonify({'status': 'success', 'message': 'No links checked'})

# --- QR CODE GENERATION (WITH LOGO) ---
@app.route('/generate_qr')
def generate_qr_code():
    url_to_encode = request.args.get('url')
    if not url_to_encode:
        return "No URL provided", 400
    
    try:
        # สร้าง QR Code
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_H, # H = High Error Correction (จำเป็นสำหรับใส่ Logo)
            box_size=10,
            border=4,
        )
        qr.add_data(url_to_encode)
        qr.make(fit=True)

        # สร้างรูปภาพ QR (ต้องแปลงเป็น RGBA เพื่อรองรับความโปร่งใสของ Logo)
        img = qr.make_image(fill_color="black", back_color="white").convert('RGBA')
        
        # (ใหม่!) ฝังโลโก้
        logo_path = os.path.join(app.root_path, 'pictures', 'qr_logo.png')
        
        if os.path.exists(logo_path):
            logo = Image.open(logo_path).convert('RGBA')
            
            # คำนวณขนาดโลโก้ (เช่น 25% ของ QR Code)
            qr_width, qr_height = img.size
            logo_size = int(qr_width * 0.25) 
            
            # ปรับขนาดโลโก้
            logo = logo.resize((logo_size, logo_size), Image.LANCZOS)
            
            # คำนวณตำแหน่งวาง (ตรงกลาง)
            pos = ((qr_width - logo_size) // 2, (qr_height - logo_size) // 2)
            
            # วางโลโก้ลงไป (ใช้ logo เป็น mask ด้วยเพื่อความคมชัดของขอบใส)
            img.paste(logo, pos, logo)
        
        # บันทึกลง Memory
        img_io = BytesIO()
        img.save(img_io, 'PNG')
        img_io.seek(0)
        
        return send_file(img_io, mimetype='image/png', as_attachment=False)

    except Exception as e:
        print(f"QR Error: {e}")
        return f"QR Generation Error: {e}", 500

# Indirection Layer
@app.route('/go/<link_id>')
def go_to_link(link_id):
    try:
        all_links = get_db_records()
        link = next((l for l in all_links if l.get('ID', '').strip() == link_id.strip()), None)
        
        if link and link.get('URL'):
            return redirect(link['URL']) 
        return render_template('error.html', message="ไม่พบลิงค์ที่คุณค้นหา") 
    except Exception as e:
        print(f"Go Error: {e}")
        return render_template('error.html', message="Error connecting to database") 

# --- 6. Routes (General) ---
@app.route('/')
def home():
    if db_sheet is None: 
        return render_template('index.html', session=session, bureaus=bureaus_list, districts=districts_list, links=[], error="Sheet Error")
    try:
        all_records = get_db_records()
        links_for_home = [link for link in all_records if link.get('สถานะ') == 'ใช้งาน']
        for link in links_for_home:
            date_str = link.get('วันที่อัปเดต')
            link['วันที่อัปเดต_สั้น'] = date_str.split(' ')[0] if date_str else ''

        return render_template('index.html', 
                               session=session,
                               bureaus=bureaus_list,
                               districts=districts_list,
                               links=links_for_home, 
                               error=None)
    except Exception as e: 
        return render_template('index.html', session=session, bureaus=bureaus_list, districts=districts_list, links=[], error=str(e))

@app.route('/links')
def links_page():
    if db_sheet is None: 
        return render_template('links_page.html', links=[], error="Sheet Error", session=session, agency_name="Error")
    try:
        agency_filter = request.args.get('agency') 
        all_records = get_db_records()
        
        if agency_filter:
            links_to_display = [l for l in all_records if l.get('สถานะ') == 'ใช้งาน' and l.get('ส่วนราชการ') == agency_filter]
            page_title = agency_filter
        else:
            links_to_display = [l for l in all_records if l.get('สถานะ') == 'ใช้งาน']
            page_title = "ลิงค์ทั้งหมด"
        
        for link in links_to_display:
            date_str = link.get('วันที่อัปเดต')
            link['วันที่อัปเดต_สั้น'] = date_str.split(' ')[0] if date_str else ''

        return render_template('links_page.html', 
                               links=links_to_display, 
                               error=None, 
                               session=session,
                               agency_name=page_title,
                               bureaus=bureaus_list,
                               districts=districts_list)
    except Exception as e: 
        return render_template('links_page.html', links=[], error=str(e), session=session, agency_name="Error")

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
        
        new_row_final = [
            username, hashed_password, 'Users', fullname, position, division, 
            main_agency, phone, email, current_time, current_time
        ]
        staff_sheet.append_row(new_row_final, value_input_option='USER_ENTERED')
        
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
            if user_found:
                send_reset_email(user_found['Username'], user_found['Email'])
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
        all_links = get_db_records() 
        for link in all_links:
            date_str = link.get('วันที่อัปเดต')
            link['วันที่อัปเดต_สั้น'] = date_str.split(' ')[0] if date_str else ''

        return render_template('dashboard.html', session=session, links=all_links, bureaus=bureaus_list, districts=districts_list)  
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
        data = {k: request.form.get(k) for k in ['ประเภท', 'หน่วยงาน', 'อีเมลผู้รับผิดชอบ', 'เบอร์โทรติดต่อ', 'ชื่อลิงก์', 'URL', 'รายละเอียด', 'สถานะ']}
        if data['เบอร์โทรติดต่อ'] and not data['เบอร์โทรติดต่อ'].startswith("'"): data['เบอร์โทรติดต่อ'] = "'" + data['เบอร์โทรติดต่อ']
        main_agency = session.get('main_agency', 'N/A') 
        
        new_row = [
            generate_new_id(), data['ประเภท'], data['หน่วยงาน'], main_agency, 
            data['อีเมลผู้รับผิดชอบ'], data['เบอร์โทรติดต่อ'], data['ชื่อลิงก์'], 
            data['URL'], data['สถานะ'], data['รายละเอียด'], 
            get_current_timestamp(), session.get('username'), ''
        ]
        db_sheet.append_row(new_row, value_input_option='USER_ENTERED')
        flash('เพิ่มลิงค์สำเร็จ', 'success'); return redirect(url_for('dashboard'))
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
        if session.get('level', '').strip() == 'Admin' or creator == session.get('username', '').strip():
            db_sheet.delete_rows(cell.row)
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
        print(f'[ERROR] Edit Page Error: {str(e)}')
        return redirect(url_for('dashboard'))

@app.route('/update_action/<link_id>', methods=['POST'])
def update_link_action(link_id):
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        cell = db_sheet.find(link_id)
        if not cell: return redirect(url_for('dashboard'))

        row_vals = db_sheet.row_values(cell.row)
        creator = row_vals[11].strip() 
        
        if session.get('level', '').strip() != 'Admin' and creator != session.get('username', '').strip():
             flash('ไม่มีสิทธิ์', 'error'); return redirect(url_for('dashboard'))
        
        data = {k: request.form.get(k) for k in ['ประเภท', 'หน่วยงาน', 'อีเมลผู้รับผิดชอบ', 'เบอร์โทรติดต่อ', 'ชื่อลิงก์', 'URL', 'สถานะ', 'รายละเอียด']}
        if data['เบอร์โทรติดต่อ'] and not data['เบอร์โทรติดต่อ'].startswith("'"): data['เบอร์โทรติดต่อ'] = "'" + data['เบอร์โทรติดต่อ']

        main_agency = request.form.get('ส่วนราชการ') if session.get('level', '').strip() == 'Admin' else row_vals[3]
        
        new_vals = [
            link_id, data['ประเภท'], data['หน่วยงาน'], main_agency, 
            data['อีเมลผู้รับผิดชอบ'], data['เบอร์โทรติดต่อ'], data['ชื่อลิงก์'], 
            data['URL'], data['สถานะ'], data['รายละเอียด'], 
            get_current_timestamp(), creator, row_vals[12] if len(row_vals) > 12 else ''
        ]
        range_name = f"A{cell.row}:M{cell.row}"
        db_sheet.update(range_name, [new_vals])
        flash('แก้ไขสำเร็จ', 'success'); return redirect(url_for('dashboard'))
    except: return redirect(url_for('dashboard'))

# --- 10. Routes (Admin & Analytics & Feedback) ---
@app.route('/analytics')
def analytics_page():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    if db_sheet is None: return render_template('analytics.html', session=session, chart_data={})
    try:
        links = get_db_records()
        users = get_staff_records()
        feedback = feedback_sheet.get_all_records()
        
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
        
        sat = [int(f['SatisfactionScore']) for f in feedback if f.get('SatisfactionScore')]
        ease = [int(f['EaseOfUseScore']) for f in feedback if f.get('EaseOfUseScore')]
        comments = [{'user': f['Username'], 'text': f['Comments']} for f in feedback if f.get('Comments')]
        features = [{'user': f['Username'], 'text': f['FeatureRequest']} for f in feedback if f.get('FeatureRequest')]

        chart_data = {
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

@app.route('/admin')
def admin_panel():
    if not session.get('logged_in') or session.get('level', '').strip() != 'Admin': return redirect(url_for('dashboard'))
    try:
        return render_template('admin_panel.html', session=session, staff_list=get_staff_records(), invite_codes=invite_sheet.get_all_records())
    except: return redirect(url_for('dashboard'))

# ... (Admin Actions - change_level, delete_user, generate_code, delete_code ... เหมือนเดิม) ...
@app.route('/admin/change_level', methods=['POST'])
def change_user_level():
    if not session.get('logged_in') or session.get('level', '').strip() != 'Admin': return redirect(url_for('login_page'))
    try:
        user, level = request.form.get('username'), request.form.get('level')
        if user == session.get('username'): flash('เปลี่ยนระดับตัวเองไม่ได้', 'error'); return redirect(url_for('admin_panel'))
        cell = staff_sheet.find(user)
        if cell: staff_sheet.update_cell(cell.row, 3, level); flash('สำเร็จ', 'success')
    except: flash('Error', 'error')
    return redirect(url_for('admin_panel'))

@app.route('/admin/delete_user', methods=['POST'])
def delete_user():
    if not session.get('logged_in') or session.get('level', '').strip() != 'Admin': return redirect(url_for('login_page'))
    try:
        user = request.form.get('username')
        if user == session.get('username'): flash('ลบตัวเองไม่ได้', 'error'); return redirect(url_for('admin_panel'))
        cell = staff_sheet.find(user)
        if cell: staff_sheet.delete_rows(cell.row); flash('สำเร็จ', 'success')
    except: flash('Error', 'error')
    return redirect(url_for('admin_panel'))

@app.route('/admin/generate_code', methods=['POST'])
def generate_code():
    if not session.get('logged_in') or session.get('level', '').strip() != 'Admin': return redirect(url_for('login_page'))
    try:
        code = generate_invite_code()
        invite_sheet.append_row([code, 'Available', '', ''], value_input_option='USER_ENTERED')
        flash(f'สร้างรหัส: {code}', 'success')
    except: flash('Error', 'error')
    return redirect(url_for('admin_panel'))

@app.route('/admin/delete_code', methods=['POST'])
def delete_code():
    if not session.get('logged_in') or session.get('level', '').strip() != 'Admin': return redirect(url_for('login_page'))
    try:
        cell = invite_sheet.find(request.form.get('code'))
        if cell: invite_sheet.delete_rows(cell.row); flash('ลบสำเร็จ', 'success')
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
        row = [get_current_timestamp(), session.get('username'), request.form.get('satisfaction'), 
               request.form.get('ease_of_use'), request.form.get('comments', ''), request.form.get('features', '')]
        feedback_sheet.append_row(row, value_input_option='USER_ENTERED')
        flash('ขอบคุณสำหรับข้อเสนอแนะ', 'success'); return redirect(url_for('dashboard'))
    except: return redirect(url_for('feedback_page'))

@app.route('/get_links', methods=['GET'])
def get_all_links():
    try: return jsonify({"status": "success", "data": get_db_records()})
    except: return jsonify({"status": "error"}), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)