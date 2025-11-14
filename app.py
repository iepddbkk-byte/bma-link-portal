import gspread
from oauth2client.service_account import ServiceAccountCredentials
from flask import (
    Flask, jsonify, request, render_template, 
    session, redirect, url_for, flash
)
from werkzeug.security import check_password_hash, generate_password_hash
import datetime
import random
import string
import os
import json # (เพิ่ม import json)
from collections import Counter

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

# ใช้ค่าจาก Environment หรือค่าสำรองที่คุณใส่มา
app.config['MAIL_USERNAME'] = os.environ.get('MAIL_USERNAME', 'iepdd.bkk@gmail.com') 
app.config['MAIL_PASSWORD'] = os.environ.get('MAIL_PASSWORD', 'haus xdqu afbt hgxz')  

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
    # (รองรับทั้ง Vercel และ Local)
    json_creds = os.environ.get('GOOGLE_CREDENTIALS')
    
    if json_creds:
        # กรณีรันบน Server (อ่านจากตัวแปร)
        creds_dict = json.loads(json_creds)
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, SCOPE)
    else:
        # กรณีรันในเครื่อง (อ่านจากไฟล์)
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


# --- 4. ฟังก์ชันตัวช่วย (Helper Functions) ---
def generate_new_id():
    code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
    return f"BMA-{code}"

def get_current_timestamp():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def generate_invite_code():
    code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
    return f"INVITE-{code}"

def send_reset_email(username, recipient_email):
    """ 
    สร้าง Token และส่งอีเมล (แบบ Synchronous - รอจนเสร็จ) 
    เพื่อให้ทำงานบน Vercel/Serverless ได้
    """
    try:
        token = s.dumps(username, salt='password-reset-salt')
        # _external=True เพื่อให้ได้ URL เต็ม (https://...)
        reset_url = url_for('reset_password_page', token=token, _external=True)
        
        msg_title = "คำขอรีเซ็ตรหัสผ่าน - BMA Link Registry"
        msg_body = f"""
        สวัสดีครับ,
        เราได้รับคำขอรีเซ็ตรหัสผ่านสำหรับ Username: {username}
        หากคุณเป็นผู้ร้องขอ กรุณาคลิกลิงค์ด้านล่างเพื่อตั้งรหัสผ่านใหม่:
        
        {reset_url}
        
        (ลิงค์นี้จะหมดอายุภายใน 1 ชั่วโมง)
        
        ขอบคุณครับ
        BMA Link Registry
        """
        msg = Message(msg_title, recipients=[recipient_email], body=msg_body)
        
        # ส่งทันที (รอจนกว่าจะเสร็จ)
        mail.send(msg)
        print(f"✅ ส่งอีเมลรีเซ็ตไปยัง {recipient_email} สำเร็จ")
        return True
    except Exception as e:
        print(f"❌ ส่งอีเมลล้มเหลว: {e}")
        return False


# --- API (ตรวจสอบ Username/Invite Code) ---
@app.route('/check_username', methods=['POST'])
def check_username():
    if staff_sheet is None:
        return jsonify({'available': False, 'message': 'Cannot connect to user database'})
    try:
        data = request.get_json()
        username_to_check = data.get('username')
        if not username_to_check:
            return jsonify({'available': False, 'message': 'Username is required'})
        all_usernames = staff_sheet.col_values(1) 
        is_available = username_to_check.lower() not in [u.lower() for u in all_usernames]
        return jsonify({'available': is_available})
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดที่ /check_username: {e}")
        return jsonify({'available': False, 'message': 'Server error'})

@app.route('/check_invite_code', methods=['POST'])
def check_invite_code():
    if invite_sheet is None:
        return jsonify({'available': False, 'message': 'ไม่สามารถเชื่อมต่อฐานข้อมูลรหัสเชิญได้'})
    try:
        data = request.get_json()
        code_to_check = data.get('invite_code')
        if not code_to_check:
            return jsonify({'available': False, 'message': 'กรุณากรอกรหัสเชิญ'})
        cell = invite_sheet.find(code_to_check)
        if not cell:
            return jsonify({'available': False, 'message': 'รหัสเชิญไม่ถูกต้อง'})
        status = invite_sheet.cell(cell.row, 2).value 
        if status == 'Available':
            return jsonify({'available': True, 'message': 'รหัสเชิญถูกต้อง'})
        else:
            return jsonify({'available': False, 'message': 'รหัสเชิญนี้ถูกใช้งานไปแล้ว'})
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดที่ /check_invite_code: {e}")
        return jsonify({'available': False, 'message': 'Server error'})


# --- 5. Routes (หน้าสาธารณะ) ---
@app.route('/')
def home():
    if db_sheet is None:
        return render_template('index.html', links=[], error="Sheet Error", session=session)
    try:
        all_records = db_sheet.get_all_records()
        links_to_display = [link for link in all_records if link.get('สถานะ') == 'ใช้งาน']
        return render_template('index.html', links=links_to_display, error=None, session=session)
    except Exception as e:
        return render_template('index.html', links=[], error=str(e), session=session)


# --- 6. Routes (ระบบสมาชิก Auth & Reset) ---
@app.route('/login')
def login_page():
    if session.get('logged_in'):
        return redirect(url_for('dashboard'))
    return render_template('login.html') 

@app.route('/login_action', methods=['POST'])
def login_action():
    if staff_sheet is None:
        flash('ไม่สามารถเชื่อมต่อฐานข้อมูลผู้ใช้งานได้', 'error')
        return redirect(url_for('login_page'))
    try:
        username_input = request.form.get('username')
        password_input = request.form.get('password')
        staff_list = staff_sheet.get_all_records()
        user_found = None
        for user in staff_list:
            if user['Username'].lower() == username_input.lower():
                user_found = user
                break
        if user_found and check_password_hash(user_found['PasswordHash'], password_input):
            session['logged_in'] = True
            session['username'] = user_found['Username']
            session['level'] = user_found['Level']
            session['name'] = user_found['ชื่อ']
            session['email'] = user_found['Email'] 
            flash('เข้าสู่ระบบสำเร็จ!', 'success') 
            return redirect(url_for('dashboard'))
        else:
            flash('Username หรือ รหัสผ่าน ไม่ถูกต้อง', 'error')
            return redirect(url_for('login_page'))
    except Exception as e:
        flash(f'เกิดข้อผิดพลาด: {e}', 'error')
        return redirect(url_for('login_page'))

@app.route('/logout')
def logout():
    session.clear() 
    flash('ออกจากระบบเรียบร้อยแล้ว', 'info')
    return redirect(url_for('home'))

@app.route('/register')
def register_page():
    if session.get('logged_in'):
        return redirect(url_for('dashboard'))
    return render_template('register.html') 

@app.route('/register_action', methods=['POST'])
def register_action():
    if staff_sheet is None or invite_sheet is None: 
        flash('ไม่สามารถเชื่อมต่อฐานข้อมูลผู้ใช้งานได้', 'error')
        return redirect(url_for('register_page'))
    try:
        username = request.form.get('username')
        password = request.form.get('password')
        fullname = request.form.get('fullname')
        position = request.form.get('position')
        department = request.form.get('department')
        email = request.form.get('email')
        phone = request.form.get('phone')
        invite_code = request.form.get('invite_code')

        all_usernames = staff_sheet.col_values(1) 
        if username.lower() in [u.lower() for u in all_usernames]:
            flash('Username นี้มีผู้ใช้งานแล้ว กรุณาใช้ชื่ออื่น', 'error')
            return redirect(url_for('register_page'))

        code_cell = invite_sheet.find(invite_code)
        if not code_cell:
            flash('รหัสเชิญไม่ถูกต้อง', 'error')
            return redirect(url_for('register_page'))
        status = invite_sheet.cell(code_cell.row, 2).value 
        if status != 'Available':
            flash('รหัสเชิญนี้ถูกใช้งานไปแล้ว', 'error')
            return redirect(url_for('register_page'))

        hashed_password = generate_password_hash(password)
        level = 'Users'
        current_time = get_current_timestamp()
        new_row = [
            username, hashed_password, level, fullname, position,
            department, phone, email, current_time, current_time
        ]
        staff_sheet.append_row(new_row, value_input_option='USER_ENTERED')
        invite_sheet.update_cell(code_cell.row, 2, "Used")
        invite_sheet.update_cell(code_cell.row, 3, username)
        invite_sheet.update_cell(code_cell.row, 4, current_time)
        flash('สมัครสมาชิกสำเร็จ! กรุณาเข้าสู่ระบบ', 'success')
        return redirect(url_for('login_page')) 
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการสมัครสมาชิก: {e}")
        flash(f'เกิดข้อผิดพลาด: {e}', 'error')
        return redirect(url_for('register_page'))

# --- 7. Routes (ระบบลืมรหัสผ่าน) ---
@app.route('/forgot_password', methods=['GET', 'POST'])
def forgot_password():
    if request.method == 'POST':
        if staff_sheet is None:
            flash('ไม่สามารถเชื่อมต่อฐานข้อมูลผู้ใช้งานได้', 'error')
            return redirect(url_for('forgot_password'))
        try:
            username_input = request.form.get('username')
            all_staff = staff_sheet.get_all_records()
            user_found = next((user for user in all_staff if user['Username'].lower() == username_input.lower()), None)
            if user_found:
                username = user_found['Username']
                email = user_found['Email']
                # (แก้ไข!) ส่งแบบ Synchronous
                success = send_reset_email(username, email)
                if not success:
                     flash('เกิดข้อผิดพลาดในการส่งอีเมล', 'error')
                     return redirect(url_for('forgot_password'))

            flash('หาก Username นี้มีอยู่ในระบบ เราได้ส่งลิงค์รีเซ็ตรหัสผ่านไปให้แล้ว', 'info')
            return redirect(url_for('login_page'))
        except Exception as e:
            print(f"❌ เกิดข้อผิดพลาดในการส่งอีเมลรีเซ็ต: {e}")
            flash('เกิดข้อผิดพลาดในการส่งอีเมล (Error)', 'error')
            return redirect(url_for('forgot_password'))
    return render_template('forgot_password.html')

@app.route('/reset_password/<token>', methods=['GET', 'POST'])
def reset_password_page(token):
    try:
        username = s.loads(token, salt='password-reset-salt', max_age=3600)
    except SignatureExpired:
        flash('ลิงค์รีเซ็ตรหัสผ่านหมดอายุแล้ว กรุณาลองอีกครั้ง', 'error')
        return redirect(url_for('forgot_password'))
    except Exception:
        flash('ลิงค์รีเซ็ตรหัสผ่านไม่ถูกต้อง', 'error')
        return redirect(url_for('forgot_password'))
    if request.method == 'POST':
        password = request.form.get('password')
        confirm_password = request.form.get('confirm_password')
        if password != confirm_password:
            flash('รหัสผ่านทั้งสองช่องไม่ตรงกัน', 'error')
            return render_template('reset_password.html', token=token)
        try:
            new_hash = generate_password_hash(password)
            current_time = get_current_timestamp()
            cell = staff_sheet.find(username) 
            if not cell:
                flash('ไม่พบผู้ใช้งานในระบบ', 'error')
                return redirect(url_for('login_page'))
            staff_sheet.update_cell(cell.row, 2, new_hash) 
            staff_sheet.update_cell(cell.row, 10, current_time) 
            flash('อัปเดตรหัสผ่านสำเร็จ! กรุณาเข้าสู่ระบบด้วยรหัสผ่านใหม่', 'success')
            return redirect(url_for('login_page'))
        except Exception as e:
            print(f"❌ เกิดข้อผิดพลาดในการอัปเดต Sheet: {e}")
            flash('เกิดข้อผิดพลาดในการอัปเดตรหัสผ่าน', 'error')
            return redirect(url_for('login_page'))
    return render_template('reset_password.html', token=token)

# --- 7.5: Routes (หน้า Analytics Dashboard) ---
@app.route('/analytics')
def analytics_page():
    if not session.get('logged_in'):
        flash('กรุณาเข้าสู่ระบบก่อน', 'error')
        return redirect(url_for('login_page'))
    
    if db_sheet is None or staff_sheet is None or feedback_sheet is None:
        flash('ไม่สามารถเชื่อมต่อฐานข้อมูลได้', 'error')
        return render_template('analytics.html', session=session, chart_data={})

    try:
        all_links = db_sheet.get_all_records()
        all_staff = staff_sheet.get_all_records()
        total_links = len(all_links)
        total_users = len(all_staff)
        active_links = sum(1 for link in all_links if link.get('สถานะ') == 'ใช้งาน')
        category_counts = Counter(link['ประเภท'] for link in all_links if link.get('ประเภท'))
        category_labels = list(category_counts.keys())
        category_data = list(category_counts.values())
        department_counts = Counter(link['หน่วยงาน'] for link in all_links if link.get('หน่วยงาน'))
        top_5_departments = department_counts.most_common(5)
        dept_labels = [dept[0] for dept in top_5_departments]
        dept_data = [dept[1] for dept in top_5_departments]
        monthly_counts = {}
        for link in all_links:
            timestamp_str = link.get('วันที่อัปเดต') 
            if timestamp_str:
                try:
                    dt = datetime.datetime.strptime(timestamp_str.split(' ')[0], '%Y-%m-%d')
                    month_key = dt.strftime('%Y-%m') 
                    monthly_counts[month_key] = monthly_counts.get(month_key, 0) + 1
                except ValueError: continue 
        sorted_months_items = sorted(monthly_counts.items())
        month_labels = [item[0] for item in sorted_months_items]
        month_data = [item[1] for item in sorted_months_items]
        
        all_feedback = feedback_sheet.get_all_records()
        sat_scores = []
        ease_scores = []
        recent_comments = []
        recent_features = []
        total_responses = len(all_feedback)

        for fb in all_feedback:
            try: sat_scores.append(int(fb['SatisfactionScore']))
            except (ValueError, KeyError): pass 
            try: ease_scores.append(int(fb['EaseOfUseScore']))
            except (ValueError, KeyError): pass 
            
            if fb.get('Comments'):
                recent_comments.append({"user": fb.get('Username'), "text": fb.get('Comments')})
            if fb.get('FeatureRequest'):
                recent_features.append({"user": fb.get('Username'), "text": fb.get('FeatureRequest')})
        
        avg_sat = round(sum(sat_scores) / len(sat_scores), 1) if sat_scores else 0
        avg_ease = round(sum(ease_scores) / len(ease_scores), 1) if ease_scores else 0
        
        recent_comments = recent_comments[-5:][::-1]
        recent_features = recent_features[-5:][::-1]

        chart_data = {
            "total_links": total_links, "total_users": total_users, "active_links": active_links,
            "category_labels": category_labels, "category_data": category_data,
            "dept_labels": dept_labels, "dept_data": dept_data,
            "month_labels": month_labels, "month_data": month_data,
            "total_responses": total_responses, "avg_sat": avg_sat, "avg_ease": avg_ease,
            "recent_comments": recent_comments, "recent_features": recent_features
        }
        
        return render_template('analytics.html', session=session, chart_data=chart_data)

    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการโหลด Analytics: {e}")
        flash(f'เกิดข้อผิดพลาดในการประมวลผล: {e}', 'error')
        return render_template('analytics.html', session=session, chart_data={})


# --- 8. Routes (ระบบจัดการ Dashboard & CRUD) ---
@app.route('/dashboard')
def dashboard():
    if not session.get('logged_in'):
        flash('กรุณาเข้าสู่ระบบก่อน', 'error')
        return redirect(url_for('login_page'))
    if db_sheet is None:
        flash('ไม่สามารถเชื่อมต่อฐานข้อมูลลิงค์ได้', 'error')
        return render_template('dashboard.html', session=session, links=[])
    try:
        all_links_data = db_sheet.get_all_records()
        return render_template('dashboard.html', session=session, links=all_links_data)
    except Exception as e:
        flash(f'เกิดข้อผิดพลาดในการโหลด Dashboard: {e}', 'error')
        return redirect(url_for('home'))

@app.route('/add')
def add_link_page():
    if not session.get('logged_in'):
        return redirect(url_for('login_page'))
    return render_template('add_link.html', session=session)

@app.route('/add_action', methods=['POST'])
def add_link_action():
    if not session.get('logged_in'):
        return redirect(url_for('login_page'))
    try:
        data = {
            'ประเภท': request.form.get('ประเภท'), 'หน่วยงาน': request.form.get('หน่วยงาน'),
            'อีเมลผู้รับผิดชอบ': request.form.get('อีเมลผู้รับผิดชอบ'), 'เบอร์โทรติดต่อ': request.form.get('เบอร์โทรติดต่อ'),
            'ชื่อลิงก์': request.form.get('ชื่อลิงก์'), 'URL': request.form.get('URL'),
            'หมายเหตุ': request.form.get('หมายเหตุ', ''), 'สถานะ': request.form.get('สถานะ')
        }
        new_id = generate_new_id()
        current_time = get_current_timestamp()
        creator_username = session.get('username') 
        new_row = [
            new_id, data['ประเภท'], data['หน่วยงาน'], data['อีเมลผู้รับผิดชอบ'],
            data['เบอร์โทรติดต่อ'], data['ชื่อลิงก์'], data['URL'],
            data['สถานะ'], data['หมายเหตุ'], current_time,
            creator_username
        ]
        db_sheet.append_row(new_row, value_input_option='USER_ENTERED')
        flash(f"เพิ่มลิงค์ใหม่สำเร็จ! (สถานะ: {data['สถานะ']})", 'success')
        return redirect(url_for('dashboard'))
    except Exception as e:
        flash(f'เกิดข้อผิดพลาด: {e}', 'error')
        return redirect(url_for('add_link_page'))

@app.route('/delete/<link_id>', methods=['POST'])
def delete_link_action(link_id):
    if not session.get('logged_in'):
        return redirect(url_for('login_page'))
    try:
        cell = db_sheet.find(link_id)
        if not cell:
            flash('ไม่พบลิงค์ที่ต้องการลบ', 'error')
            return redirect(url_for('dashboard'))
        row_data = db_sheet.get_all_records()
        link_to_delete = next((link for link in row_data if link['ID'] == link_id), None)
        if not link_to_delete:
            flash('ไม่พบข้อมูลลิงค์', 'error')
            return redirect(url_for('dashboard'))
        user_level = session.get('level')
        user_username = session.get('username')
        can_delete = False
        if user_level == 'Admin' or (user_level == 'Users' and link_to_delete.get('CreatorUsername') == user_username):
            can_delete = True
        if can_delete:
            db_sheet.delete_rows(cell.row) 
            flash(f'ลบลิงค์ ID: {link_id} สำเร็จ', 'success')
        else:
            flash(f'คุณไม่มีสิทธิ์ลบลิงค์ ID: {link_id}', 'error')
        return redirect(url_for('dashboard'))
    except Exception as e:
        flash(f'เกิดข้อผิดพลาดในการลบ: {e}', 'error')
        return redirect(url_for('dashboard'))

@app.route('/edit/<link_id>')
def edit_link_page(link_id):
    if not session.get('logged_in'):
        return redirect(url_for('login_page'))
    try:
        all_links = db_sheet.get_all_records()
        link_to_edit = next((link for link in all_links if link['ID'] == link_id), None)
        if not link_to_edit:
            flash(f'ไม่พบลิงค์ ID: {link_id}', 'error')
            return redirect(url_for('dashboard'))
        user_level = session.get('level')
        user_username = session.get('username')
        can_edit = False
        if user_level == 'Admin' or (user_level == 'Users' and link_to_edit.get('CreatorUsername') == user_username):
            can_edit = True
        if not can_edit:
            flash(f'คุณไม่มีสิทธิ์แก้ไขลิงค์ ID: {link_id}', 'error')
            return redirect(url_for('dashboard'))
        return render_template('edit_link.html', session=session, link=link_to_edit)
    except Exception as e:
        flash(f'เกิดข้อผิดพลาด: {e}', 'error')
        return redirect(url_for('dashboard'))

@app.route('/update_action/<link_id>', methods=['POST'])
def update_link_action(link_id):
    if not session.get('logged_in'):
        return redirect(url_for('login_page'))
    try:
        cell = db_sheet.find(link_id)
        if not cell:
            flash(f'ไม่พบลิงค์ ID: {link_id}', 'error')
            return redirect(url_for('dashboard'))
        original_data = db_sheet.get_all_records()
        link_data = next((link for link in original_data if link['ID'] == link_id), None)
        user_level = session.get('level')
        user_username = session.get('username')
        can_edit = False
        if user_level == 'Admin' or (user_level == 'Users' and link_data.get('CreatorUsername') == user_username):
            can_edit = True
        if not can_edit:
            flash(f'คุณไม่มีสิทธิ์อัปเดตลิงค์ ID: {link_id}', 'error')
            return redirect(url_for('dashboard'))
        updated_status = request.form.get('สถานะ')
        updated_row = [
            link_data['ID'], request.form.get('ประเภท'), request.form.get('หน่วยงาน'),
            request.form.get('อีเมลผู้รับผิดชอบ'), request.form.get('เบอร์โทรติดต่อ'), request.form.get('ชื่อลิงก์'),
            request.form.get('URL'), updated_status, request.form.get('หมายเหตุ'),
            get_current_timestamp(), link_data['CreatorUsername']
        ]
        range_to_update = f'A{cell.row}:K{cell.row}'
        db_sheet.update(range_to_update, [updated_row])
        flash(f'แก้ไขลิงค์ ID: {link_id} สำเร็จ!', 'success')
        return redirect(url_for('dashboard'))
    except Exception as e:
        flash(f'เกิดข้อผิดพลาดในการอัปเดต: {e}', 'error')
        return redirect(url_for('edit_link_page', link_id=link_id))

@app.route('/get_links', methods=['GET'])
def get_all_links():
    if db_sheet is None:
        return jsonify({"status": "error", "message": "Sheet connection failed"}), 500
    try:
        records = db_sheet.get_all_records() 
        return jsonify({"status": "success", "count": len(records), "data": records}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

# --- 9. (ใหม่!) Routes (Admin Control Panel) ---
@app.route('/admin')
def admin_panel():
    if not session.get('logged_in') or session.get('level') != 'Admin':
        flash('คุณไม่มีสิทธิ์เข้าถึงหน้านี้', 'error')
        return redirect(url_for('dashboard')) 
    try:
        all_staff = staff_sheet.get_all_records()
        all_codes = invite_sheet.get_all_records()
        return render_template('admin_panel.html', 
                               session=session, 
                               staff_list=all_staff, 
                               invite_codes=all_codes)
    except Exception as e:
        flash(f'เกิดข้อผิดพลาดในการโหลด Admin Panel: {e}', 'error')
        return redirect(url_for('dashboard'))

@app.route('/admin/change_level', methods=['POST'])
def change_user_level():
    if not session.get('logged_in') or session.get('level') != 'Admin':
        return redirect(url_for('login_page'))
    try:
        username_to_change = request.form.get('username')
        new_level = request.form.get('level')
        if username_to_change == session.get('username'):
            flash('คุณไม่สามารถเปลี่ยนระดับของตัวเองได้', 'error')
            return redirect(url_for('admin_panel'))
        cell = staff_sheet.find(username_to_change) 
        if not cell:
            flash('ไม่พบผู้ใช้', 'error')
            return redirect(url_for('admin_panel'))
        staff_sheet.update_cell(cell.row, 3, new_level) 
        flash(f'เปลี่ยนระดับของ {username_to_change} เป็น {new_level} สำเร็จ', 'success')
    except Exception as e:
        flash(f'เกิดข้อผิดพลาด: {e}', 'error')
    return redirect(url_for('admin_panel'))

@app.route('/admin/delete_user', methods=['POST'])
def delete_user():
    if not session.get('logged_in') or session.get('level') != 'Admin':
        return redirect(url_for('login_page'))
    try:
        username_to_delete = request.form.get('username')
        if username_to_delete == session.get('username'):
            flash('คุณไม่สามารถลบตัวเองได้', 'error')
            return redirect(url_for('admin_panel'))
        cell = staff_sheet.find(username_to_delete)
        if not cell:
            flash('ไม่พบผู้ใช้', 'error')
            return redirect(url_for('admin_panel'))
        staff_sheet.delete_rows(cell.row)
        flash(f'ลบผู้ใช้ {username_to_delete} สำเร็จ', 'success')
    except Exception as e:
        flash(f'เกิดข้อผิดพลาด: {e}', 'error')
    return redirect(url_for('admin_panel'))

@app.route('/admin/generate_code', methods=['POST'])
def generate_code():
    if not session.get('logged_in') or session.get('level') != 'Admin':
        return redirect(url_for('login_page'))
    try:
        new_code = generate_invite_code()
        new_row = [new_code, 'Available', '', '']
        invite_sheet.append_row(new_row, value_input_option='USER_ENTERED')
        flash(f'สร้างรหัสเชิญใหม่สำเร็จ: {new_code}', 'success')
    except Exception as e:
        flash(f'เกิดข้อผิดพลาด: {e}', 'error')
    return redirect(url_for('admin_panel'))

@app.route('/admin/delete_code', methods=['POST'])
def delete_code():
    if not session.get('logged_in') or session.get('level') != 'Admin':
        return redirect(url_for('login_page'))
    try:
        code_to_delete = request.form.get('code')
        cell = invite_sheet.find(code_to_delete)
        if not cell:
            flash('ไม่พบรหัสเชิญ', 'error')
            return redirect(url_for('admin_panel'))
        invite_sheet.delete_rows(cell.row)
        flash(f'ลบรหัสเชิญ {code_to_delete} สำเร็จ', 'success')
    except Exception as e:
        flash(f'เกิดข้อผิดพลาด: {e}', 'error')
    return redirect(url_for('admin_panel'))


# --- 10. Routes (ระบบ Feedback) ---
@app.route('/feedback')
def feedback_page():
    if not session.get('logged_in'):
        flash('กรุณาเข้าสู่ระบบก่อน', 'error')
        return redirect(url_for('login_page'))
    return render_template('feedback.html', session=session)

@app.route('/feedback_action', methods=['POST'])
def feedback_action():
    if not session.get('logged_in'):
        return redirect(url_for('login_page'))
    if feedback_sheet is None:
        flash('ไม่สามารถเชื่อมต่อระบบ Feedback ได้', 'error')
        return redirect(url_for('feedback_page'))
    try:
        satisfaction = request.form.get('satisfaction')
        ease_of_use = request.form.get('ease_of_use')
        comments = request.form.get('comments', '')
        features = request.form.get('features', '')
        username = session.get('username')
        timestamp = get_current_timestamp()
        new_row = [
            timestamp, username, satisfaction,
            ease_of_use, comments, features
        ]
        feedback_sheet.append_row(new_row, value_input_option='USER_ENTERED')
        flash('ขอบคุณสำหรับข้อเสนอแนะ! ทีมงานจะนำไปปรับปรุงต่อไป', 'success')
        return redirect(url_for('dashboard')) 
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการบันทึก Feedback: {e}")
        flash(f'เกิดข้อผิดพลาด: {e}', 'error')
        return redirect(url_for('feedback_page'))


# --- 11. รันเซิร์ฟเวอร์ ---
if __name__ == '__main__':
    # ใช้พอร์ตที่ Render กำหนด (ถ้ามี) ถ้าไม่มีใช้ 5000
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)