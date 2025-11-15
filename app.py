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
import json
import requests
from collections import Counter
from threading import Thread 

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


# --- 4. ฟังก์ชันตัวช่วย ---
def generate_new_id():
    code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
    return f"BMA-{code}"

def get_current_timestamp():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def generate_invite_code():
    code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
    return f"INVITE-{code}"

def send_reset_email(username, recipient_email):
    """ ส่งอีเมลแบบ Synchronous (รอจนเสร็จ) เพื่อความชัวร์บน Vercel """
    try:
        token = s.dumps(username, salt='password-reset-salt')
        reset_url = url_for('reset_password_page', token=token, _external=True)
        
        msg = Message("คำขอรีเซ็ตรหัสผ่าน - BMA Link Registry", recipients=[recipient_email])
        msg.body = f"""สวัสดีครับ,\nเราได้รับคำขอรีเซ็ตรหัสผ่านสำหรับ Username: {username}\nคลิกลิงค์เพื่อตั้งรหัสผ่านใหม่: {reset_url}\n(ลิงค์หมดอายุใน 1 ชั่วโมง)"""
        
        mail.send(msg)
        print(f"✅ ส่งอีเมลสำเร็จไปยัง: {recipient_email}")
        return True
    except Exception as e:
        print(f"❌ ส่งอีเมลล้มเหลว: {e}")
        return False


# --- 5. API (Check Username/Invite/Checker) ---
@app.route('/check_username', methods=['POST'])
def check_username():
    if staff_sheet is None: return jsonify({'available': False, 'message': 'DB Error'})
    try:
        data = request.get_json()
        username = data.get('username')
        if not username: return jsonify({'available': False})
        all_usernames = staff_sheet.col_values(1) 
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
    """ API สำหรับ Cron Job (ตรวจสอบลิงค์เสีย) """
    key = request.args.get('key')
    secret_key = os.environ.get('CHECKER_SECRET', 'my_super_secret_checker_key')
    
    if key != secret_key:
        return jsonify({'status': 'error', 'message': 'Unauthorized'}), 401
    if db_sheet is None:
        return jsonify({'status': 'error', 'message': 'Database not connected'}), 500

    print("🚀 (CHECKER) เริ่มตรวจสอบลิงค์...")
    try:
        records = db_sheet.get_all_records()
        updates = []
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}

        # ตรวจสอบ 30 ลิงค์แรกเพื่อกัน Timeout (ปรับเลขได้)
        for i, record in enumerate(records[:30], start=2): 
            url = record.get('URL')
            if not url: continue
            if not url.startswith('http'): url = 'http://' + url

            status_msg = "Unknown"
            try:
                resp = requests.get(url, headers=headers, timeout=3)
                if 200 <= resp.status_code < 300: status_msg = "OK"
                elif resp.status_code == 403: status_msg = "OK"
                else: status_msg = f"{resp.status_code} Error"
            except:
                status_msg = "Error/Timeout"

            updates.append({'range': f'L{i}', 'values': [[status_msg]]})
        
        if updates:
            db_sheet.batch_update(updates, value_input_option='RAW')
            return jsonify({'status': 'success', 'message': f'Checked {len(updates)} links successfully'})
            
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500
    return jsonify({'status': 'success', 'message': 'No links checked'})


# --- 6. Routes (General) ---
@app.route('/')
def home():
    if db_sheet is None: return render_template('index.html', links=[], error="Sheet Error", session=session)
    try:
        all_records = db_sheet.get_all_records()
        links_to_display = [link for link in all_records if link.get('สถานะ') == 'ใช้งาน']
        return render_template('index.html', links=links_to_display, error=None, session=session)
    except Exception as e: return render_template('index.html', links=[], error=str(e), session=session)


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
        staff_list = staff_sheet.get_all_records()
        user_found = next((u for u in staff_list if u['Username'].lower() == username.lower()), None)
        
        if user_found and check_password_hash(user_found['PasswordHash'], password):
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
        position = request.form.get('position')
        department = request.form.get('department')
        email = request.form.get('email')
        phone = request.form.get('phone')
        invite_code = request.form.get('invite_code')

        all_usernames = staff_sheet.col_values(1)
        if username.lower() in [u.lower() for u in all_usernames]:
            flash('Username ซ้ำ', 'error'); return redirect(url_for('register_page'))
        
        code_cell = invite_sheet.find(invite_code)
        if not code_cell or invite_sheet.cell(code_cell.row, 2).value != 'Available':
            flash('รหัสเชิญไม่ถูกต้อง', 'error'); return redirect(url_for('register_page'))

        hashed_password = generate_password_hash(password)
        current_time = get_current_timestamp()
        # Structure A-J
        new_row = [username, hashed_password, 'Users', fullname, position, department, phone, email, current_time, current_time]
        staff_sheet.append_row(new_row, value_input_option='USER_ENTERED')
        
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
            all_staff = staff_sheet.get_all_records()
            user_found = next((u for u in all_staff if u['Username'].lower() == username_input.lower()), None)
            if user_found:
                success = send_reset_email(user_found['Username'], user_found['Email'])
                if not success:
                     flash('เกิดข้อผิดพลาดในการส่งอีเมล', 'error')
                     return redirect(url_for('forgot_password'))
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
                staff_sheet.update_cell(cell.row, 10, get_current_timestamp())
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
        all_staff = staff_sheet.get_all_records()
        user_info = next((u for u in all_staff if u['Username'] == username), None)
        if not user_info: flash(f'ไม่พบผู้ใช้: {username}', 'error'); return redirect(url_for('dashboard'))

        all_links = db_sheet.get_all_records()
        count = sum(1 for link in all_links if link.get('CreatorUsername') == username)
        is_own = (username == session.get('username'))
        return render_template('profile.html', session=session, user=user_info, links_count=count, is_own_profile=is_own)
    except: return redirect(url_for('dashboard'))

@app.route('/edit_profile')
def edit_profile_page():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        all_staff = staff_sheet.get_all_records()
        user_info = next((u for u in all_staff if u['Username'] == session.get('username')), None)
        return render_template('edit_profile.html', session=session, user=user_info)
    except: return redirect(url_for('profile_page'))

@app.route('/edit_profile_action', methods=['POST'])
def edit_profile_action():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        fullname = request.form.get('fullname')
        position = request.form.get('position')
        department = request.form.get('department')
        email = request.form.get('email')
        phone = request.form.get('phone')
        
        cell = staff_sheet.find(session.get('username'))
        if cell:
            staff_sheet.update_cell(cell.row, 4, fullname)
            staff_sheet.update_cell(cell.row, 5, position)
            staff_sheet.update_cell(cell.row, 6, department)
            staff_sheet.update_cell(cell.row, 7, phone)
            staff_sheet.update_cell(cell.row, 8, email)
            staff_sheet.update_cell(cell.row, 10, get_current_timestamp())
            
            session['name'] = fullname
            session['email'] = email
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
    try: return render_template('dashboard.html', session=session, links=db_sheet.get_all_records())
    except: return redirect(url_for('home'))

@app.route('/add')
def add_link_page():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    return render_template('add_link.html', session=session)

@app.route('/add_action', methods=['POST'])
def add_link_action():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        data = {k: request.form.get(k) for k in ['ประเภท','หน่วยงาน','อีเมลผู้รับผิดชอบ','เบอร์โทรติดต่อ','ชื่อลิงก์','URL','หมายเหตุ','สถานะ']}
        new_row = [
            generate_new_id(), data['ประเภท'], data['หน่วยงาน'], data['อีเมลผู้รับผิดชอบ'], 
            data['เบอร์โทรติดต่อ'], data['ชื่อลิงก์'], data['URL'], data['สถานะ'], 
            data['หมายเหตุ'], get_current_timestamp(), session.get('username')
        ]
        db_sheet.append_row(new_row, value_input_option='USER_ENTERED')
        flash('เพิ่มลิงค์สำเร็จ', 'success'); return redirect(url_for('dashboard'))
    except Exception as e: flash(f'Error: {e}', 'error'); return redirect(url_for('add_link_page'))

@app.route('/delete/<link_id>', methods=['POST'])
def delete_link_action(link_id):
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        cell = db_sheet.find(link_id)
        if not cell: return redirect(url_for('dashboard'))
        row_data = db_sheet.row_values(cell.row)
        creator = row_data[10] 
        if session['level'] == 'Admin' or creator == session['username']:
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
        all_links = db_sheet.get_all_records()
        link = next((l for l in all_links if l['ID'] == link_id), None)
        if not link: return redirect(url_for('dashboard'))
        if session['level'] != 'Admin' and link['CreatorUsername'] != session['username']:
            flash('ไม่มีสิทธิ์', 'error'); return redirect(url_for('dashboard'))
        return render_template('edit_link.html', session=session, link=link)
    except: return redirect(url_for('dashboard'))

@app.route('/update_action/<link_id>', methods=['POST'])
def update_link_action(link_id):
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        cell = db_sheet.find(link_id)
        row_vals = db_sheet.row_values(cell.row)
        creator = row_vals[10] 
        if session['level'] != 'Admin' and creator != session['username']:
             flash('ไม่มีสิทธิ์', 'error'); return redirect(url_for('dashboard'))
        
        data = {k: request.form.get(k) for k in ['ประเภท','หน่วยงาน','อีเมลผู้รับผิดชอบ','เบอร์โทรติดต่อ','ชื่อลิงก์','URL','สถานะ','หมายเหตุ']}
        new_vals = [
            link_id, data['ประเภท'], data['หน่วยงาน'], data['อีเมลผู้รับผิดชอบ'], 
            data['เบอร์โทรติดต่อ'], data['ชื่อลิงก์'], data['URL'], data['สถานะ'], 
            data['หมายเหตุ'], get_current_timestamp(), creator
        ]
        range_name = f"A{cell.row}:K{cell.row}"
        db_sheet.update(range_name, [new_vals])
        flash('แก้ไขสำเร็จ', 'success'); return redirect(url_for('dashboard'))
    except: return redirect(url_for('dashboard'))


# --- 10. Routes (Admin & Analytics & Feedback) ---
@app.route('/analytics')
def analytics_page():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    if db_sheet is None: return render_template('analytics.html', session=session, chart_data={})
    try:
        links = db_sheet.get_all_records()
        users = staff_sheet.get_all_records()
        feedback = feedback_sheet.get_all_records()
        
        # Process Data
        cat_counts = Counter(l['ประเภท'] for l in links if l.get('ประเภท'))
        dept_counts = Counter(l['หน่วยงาน'] for l in links if l.get('หน่วยงาน')).most_common(5)
        monthly = {}
        for l in links:
            if l.get('วันที่อัปเดต'):
                try: 
                    m = datetime.datetime.strptime(l['วันที่อัปเดต'].split()[0], '%Y-%m-%d').strftime('%Y-%m')
                    monthly[m] = monthly.get(m, 0) + 1
                except: pass
        sorted_m = sorted(monthly.items())
        
        # Feedback Data
        sat, ease, comments, features = [], [], [], []
        for f in feedback:
            try: sat.append(int(f['SatisfactionScore']))
            except: pass
            try: ease.append(int(f['EaseOfUseScore']))
            except: pass
            if f.get('Comments'): comments.append({'user': f['Username'], 'text': f['Comments']})
            if f.get('FeatureRequest'): features.append({'user': f['Username'], 'text': f['FeatureRequest']})

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
    if not session.get('logged_in') or session.get('level') != 'Admin': return redirect(url_for('dashboard'))
    try:
        return render_template('admin_panel.html', session=session, staff_list=staff_sheet.get_all_records(), invite_codes=invite_sheet.get_all_records())
    except: return redirect(url_for('dashboard'))

# ... (Admin Actions routes - change_level, delete_user etc. - same as before) ...
@app.route('/admin/change_level', methods=['POST'])
def change_user_level():
    if not session.get('logged_in') or session.get('level') != 'Admin': return redirect(url_for('login_page'))
    try:
        user, level = request.form.get('username'), request.form.get('level')
        if user == session.get('username'): flash('เปลี่ยนระดับตัวเองไม่ได้', 'error'); return redirect(url_for('admin_panel'))
        cell = staff_sheet.find(user)
        if cell: staff_sheet.update_cell(cell.row, 3, level); flash('สำเร็จ', 'success')
    except: flash('Error', 'error')
    return redirect(url_for('admin_panel'))

@app.route('/admin/delete_user', methods=['POST'])
def delete_user():
    if not session.get('logged_in') or session.get('level') != 'Admin': return redirect(url_for('login_page'))
    try:
        user = request.form.get('username')
        if user == session.get('username'): flash('ลบตัวเองไม่ได้', 'error'); return redirect(url_for('admin_panel'))
        cell = staff_sheet.find(user)
        if cell: staff_sheet.delete_rows(cell.row); flash('สำเร็จ', 'success')
    except: flash('Error', 'error')
    return redirect(url_for('admin_panel'))

@app.route('/admin/generate_code', methods=['POST'])
def generate_code():
    if not session.get('logged_in') or session.get('level') != 'Admin': return redirect(url_for('login_page'))
    try:
        code = generate_invite_code()
        invite_sheet.append_row([code, 'Available', '', ''], value_input_option='USER_ENTERED')
        flash(f'สร้างรหัส: {code}', 'success')
    except: flash('Error', 'error')
    return redirect(url_for('admin_panel'))

@app.route('/admin/delete_code', methods=['POST'])
def delete_code():
    if not session.get('logged_in') or session.get('level') != 'Admin': return redirect(url_for('login_page'))
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
    try: return jsonify({"status": "success", "data": db_sheet.get_all_records()})
    except: return jsonify({"status": "error"}), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)