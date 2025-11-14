# hash_gen.py
from werkzeug.security import generate_password_hash

# ⚠️ เปลี่ยน 'super_password_123' เป็นรหัสผ่านที่คุณต้องการ
password_to_hash = 'admin2' 
hashed_password = generate_password_hash(password_to_hash)

print("นี่คือ Hash ของคุณ:")
print(hashed_password)