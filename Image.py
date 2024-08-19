from PIL import Image
import os

# กำหนดโฟลเดอร์ที่เก็บไฟล์รูปภาพ
input_folder = r'C:\Users\user\Downloads\ดีลักซ์-20240817T081506Z-001\ดีลักซ์'
output_folder = r'C:\Users\user\OneDrive - Walailak University\Desktop\Img'

# ตรวจสอบว่าโฟลเดอร์ปลายทางมีอยู่หรือไม่ ถ้าไม่มีให้สร้าง
os.makedirs(output_folder, exist_ok=True)

# นับจำนวนไฟล์ที่ถูกแปลง
converted_files_count = 0

# วนลูปผ่านไฟล์ทั้งหมดในโฟลเดอร์
for filename in os.listdir(input_folder):
    # ตรวจสอบว่าเป็นไฟล์รูปภาพหรือไม่ (โดยนามสกุล)
    if filename.lower().endswith(('.jpg', '.jpeg', '.jfif', '.bmp', '.gif', '.tiff', 'webp')):
        try:
            # เปิดไฟล์รูปภาพ
            img_path = os.path.join(input_folder, filename)
            with Image.open(img_path) as img:
                # แปลงเป็น PNG
                output_path = os.path.join(output_folder, os.path.splitext(filename)[0] + '.png')
                img.save(output_path, 'PNG')
                converted_files_count += 1
        except Exception as e:
            print(f"ไม่สามารถแปลงไฟล์ {filename}: {e}")

print(f"การแปลงไฟล์เสร็จสิ้น รวมทั้งหมด {converted_files_count} ไฟล์")
