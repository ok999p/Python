import os
import pandas as pd

# ระบุโฟลเดอร์ที่เก็บไฟล์รูปภาพ
image_folder = r"C:\Users\user\Downloads\รูปผู้สมัครสภานักศึกษา 67-20240816T062907Z-001\รูปผู้สมัครสภานักศึกษา 67"

# ระบุไฟล์ Excel ที่ต้องการตรวจสอบ
excel_file = r"C:\Users\user\Downloads\รายชื่อผู้สมัครสภานักศึกษา-67.xlsx"

# ตรวจสอบว่ามีไฟล์ Excel อยู่หรือไม่
if not os.path.exists(excel_file):
    print(f"ไม่พบไฟล์ Excel ที่ {excel_file}")
else:
    # อ่านข้อมูลจากไฟล์ Excel
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการอ่านไฟล์ Excel: {e}")
        exit()

    # ตรวจสอบว่าคอลัมน์รหัสนักศึกษาชื่ออะไร (ในภาพตัวอย่างคือ "รหัสนักศึกษา")
    student_id_column = 'รหัสนักศึกษา'

    # ตรวจสอบว่าคอลัมน์ 'รหัสนักศึกษา' มีอยู่ใน DataFrame หรือไม่
    if student_id_column not in df.columns:
        print(f"ไม่พบคอลัมน์ '{student_id_column}' ในไฟล์ Excel")
    else:
        # สร้างเซ็ตของชื่อไฟล์รูปภาพที่มีอยู่
        image_files = set(f.split('.')[0] for f in os.listdir(image_folder) if f.endswith(('.jpg', '.jpeg', '.jfif', '.png', '.bmp', '.gif', '.tiff')))

        # สร้างคอลัมน์ใหม่ใน DataFrame เพื่อตรวจสอบว่ามีรูปภาพหรือไม่
        df['Check'] = df[student_id_column].apply(lambda x: 1 if str(x) in image_files else 0)

        # ระบุที่อยู่ไฟล์ Excel ที่ต้องการบันทึก
        output_excel_file = r"C:\Users\user\OneDrive - Walailak University\Desktop\student_data_with_images.xlsx"

        # บันทึก DataFrame กลับไปที่ไฟล์ Excel
        try:
            df.to_excel(output_excel_file, index=False)
            print("การตรวจสอบและบันทึกไฟล์เสร็จสิ้น")
        except Exception as e:
            print(f"เกิดข้อผิดพลาดในการบันทึกไฟล์ Excel: {e}")
