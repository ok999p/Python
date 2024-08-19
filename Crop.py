from PIL import Image
import os

# โฟลเดอร์ที่เก็บไฟล์รูปภาพ
input_folder = r'C:\Users\user\OneDrive - Walailak University\Desktop\P2'
output_folder = r'C:\Users\user\OneDrive - Walailak University\Desktop\Crop'

# ตรวจสอบว่าโฟลเดอร์ปลายทางมีอยู่หรือไม่ ถ้าไม่มีให้สร้าง
os.makedirs(output_folder, exist_ok=True)

# ฟังก์ชันสำหรับการ crop รูปภาพเป็น 4:5
def crop_to_ratio(image, ratio_width, ratio_height):
    img_width, img_height = image.size
    target_width = img_width
    target_height = int(target_width * ratio_height / ratio_width)

    if target_height > img_height:
        target_height = img_height
        target_width = int(target_height * ratio_width / ratio_height)

    left = (img_width - target_width) / 2
    top = (img_height - target_height) / 2
    right = (img_width + target_width) / 2
    bottom = (img_height + target_height) / 2

    return image.crop((left, top, right, bottom))

# นับจำนวนไฟล์ที่ถูก crop
cropped_files_count = 0

# วนลูปผ่านไฟล์ทั้งหมดในโฟลเดอร์
for filename in os.listdir(input_folder):
    if filename.lower().endswith(('.jpg', '.jpeg', '.jfif', '.bmp', '.gif', '.tiff', '.png')):
        try:
            img_path = os.path.join(input_folder, filename)
            with Image.open(img_path) as img:
                # Crop รูปภาพให้เป็นสัดส่วน 4:5
                cropped_img = crop_to_ratio(img, 4, 5)
                # บันทึกรูปภาพที่ถูก crop แล้ว
                output_path = os.path.join(output_folder, filename)
                cropped_img.save(output_path)
                cropped_files_count += 1
        except Exception as e:
            print(f"ไม่สามารถ crop ไฟล์ {filename}: {e}")

print(f"การ crop ไฟล์เสร็จสิ้น รวมทั้งหมด {cropped_files_count} ไฟล์")
