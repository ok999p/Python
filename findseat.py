import pandas as pd

# โหลดไฟล์ Excel
file_path = r'C:\Users\user\Downloads\Database_ThaiburiSeat.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

# ฟังก์ชันสำหรับจัดที่นั่งตามแถว
def categorize_seats(df):
    seat_dict = {}
    
    for index, row in df.iterrows():
        seat = row['หมายเลขที่นั่ง']
        row_letter = seat[0]  # ใช้ตัวอักษรแรกเพื่อหาว่าเป็นแถวอะไร
        
        if row_letter not in seat_dict:
            seat_dict[row_letter] = []
        
        seat_dict[row_letter].append(seat)
    
    # จัดเรียงที่นั่งในแต่ละแถวจากมากไปน้อย
    for key in seat_dict:
        # เรียงจากมากไปน้อยโดยใช้การแปลงตัวเลขที่นั่งเป็น int
        seat_dict[key].sort(key=lambda x: int(x[1:]), reverse=True)
    
    # จัดเรียงตัวอักษรแถวจาก A-Z
    seat_dict = dict(sorted(seat_dict.items()))
    
    return seat_dict

# จัดประเภทที่นั่งตามแถว
categorized_seats = categorize_seats(df)

# แสดงผลลัพธ์ในรูปแบบโค้ดที่คุณต้องการ พร้อม \n หลังจากจบแต่ละแถว
output = "const rowSeats = {\n"
for row, seats in categorized_seats.items():
    output += f"    '{row}': {seats},\n\n"  # เพิ่ม \n หลังจากแถวจบเพื่อสร้างบรรทัดใหม่
output += "};"

print(output)
