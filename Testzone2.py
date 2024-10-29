import pandas as pd
import re

# Load the Excel file
input_file = r'C:\Users\user\Downloads\ผังการจัดที่นังวันที่8 มิย 66.xlsx'
df = pd.read_excel(input_file, header=None)

# Create a list to store the seat data with the zone
seat_data = []

# Define a pattern for seat numbers (A01 - Z30)
seat_pattern = re.compile(r'^[A-Z]\d{2}$')

# Iterate through each column to find seat numbers and their zones
for col in df.columns:
    # Get the zone number from row 1 and 15
    zone_row_1 = df.iloc[0, col]
    zone_row_15 = df.iloc[14, col]

    # Iterate through rows 1 to 12 for zone_row_1
    for row in range(1, 13):  # From row 1 to row 12
        seat_number = df.iloc[row, col]
        if isinstance(seat_number, str) and seat_pattern.match(seat_number):
            # Add the zone and seat number to the list from zone_row_1
            seat_data.append([zone_row_1, seat_number])
    
    # Iterate through rows 15 to 30 for zone_row_15
    for row in range(15, 31):  # From row 15 to row 30
        seat_number = df.iloc[row, col]
        if isinstance(seat_number, str) and seat_pattern.match(seat_number):
            # Add the zone and seat number to the list from zone_row_15
            seat_data.append([zone_row_15, seat_number])

# Create a new DataFrame for the output
output_df = pd.DataFrame(seat_data, columns=['โซน', 'หมายเลขที่นั่ง'])

# Save the output to a new Excel file
output_file = r'C:\Users\user\OneDrive - Walailak University\Desktop\New folder\Database_ThaiburiSeat.xlsx'
output_df.to_excel(output_file, index=False)

# Verify if the new file contains all seat numbers (A01-Z30)
expected_seats = [f"{chr(i)}{str(j).zfill(2)}" for i in range(ord('A'), ord('Z')+1) for j in range(1, 31)]

# Check if all expected seats are in the new output
missing_seats = set(expected_seats) - set(output_df['หมายเลขที่นั่ง'])

if not missing_seats:
    print(f"ไฟล์ใหม่สร้างสำเร็จและมีหมายเลขที่นั่งครบถ้วน: {output_file}")
else:
    print(f"ไฟล์ใหม่สร้างสำเร็จ แต่มีหมายเลขที่นั่งที่ขาดหาย: {', '.join(missing_seats)}")
