import pandas as pd
import re

# Load the Excel file into a DataFrame
file_path = r'C:\Users\user\Downloads\ผังการจัดที่นังวันที่8 มิย 66.xlsx'
df = pd.read_excel(file_path)

# Function to check if a string is a valid seat number (e.g., A01, B01, etc.)
def is_seat_number(value):
    return bool(re.match(r'^[A-Z]\d{2}$', value))

# Filter out only the valid seat numbers
valid_seat_list = []

for col in df.columns:
    seat_numbers = df[col].dropna().astype(str).tolist()
    for seat in seat_numbers:
        if is_seat_number(seat):
            valid_seat_list.append({'Seat': seat})

# Convert to a DataFrame for export
valid_seat_df = pd.DataFrame(valid_seat_list)

# Save to a new Excel file with only seat numbers
output_file_seats = r'C:\Users\user\OneDrive - Walailak University\Desktop\New folder\valid_seat_list.xlsx'
valid_seat_df.to_excel(output_file_seats, index=False)

print("File saved at:", output_file_seats)
