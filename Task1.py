import openpyxl
import pandas as pd
import json

workbook = openpyxl.load_workbook(r"D:\Ashish\Coding\SUTT\Timetable Workbook - SUTT Task 1.xlsx")
data = []

for i in range(1, 7):
    worksheet = workbook[f'S{i}']
    course_data = {
        "course_code": worksheet['B4'].value,
        "course_title": worksheet['C4'].value,
        "credits": {
            "L": worksheet['D4'].value,
            "P": worksheet['E4'].value,
            "U": worksheet['F4'].value
        },
        "sections": []
    }
    
    row = 4
    last_section_number = None
    last_room = None
    last_timing = None
    instructors = []
    section_type = None  # Initialize section_type
    
    while worksheet[f'H{row}'].value is not None or worksheet[f'I{row}'].value is not None or worksheet[f'J{row}'].value is not None:
        section_number = worksheet[f'G{row}'].value
        room = worksheet[f'I{row}'].value
        timing = worksheet[f'J{row}'].value
        instructor = worksheet[f'H{row}'].value  # Corrected to read from H column
        
        if section_number is not None:
            if last_section_number is not None and last_section_number != section_number:
                # Append the previous section data before resetting
                section_data = {
                    "section_type": section_type,
                    "section_number": last_section_number,
                    "instructors": instructors,
                    "room": last_room,
                    "timing": last_timing
                }
                course_data["sections"].append(section_data)
                # Reset instructors list when moving to a new section
                instructors = []
            
            last_section_number = section_number
            if section_number.startswith('L'):
                section_type = 'Lecture'
            elif section_number.startswith('T'):
                section_type = 'Tutorial'
            elif section_number.startswith('P'):
                section_type = 'Laboratory'
        
        if room is not None:
            last_room = room
        else:
            room = last_room
        
        if timing is not None:
            last_timing = timing
        else:
            timing = last_timing
        
        if instructor is not None:
            instructors.append(instructor)
        
        row += 1
    
    # Append the last section data
    if last_section_number is not None:
        section_data = {
            "section_type": section_type,
            "section_number": last_section_number,
            "instructors": instructors,
            "room": last_room,
            "timing": last_timing
        }
        course_data["sections"].append(section_data)
    
    data.append(course_data)

# Convert the data to a pandas DataFrame
df = pd.DataFrame(data)

# Convert the DataFrame to JSON
json_data = df.to_json(orient='records', indent=4)

# Save the JSON data to a file
with open('output.json', 'w') as json_file:
    json_file.write(json_data)