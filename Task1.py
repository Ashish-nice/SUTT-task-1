import pandas as pd

# Read the Excel file
df = pd.read_excel("D:\Ashish\Coding\SUTT\Timetable Workbook - SUTT Task 1.xlsx")

# Convert the DataFrame to JSON
output = df.to_json(orient='records')

# Print the JSON data
print(output)