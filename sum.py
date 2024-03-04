import pandas as pd
import json
import os
from openpyxl import Workbook
from datetime import datetime

# Read the JSON file
with open('verified_emails.json') as file:
    data = json.load(file)

# Create a list to store the values of the specified key
email = []

# Iterate over each item in the JSON data
for item in data:
    # Get the value of the specified key
    email.append(item['email'])

# Create a DataFrame from the values
df = pd.DataFrame(email)
df.columns = [None]  # Set the column name to None

# Save the DataFrame to an Excel file
current_datetime = datetime.now().strftime("%m%d%H%M")
filename = os.path.join(".", current_datetime + ".xlsx")

df.to_excel(filename, index=False, header=False)
print("Exported successfully to:", filename)