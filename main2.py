import requests
import os
import datetime
import openpyxl

# Your personal data. Used by Nutritionix to calculate calories.
GENDER = "male"
WEIGHT_KG = 84
HEIGHT_CM = 180
AGE = 32

# Nutritionix APP ID and API Key. Actual values are stored as environment variables.
APP_ID = os.environ["ENV_NIX_APP_ID"]
API_KEY = os.environ["ENV_NIX_API_KEY"]

exercise_endpoint = "https://trackapi.nutritionix.com/v2/natural/exercise"

exercise_text = input("Tell me which exercises you did: ")

# Nutritionix API Call
headers = {
    "x-app-id": APP_ID,
    "x-app-key": API_KEY,
}

parameters = {
    "query": exercise_text,
    "gender": GENDER,
    "weight_kg": WEIGHT_KG,
    "height_cm": HEIGHT_CM,
    "age": AGE
}

response = requests.post(exercise_endpoint, json=parameters, headers=headers)
result = response.json()
print(f"Nutritionix API call: \n {result} \n")

# Adding date and time
today_date = datetime.datetime.now().strftime("%Y-%m-%d")
now_time = datetime.datetime.now().strftime("%H:%M:%S")

# Append exercise data to Excel file
excel_file = "exercise_data.xlsx"
wb = openpyxl.load_workbook(excel_file)
sheet = wb.active
for exercise in result["exercises"]:
    row = [today_date, now_time, exercise["name"].title(), exercise["duration_min"], exercise["nf_calories"]]
    sheet.append(row)
wb.save(excel_file)
print("Exercise data appended to Excel file")
