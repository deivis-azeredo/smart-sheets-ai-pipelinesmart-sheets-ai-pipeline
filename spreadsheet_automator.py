import openpyxl
from groq import Groq
import time
import os

# ========================
# CONFIGURATIONS
# ========================
GROQ_API_KEY = ""
SPREADSHEET_FILE = "reports.xlsx"
SHEET_NAME = "reports"
START_ROW = 2
END_ROW = 10
TARGET_COLUMN = "B"

# ========================
# CONNECT TO LOCAL SPREADSHEET
# ========================
print(f"🔌 Opening local spreadsheet: {SPREADSHEET_FILE}...")

if not os.path.exists(SPREADSHEET_FILE):
    print(f"❌ Error: The file '{SPREADSHEET_FILE}' was not found in this folder.")
    print("Please make sure the Excel file is in the same directory as this script.")
    exit()

# Load the workbook and select the worksheet
workbook = openpyxl.load_workbook(SPREADSHEET_FILE)
sheet = workbook[SHEET_NAME]
print("✅ Loaded successfully!")

# ========================
# Data Processing Loop
# ========================
print(f"\n📖 Reading column {TARGET_COLUMN} from row {START_ROW} to {END_ROW}...")
groq_client = Groq(api_key=GROQ_API_KEY)
total_updated = 0

for current_row in range(START_ROW, END_ROW + 1):
    cell_address = f"{TARGET_COLUMN}{current_row}"
    cell_value = sheet[cell_address].value
    print(f"  Cell {cell_address}: '{cell_value}'")

    if not cell_value or str(cell_value).strip() == "":
        print(f"  ⏭️  Empty, skipping...")
        continue

    print(f"  ✏️  Sending to Groq API...")
    try:
        response = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are an expert English text editor and rewriter. "
                        "Fix all grammatical, spelling, and semantic errors. "
                        "Rewrite in a natural, fluid, humanized way as if written by a real person. "
                        "Reduce plagiarism with original vocabulary and structure. "
                        "Return ONLY the final rewritten text, no explanations, no quotes."
                    )
                },
                {
                    "role": "user",
                    "content": str(cell_value)
                }
            ]
        )
        corrected_text = response.choices[0].message.content.strip()
        
        # Update the cell in memory
        sheet[cell_address] = corrected_text
        print(f"  ✅ Corrected!")
        total_updated += 1
        time.sleep(3) # Pausing to respect API rate limits
        
    except Exception as e:
        print(f"  ❌ Error: {e}")

# Save the changes back to the local file
workbook.save(SPREADSHEET_FILE)
print(f"\n🎉 Done! {total_updated} cells updated and saved to '{SPREADSHEET_FILE}' successfully.")
