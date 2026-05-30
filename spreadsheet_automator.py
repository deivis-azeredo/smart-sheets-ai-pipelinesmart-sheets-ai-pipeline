import gspread
from google.oauth2.service_account import Credentials
from groq import Groq
import time

# ========================
# CONFIGURATIONS
# ========================
GROQ_API_KEY = "YOUR_API_KEY_HERE"
SPREADSHEET_FILE = "reports.xlsx"
SHEET_NAME = "reports"
START_ROW = 2
END_ROW = 10
TARGET_COLUMN = "B"

# ========================
# CONNECT GOOGLE SHEETS
# ========================
print("🔌 Connecting to Google Sheets...")
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
# Assuming credentials.json is in the same folder
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
gc = gspread.authorize(creds)

# Using the standard variables we defined earlier
sheet = gc.open_by_key(SPREADSHEET_FILE).worksheet(SHEET_NAME)
print("✅ Connected!")

# ========================
# Data Processing Loop
# ========================
print(f"\n📖 Reading column {TARGET_COLUMN} from row {START_ROW} to {END_ROW}...")
groq_client = Groq(api_key=GROQ_API_KEY)
total_updated = 0

for current_row in range(START_ROW, END_ROW + 1):
    cell_address = f"{TARGET_COLUMN}{current_row}"
    cell_value = sheet.acell(cell_address).value
    print(f"  Cell {cell_address}: '{cell_value}'")

    if not cell_value or cell_value.strip() == "":
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
                    "content": cell_value
                }
            ]
        )
        corrected_text = response.choices[0].message.content.strip()
        sheet.update_acell(cell_address, corrected_text)
        print(f"  ✅ Saved!")
        total_updated += 1
        time.sleep(3) # Pausing to respect API rate limits
        
    except Exception as e:
        print(f"  ❌ Error: {e}")

print(f"\n🎉 Done! {total_updated} cells updated successfully.")
