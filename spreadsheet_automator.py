import gspread
from google.oauth2.service_account import Credentials
from groq import Groq
import time

# ========================
# CONFIGURAÇÕES
# ========================
GROQ_API_KEY = " ■■■■■■■■■■■Due to restrictions my API won't be shown"
SPREADSHEET_ID = "reports.xlsx"
SHEET_NAME = "reports"
LINHA_INICIO = 2
LINHA_FIM = 10
COLUNA = "B"

# ========================
# CONECTAR GOOGLE SHEETS
# ========================
print("🔌 Conectando ao Google Sheets...")
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
gc = gspread.authorize(creds)
sheet = gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
print("✅ Conectado!")

# ========================
# LER TODAS AS CÉLULAS
# ========================
print(f"\n📖 Lendo coluna {COLUNA} linhas {LINHA_INICIO} até {LINHA_FIM}...")
client_groq = Groq(api_key=GROQ_API_KEY)
total = 0

for linha in range(LINHA_INICIO, LINHA_FIM + 1):
    celula = f"{COLUNA}{linha}"
    valor = sheet.acell(celula).value
    print(f"  Célula {celula}: '{valor}'")

    if not valor or valor.strip() == "":
        print(f"  ⏭️  Vazia, pulando...")
        continue

    print(f"  ✏️  Enviando para o Groq...")
    try:
        response = client_groq.chat.completions.create(
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
                    "content": valor
                }
            ]
        )
        texto_corrigido = response.choices[0].message.content.strip()
        sheet.update_acell(celula, texto_corrigido)
        print(f"  ✅ Gravado!")
        total += 1
        time.sleep(3)
    except Exception as e:
        print(f"  ❌ Erro: {e}")

print(f"\n🎉 Concluído! {total} células atualizadas.")
