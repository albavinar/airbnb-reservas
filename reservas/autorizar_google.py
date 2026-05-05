#!/usr/bin/env python3
"""Autorización Google OAuth — ejecutar una sola vez."""
import subprocess, warnings, webbrowser
warnings.filterwarnings("ignore")

from google_auth_oauthlib.flow import InstalledAppFlow
from pathlib import Path

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/drive",
]
CREDS = Path(__file__).parent / "credentials.json"
TOKEN = Path(__file__).parent / "token.json"

# Interceptar la llamada al navegador para abrirlo con 'open' de macOS
# y mostrar la URL en pantalla por si falla
def abrir_navegador(url, *args, **kwargs):
    print(f"\nAbriendo navegador...\nSi no se abre, copia esta URL:\n{url}\n")
    subprocess.run(["open", url])
    return True

webbrowser.open = abrir_navegador

flow = InstalledAppFlow.from_client_secrets_file(str(CREDS), SCOPES)
creds = flow.run_local_server(port=8080)   # usa la misma URL internamente
TOKEN.write_text(creds.to_json())
print("✅ Autorización completada. token.json guardado.")
