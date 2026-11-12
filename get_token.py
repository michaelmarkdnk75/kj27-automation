"""Kør én gang på din pc for hver konto:  python get_token.py client_secret.json mail
                                          python get_token.py client_secret.json drive
'mail'  -> log ind som michaelmarkdnk@gmail.com   (Gmail send/læs + kalender læs)
'drive' -> log ind som koegejernmand@gmail.com    (Drive – ejer af KJ27-filerne)
Kræver:  pip install google-auth-oauthlib
"""
import sys
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = {
    "mail": ["https://www.googleapis.com/auth/gmail.send", "https://www.googleapis.com/auth/gmail.readonly",
             "https://www.googleapis.com/auth/calendar.readonly"],
    "drive": ["https://www.googleapis.com/auth/drive"],
}
secret, kind = sys.argv[1], sys.argv[2]
creds = InstalledAppFlow.from_client_secrets_file(secret, SCOPES[kind]).run_local_server(port=0, prompt="consent")
print("\nREFRESH_TOKEN_" + kind.upper() + " =", creds.refresh_token)
