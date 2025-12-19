import requests
import time
import json
import os
import urllib3
from dotenv import load_dotenv

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
load_dotenv()

TOKEN_URL = "https://apis-b2b-dev.lowes.com/v1/oauthprovider/oauth2/token"
CHAT_URL  = "https://apis-b2b-dev.lowes.com/llama3/v1/chat/completions"

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
MODEL_NAME = "meta-llama/Meta-Llama-3.1-8B-Instruct"

_access_token = None
_token_expiry = 0

SYSTEM_PROMPT = """
You are an AI assistant for tabular data (CSV / Excel).

RULES:
- ALWAYS return valid JSON
- NEVER write files unless user explicitly says export/save
- Analysis is READ-ONLY

ONLY ONE RESPONSE TYPE HERE:

{
  "type": "analysis",
  "action": {
    "name": "group_count",
    "params": {
      "group_by": "column_name",
      "filter_column": "column_name",
      "operator": "contains | equals",
      "value": "keyword"
    }
  }
}
"""

def get_access_token():
    global _access_token, _token_expiry

    if _access_token and time.time() < _token_expiry:
        return _access_token

    payload = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials"
    }

    r = requests.post(
        TOKEN_URL,
        data=payload,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        verify=False
    )
    r.raise_for_status()

    data = r.json()
    _access_token = data["access_token"]
    _token_expiry = time.time() + data.get("expires_in", 900) - 30
    return _access_token

def ask_llm(user_input, columns):
    token = get_access_token()

    payload = {
        "model": MODEL_NAME,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "system", "content": f"Available columns: {columns}"},
            {"role": "user", "content": user_input}
        ],
        "temperature": 0
    }

    r = requests.post(
        CHAT_URL,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        },
        json=payload,
        verify=False
    )
    r.raise_for_status()

    return json.loads(r.json()["choices"][0]["message"]["content"])
