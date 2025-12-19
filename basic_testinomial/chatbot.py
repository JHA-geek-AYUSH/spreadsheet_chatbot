#pip install python-dotenv
import requests
import time
import json
import warnings
import os
from dotenv import load_dotenv

warnings.filterwarnings("ignore")  # suppress SSL warnings for internal APIs

# --------------------------------------------------
# LOAD ENV VARIABLES
# --------------------------------------------------
load_dotenv()   # reads .env file

# --------------------------------------------------
# CONFIG (MATCH YOUR POSTMAN SETUP)
# --------------------------------------------------
TOKEN_URL = "https://apis-b2b-dev.lowes.com/v1/oauthprovider/oauth2/token"
CHAT_URL  = "https://apis-b2b-dev.lowes.com/llama3/v1/chat/completions"

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

MODEL_NAME = "meta-llama/Meta-Llama-3.1-8B-Instruct"

# --------------------------------------------------
# TOKEN CACHE (AUTO-REFRESH)
# --------------------------------------------------
_access_token = None
_token_expiry = 0  # epoch seconds


def get_access_token():
    """
    Returns a valid access token.
    Automatically refreshes if expired.
    """
    global _access_token, _token_expiry

    # Reuse token if still valid
    if _access_token and time.time() < _token_expiry:
        return _access_token

    payload = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials"
    }

    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }

    resp = requests.post(
        TOKEN_URL,
        data=payload,
        headers=headers,
        verify=False
    )

    resp.raise_for_status()
    data = resp.json()

    _access_token = data["access_token"]
    expires_in = data.get("expires_in", 900)

    # Refresh 30 seconds before actual expiry
    _token_expiry = time.time() + expires_in - 30

    print("🔑 Token refreshed")

    return _access_token


# --------------------------------------------------
# CHAT CONTEXT (THIS IS YOUR MEMORY)
# --------------------------------------------------
conversation = [
    {
        "role": "system",
        "content": "You are a helpful assistant. Answer clearly and concisely."
    }
]


# --------------------------------------------------
# CHAT FUNCTION
# --------------------------------------------------
def chat_with_llm():
    token = get_access_token()

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Accept": "application/json"
    }

    payload = {
        "model": MODEL_NAME,
        "messages": conversation,
        "temperature": 0,
        "top_p": 1,
        "max_tokens": 200,
        "presence_penalty": 0,
        "frequency_penalty": 0,
        "user": "generic-chatbot"
    }

    response = requests.post(
        CHAT_URL,
        headers=headers,
        data=json.dumps(payload),
        verify=False
    )

    # If token expired unexpectedly, retry once
    if response.status_code == 401:
        print("⚠ Token expired mid-call, retrying...")
        get_access_token()
        headers["Authorization"] = f"Bearer {_access_token}"
        response = requests.post(
            CHAT_URL,
            headers=headers,
            data=json.dumps(payload),
            verify=False
        )

    response.raise_for_status()

    return response.json()["choices"][0]["message"]["content"]


# --------------------------------------------------
# MAIN CHAT LOOP
# --------------------------------------------------
if __name__ == "__main__":
    print("💬 Generic LLM Chatbot (type 'exit' to quit)\n")

    while True:
        user_input = input("You: ").strip()

        if user_input.lower() == "exit":
            break

        conversation.append({
            "role": "user",
            "content": user_input
        })

        try:
            reply = chat_with_llm()
        except Exception as e:
            print("❌ Error:", e)
            continue

        conversation.append({
            "role": "assistant",
            "content": reply
        })

        print("Bot:", reply, "\n")