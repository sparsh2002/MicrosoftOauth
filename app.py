import os
import msal
import requests
from flask import Flask, redirect, request, session, url_for
from dotenv import load_dotenv
import os

load_dotenv()

app = Flask(__name__)
app.secret_key = "mysecretkey"
PORT = os.getenv('PORT')

# Microsoft App Credentials (Replace these)
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
TENANT_ID = os.getenv('TENANT_ID')
# AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
AUTHORITY = "https://login.microsoftonline.com/common"
REDIRECT_URI = "http://localhost:5001/callback"
SCOPES = ['User.Read' , 'Mail.Read' ,'Mail.ReadWrite'] 

print(PORT ,CLIENT_ID)

# MSAL App instance
msal_app = msal.ConfidentialClientApplication(CLIENT_ID, CLIENT_SECRET, authority=AUTHORITY)

@app.route("/")
def home():
    return '<a href="/login">Login with Microsoft</a>'

@app.route("/login")
def login():
    auth_url = msal_app.get_authorization_request_url(
        SCOPES, 
        redirect_uri=REDIRECT_URI, 
        prompt='select_account'
        )
    return redirect(auth_url)

@app.route("/callback")
def callback():
    code = request.args.get("code")
    if not code:
        return "Login failed", 400

    token_response = msal_app.acquire_token_by_authorization_code(code, SCOPES, redirect_uri=REDIRECT_URI)
    print(token_response['refresh_token'])
    if "access_token" in token_response:
        session["user"] = {
                "access_token": token_response["access_token"],
                "refresh_token": token_response.get("refresh_token"),  # Store refresh token
                "expires_in": token_response["expires_in"]
            }
        return redirect(url_for("profile"))
    return "Login failed", 400

@app.route("/profile")
def profile():
    if "user" not in session:
        return redirect(url_for("login"))

    access_token = session["user"]["access_token"]
    refresh_token = session["user"]["refresh_token"]
    print(access_token)
    print(refresh_token)
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get("https://graph.microsoft.com/v1.0/me", headers=headers)
    
    if response.status_code == 200:
        return response.json()
    return "Failed to fetch profile", response.status_code

@app.route("/refresh")
def refresh_access_token():
    if "user" not in session or "refresh_token" not in session["user"]:
        return None

    token_response = msal_app.acquire_token_by_refresh_token(
        session["user"]["refresh_token"], SCOPES
    )

    if "access_token" in token_response:
        session["user"]["access_token"] = token_response["access_token"]
        session["user"]["refresh_token"] = token_response.get("refresh_token", session["user"]["refresh_token"])
        session["user"]["expires_in"] = token_response["expires_in"]
        return session["user"]["access_token"]
    
    return None  # Refresh token expired or invalid


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("home"))


@app.route('/fetch_emails')
def get_messages():
    # if "user" not in session:
    #     return redirect(url_for("login"))

    # access_token = session["user"]["access_token"]
    access_token = request.get_json()['access_token']
    # refresh_token = session["user"]["refresh_token"]
    # print(refresh_token)
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(f"https://graph.microsoft.com/v1.0/me/messages?$top={request.args.get('top')}", headers=headers)
    print(response)
    if response.status_code == 200:
        return response.json()
    return "Failed to fetch mails", response.status_code

if __name__ == "__main__":
    app.run(port=5001 , debug=True)
