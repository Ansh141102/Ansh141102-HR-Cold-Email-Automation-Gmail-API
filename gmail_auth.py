from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

def authenticate_gmail():
    creds = None

    # Check whether we already have a saved login token
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file(
            "token.json",
            SCOPES
        )

    # If there is no valid login, start OAuth authentication
    if not creds or not creds.valid:

        # Refresh an expired token if possible
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())

        else:
            if not os.path.exists("credentials.json"):
                raise FileNotFoundError(
                    "credentials.json was not found."
                )

            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json",
                SCOPES
            )

            creds = flow.run_local_server(port=0)

        # Save the authorization for future runs
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    print("Gmail authentication successful!")
    print("token.json has been created.")

if __name__ == "__main__":
    authenticate_gmail()