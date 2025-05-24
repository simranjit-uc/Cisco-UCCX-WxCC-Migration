import os, requests, json, webbrowser, time, socket
from requests_oauthlib import OAuth2Session


client_ID = os.getenv("WxCC_CLIENT_ID")
client_Secret = os.getenv("WxCC_CLIENT_SECRET")
auth_Base_URL = os.getenv("WxCC_AUTH_URL")
token_URL = os.getenv("WxCC_TOKEN_URL")
redirect_URI = os.getenv("WxCC_REDIRECT_URI")

oauth = OAuth2Session(client_ID, redirect_uri=redirect_URI)

def oauth_Flow():
    # Get the authorization URL
    auth_URL, state = oauth.authorization_url(auth_Base_URL)

    # Save the state value in a file
    with open('state.txt', 'w') as f:
        f.write(state)

    # You can open the authorization URL in the default web browser as well.
    # webbrowser.open(authorization_url)

    # But I prefer Firefox but it's not my default browser. So, I've provided a path to its EXE file
    webbrowser.Mozilla(r"C:\Program Files\Mozilla Firefox\firefox.exe").open(auth_URL)

    # Create a socket to listen for incoming messages
    server_Socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server_Socket.bind(('localhost', 5001))                                 # Client port
    server_Socket.listen(1)

    print("Waiting for the authorization code from the server.....")

    client_Socket, client_Address = server_Socket.accept()
    data = client_Socket.recv(1024).decode()
    auth_Code, state = data.split('|')
    state = state.strip()
    server_Socket.close()
    return auth_Code, state


def get_Access_Token():
    # Grab Auth code and state values
    auth_Code, state = oauth_Flow()

    # Send OAuth request to the token URL to exchange Auth  code for a token
    token = oauth.fetch_token(token_URL, client_secret=client_Secret, authorization_response=f"{redirect_URI}?code={auth_Code}&state={state}", verify=False)
    token_Value = token['access_token']
    refresh_Value = token['refresh_token']
    expiry_Value = token['expires_in']
    token_Expiry = time.time() + expiry_Value

    # Compare token expiry time with current time. If it's exceeded then send a request to refresh the token.
    if time.time() > token_Expiry:
        print("Access token expired. Refreshing...")
        access_token, refresh_token, token_expiry = refresh_Access_Token(refresh_Value)
        return access_token, refresh_token, token_expiry
    else:
        return token_Value, refresh_Value


def refresh_Access_Token(refresh_token):
    data = {
        "grant_type": "refresh_token",
        "client_id": client_ID,
        "client_secret": client_Secret,
        "refresh_token": refresh_token
    }
    response = requests.post(token_URL,data=data)
    response_data = response.json()
    new_Access_Token = response_data['access_token']
    new_Refresh_Token = response_data['refresh_token']
    expires_in = response_data['expires_in']
    token_Expiry = time.time() + expires_in
    return new_Access_Token, new_Refresh_Token, token_Expiry


if __name__ == '__main__':
    get_Access_Token()