# A script made in python for linking with OneDrive, get pictures and return a HTML bootstrap carousel to use when needed, this way is much faster then manually get URLs one by one.
# Made by Gyuji.

#Notes: Kinda dissapointed that this was made to be used for ebay's html custom page and it doesn't even let JS scripts load in there, total bummer.

import json
import time
import tkinter as tk
from requests_oauthlib import OAuth2Session
from oauthlib.oauth2 import TokenExpiredError
from tkinter import simpledialog
import webbrowser
from urllib.parse import quote

# Microsoft Azure Graph API credentials, permissions must be granted to fetch data.
client_id = ''
client_secret = ''
redirect_uri = ''

token_file = 'token.json'  # Make a file or retrieve data for json information.

# Custom information to be fetched, folder "Ebay pictures" with the SKU inside.
SKU = 'Z22243'
folder_name = 'Ebay pictures'
encoded_folder_name = quote(folder_name)

# OneDrive API URLs
authorize_url = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize'
token_url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
api_url = f'https://graph.microsoft.com/v1.0/me/drive/root:/{encoded_folder_name}/{SKU}:/children'

# Microsoft API scope, this need to be granted permission to use.
scope = ['Files.Read']

# This tries to fetch data from a token.json file, if not, returns an error.
def get_token():
    try:
        with open(token_file, 'r') as file:
            token = json.load(file)
        return token
    except FileNotFoundError:
        return None

# Makes the session authentication.
token = get_token()
oauth = OAuth2Session(client_id, redirect_uri=redirect_uri, scope=scope, token=token)

# I stole this code, no clue how it refreshes and i don't know if it works, it doesn't break the program tho.
def refresh_token():
    try:
        token = oauth.refresh_token(token_url, client_id=client_id, client_secret=client_secret)
        with open(token_file, 'w') as file:
            json.dump(token, file)
    except Exception as e:
        print(f"Token refresh failed: {e}")

# If the token is not available or has expired, get a new one
if not token or token.get('expires_at') < time.time():
    authorization_url, state = oauth.authorization_url(authorize_url)
    tk.Tk().withdraw()  # Hide the main window

    # From now on it's not related to refreshing tokens, it just need the redirect uri to proceed, you need to login in ur microsoft account and then get the URL WITH the https part.
    webbrowser.open(authorization_url)
    redirect_response = simpledialog.askstring("Authorization", f"Please go to the opened web browser and authorize access. Paste the full redirect URL here:")

    # Fetch the access token
    try:
        oauth.fetch_token(token_url, authorization_response=redirect_response, client_secret=client_secret)
    except TokenExpiredError:
        refresh_token()

    # Save the token for future use
    with open(token_file, 'w') as file:
        json.dump(oauth.token, file)

# Make a request and check status.
try:
    response = oauth.get(api_url)
    response.raise_for_status()
    data = response.json()
except TokenExpiredError:
    refresh_token()
    response = oauth.get(api_url)
    response.raise_for_status()
    data = response.json()

#this variable will store all the html to be used when needed, in this case, it turns all the pictures into URLs and dynamically fills a bootstrap caroulsel.
html_content = ''

if 'error' in data:
    html_content += 'Error: {}'.format(data["error"]["message"])
else:
    html_content += '<style>'
    html_content += '.carousel-control-prev,'
    html_content += '.carousel-control-next {'
    html_content += '  filter: invert(100%);'
    html_content += '}'
    html_content += '</style>'

    html_content += '<div class="container mx-auto">'
    html_content += '  <div class="row">'
    html_content += '    <div class="col-8">'
    html_content += '      <style>.carousel-control-prev,.carousel-control-next {  filter: invert(100%);}</style>'
    html_content += '      <div id="carouselExampleIndicators" class="carousel slide">'
    html_content += '        <div class="carousel-indicators">'

    for i, item in enumerate(data['value']):
        html_content += '<button type="button" data-bs-target="#carouselExampleIndicators" data-bs-slide-to="{}" {}aria-label="Slide {}"></button>'.format(i, 'class="active" ' if i == 0 else '', i + 1)

    html_content += '        </div>'
    html_content += '        <div class="carousel-inner">'

    for i, item in enumerate(data['value']):
        if 'image' in item['file']['mimeType']:
            name = item['name']
            web_url = item['webUrl']
            direct_download_url = item.get('@microsoft.graph.downloadUrl')

            html_content += '<div class="carousel-item {}">'.format('active' if i == 0 else '')
            if direct_download_url:
                html_content += '  <img src="{}" class="d-block w-100" alt="{}">'.format(direct_download_url, name)
            else:
                html_content += '  <img src="{}" class="d-block w-100" alt="{}">'.format(web_url, name)
            html_content += '</div>'

    html_content += '        </div>'
    html_content += '        <button class="carousel-control-prev" type="button" data-bs-target="#carouselExampleIndicators" data-bs-slide="prev">'
    html_content += '          <span class="carousel-control-prev-icon" aria-hidden="true"></span>'
    html_content += '          <span class="visually-hidden">Previous</span>'
    html_content += '        </button>'
    html_content += '        <button class="carousel-control-next" type="button" data-bs-target="#carouselExampleIndicators" data-bs-slide="next">'
    html_content += '          <span class="carousel-control-next-icon" aria-hidden="true"></span>'
    html_content += '          <span class="visually-hidden">Next</span>'
    html_content += '        </button>'
    html_content += '      </div>'
    html_content += '    </div>'

    # This column is just for a small "gallery"
    html_content += '    <div class="col-4" style="max-height: 500px; overflow-y: auto;">'
    for i, item in enumerate(data['value']):
        if 'image' in item['file']['mimeType']:
            direct_download_url = item.get('@microsoft.graph.downloadUrl')
            html_content += '      <a data-bs-target="#carouselExampleIndicators" href="#" data-bs-slide-to="{}" {}aria-label="Slide {}">'.format(i, 'class="active" ' if i == 0 else '', i + 1)
            html_content += '        <img src="{}" style="width: 100px; height: 100px; object-fit: cover;" alt="{}">'.format(direct_download_url, name)
            html_content += '      </a>'

    html_content += '    </div>'
    html_content += '  </div>'
    html_content += '</div>'

#Then you コピペ the content in the terminal and put wherever u like.
print(html_content)
