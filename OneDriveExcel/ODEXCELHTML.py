''' 
This script is basically a mix of the OneDrive API picture fetcher plus the excel to HTML,
it just wraps everything into a single program, the other programs are commented for more info, 
but i will just point the small changes in this one, look the other scripts for detailed walkthrough.
'''
# Made by Gyuji.

import json
import time
from requests_oauthlib import OAuth2Session
from oauthlib.oauth2 import TokenExpiredError
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import scrolledtext
import webbrowser
from tkinter import simpledialog
from urllib.parse import quote

def run_onedrive(sku): #This will run the onedrive API PLUS the Excel to HTML one after the other, it will ask for auth.
    
    client_id = ''
    client_secret = ''
    redirect_uri = 'https://localhost:5000/callback'
    token_file = 'token.json'

    # Custom folders inside OneDrive, these can change anytime.
    folder_name = 'Ebay pictures'
    encoded_folder_name = quote(folder_name)

    authorize_url = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize'
    token_url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
    api_url = f'https://graph.microsoft.com/v1.0/me/drive/root:/{encoded_folder_name}/{sku}:/children'
    #api_url = f'https://graph.microsoft.com/v1.0/me/drive/root:/{sku}:/children'
    #f'https://graph.microsoft.com/v1.0/me/drive/root:/LDE:/IN%STOCK:/{sku}:/children' - Possible path to the In Stock files.

    scope = ['Files.Read']

    def get_token():
        try:
            with open(token_file, 'r') as file:
                token = json.load(file)
            return token
        except FileNotFoundError:
            return None

    token = get_token()
    oauth = OAuth2Session(client_id, redirect_uri=redirect_uri, scope=scope, token=token)

    def refresh_token():
        try:
            token = oauth.refresh_token(token_url, client_id=client_id, client_secret=client_secret)
            with open(token_file, 'w') as file:
                json.dump(token, file)
        except Exception as e:
            print(f"Token refresh failed: {e}")

    if not token or token.get('expires_at') < time.time():
        authorization_url, state = oauth.authorization_url(authorize_url)


        tk.Tk().withdraw()
        webbrowser.open(authorization_url)
        redirect_response = simpledialog.askstring("Authorization", f"Please go to the opened web browser and authorize access. Paste the full redirect URL here:")

        try:
            oauth.fetch_token(token_url, authorization_response=redirect_response, client_secret=client_secret)
        except TokenExpiredError:
            refresh_token()

        with open(token_file, 'w') as file:
            json.dump(oauth.token, file)

    try:
        response = oauth.get(api_url)
        response.raise_for_status()
        data = response.json()
    except TokenExpiredError:
        refresh_token()
        response = oauth.get(api_url)
        response.raise_for_status()
        data = response.json()

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


    # Instead of a print we return the content, cause we will need to call it again, so we just "keep" the result in to that var.
    return(html_content)


def get_info_by_sku():
  
    sku = sku_entry.get().strip()
    excel_file_path = r"C:/Users/user/Desktop/Ebay Item list.xlsx" # Custom PATH, this can be changed to fit any need.
    html_content = run_onedrive(sku) # we just get the content that we needed from the API function back there with the return to be used here.

    try:
        df = pd.read_excel(excel_file_path)
    except pd.errors.ParserError:
        result_label.config(text="Error: Invalid Excel file")
        return

    sku_row = df[df['SKU'].str.lower() == sku.lower()]

    result_text = scrolledtext.ScrolledText(window, wrap=tk.WORD, width=40, height=10)
    result_text.pack()
    
    sku_row = df[df['SKU'] == sku]
    if not sku_row.empty:
        info = {
            'Brand': sku_row.iloc[0]['Brand'],
            'Item Name': sku_row.iloc[0]['Item Name'],
            'Condition': sku_row.iloc[0]['Condition'],
            'Model Number': sku_row.iloc[0]['Model Number'],
            'Serial No.': sku_row.iloc[0]['Serial No.'],
            'Manufacture': sku_row.iloc[0]['Manufacture'],
            'Dimensions': sku_row.iloc[0]['Dimentions'],
            'Material': sku_row.iloc[0]['Material'],
            'Accessory': sku_row.iloc[0]['Accessory'],
            'Comment': sku_row.iloc[0]['Comment']
        }

        comment_lines = info['Comment'].split('\n')

        formatted_comment = '<br>'.join(comment_lines)

        info['Comment'] = formatted_comment

        html_output = f'''
        <!DOCTYPE html>
<html>
<head>
    <!-- Meta tags, title, and links to CSS or other resources go here -->
    <!-- Specify the character encoding for Japanese -->
    <meta charset="Shift_JIS">
    
    <!-- Specify UTF-8 as the default encoding for the rest of the document -->
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    
    <!-- Other meta tags, title, and links to resources go here -->
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Lux Fleek Japan</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-9ndCyUaIbzAi2FUVXJi0CjmCapSmO7SnpJef0486qhLnuZ2cdeRhO02iuK6FUUVM" crossorigin="anonymous">
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js" integrity="sha384-geWF76RCwLtnZ8qwWowPQNguL3RmwHVBC9FhGdlKrxdiJJigb/j/68SIy3Te4Bkz" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.11.8/dist/umd/popper.min.js" integrity="sha384-I7E8VVD/ismYTF4hNIPjVp/Zjvgyol6VFvRkX/vR+Vc4jQkC+hVqc2pM8ODewa9r" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.min.js" integrity="sha384-fbbOQedDUMZZ5KreZpsbe1LCZPVmfTnH7ois6mU1QK+m14rQ1l2bGBq41eYeM/fS" crossorigin="anonymous"></script>
    <script src="https://kit.fontawesome.com/af3e0fd724.js" crossorigin="anonymous"></script>

</head>

<body >

  <br>
    <img src="https://onedrive.live.com/embed?resid=544986CF56E6B0E4%2118699&authkey=%21AOwyfEJ-yFkJT3U&width=1200&height=270" class="mx-auto d-flex justify-content-center align-items-center"/>
    <br><br>

    <div class="card mx-auto text-center border-0">
    <h2>{info['Item Name']}</h2>
    </div>

    {html_content}

    <br><br>

    <card class="card w-75 text-left mx-auto border-0">

    <h2 class="text-left"><i class="fa-solid fa-magnifying-glass"></i>&middot; Product Details</h2><br>
    <table class="table table-bordered border-secondary text-left w-75 table-striped">
        <tbody>
          <tr>
            <th scope="row">Item Name</th>
            <td class="text-left">{info['Item Name']}</td>
          </tr>
          <tr>
            <th scope="row">Brand</th>
            <td>{info['Brand']}</td>
          </tr>
          <tr>
            <th scope="row">Condition</th>
            <td >{info['Condition']}</td>
          </tr>
          <tr>
            <th scope="row">Model Number</th>
            <td >{info['Model Number']}</td>
          </tr>
          <tr>
            <th scope="row">Serial No.</th>
            <td >{info['Serial No.']}</td>
          </tr>
          <tr>
            <th scope="row">Manufacture</th>
            <td >{info['Manufacture']}</td>
          </tr>
          <tr>
            <th scope="row">Dimensions</th>
            <td >{info['Dimensions']}</td>
          </tr>
          <tr>
            <th scope="row">Material</th>
            <td >{info['Material']}</td>
          </tr>
          <tr>
            <th scope="row">Accessory</th>
            <td >{info['Accessory']}</td>
          </tr>
          <tr>
            <th scope="row">Comment</th>
            <td >{info['Comment']}</td>
          </tr>
        </tbody>
      </table>

      <br>

    </card>

      
      <card class="card w-75 text-left mx-auto border-0">

      <hr>
      <br>

      <h2 class="text-left"><i class="fa-solid fa-star"></i>&middot; Item Ranking</h2><br>
      <table class="table rounded table-bordered border-secondary w-75">
        <tbody>
          <tr>
            <th class="text-center" scope="row" style="background: linear-gradient(45deg, #ffa155, #ffffff);">N</th>
            <td class="text-left"><b>New</b>, Perfect Condition or only without a tag.</td>
          </tr>
          <tr>
            <th class="text-center" scope="row" style="background: linear-gradient(45deg, #ff4747, #ffffff);">S</th>
            <td class="text-left"><b>Mint</b>. As if almost New.</td>
          </tr>
          <tr>
            <th class="text-center" scope="row" style="background: linear-gradient(45deg, #d64eff, #ffffff);">A</th>
            <td class="text-left"><b>Excellent</b>, near Mint, slightly used or new old stock.</td>
          </tr>
          <tr>
            <th class="text-center" scope="row" style="background: linear-gradient(45deg, #53ebff, #ffffff);">B</th>
            <td class="text-left"><b>Good</b>, some scratches, stains & used feeling.</td>
          </tr>
          <tr>
            <th class="text-center" scope="row" style="background: linear-gradient(45deg, #2e79be, #ffffff);">C</th>
            <td class="text-left"><b>Fair</b>, has obvious used feeling, but still acceptable.</td>
          </tr>
          <tr>
            <th class="text-center" scope="row" style="background: linear-gradient(45deg, #387e2f, #ffffff);">D</th>
            <td class="text-left"><b>Poor</b>, has damages, but still usable.</td>
          </tr>
            <th class="text-center" scope="row" style="background: linear-gradient(45deg, #9c9696, #ffffff);">E</th>
            <td class="text-left"><b>Junk</b> items.</td>
          </tr>
        </tbody>
      </table>
      <br><hr><br>

      <h2 class="text-left"><i class="fa-solid fa-circle-info"></i>&nbsp;&middot; Note for International Customers</h2><br>
      <br>
      <p class="text-left">
        &middot; Please check all the necessary policies.<br><br>

        &middot; Please check the product conditions and pictures very carefully before purchasing.<br>
                
        <p style="color: red;">Import duties, taxes, and charges are NOT INCLUDED in item prices or shipping costs. These charges are customer's responsibility. Please check with your country's customs office to determine what those additional charges will be prior to bidding or buying. Customs fees are normally charged by shipping companies or collected when you pick up the item. These fees will not be added to shipping charges.
        We don't under-value merchandise or mark the item as a gift on customs forms. Doing that is against international laws.</p>
</p>
    
<hr><br><br>

      <h2 class="text-left"><i class="fa-solid fa-truck-fast"></i>&nbsp;&middot; Shipping Policy</h2>
      <br><br>
<p class="text-left">
Delivery method: <b>FedEx</b>, <b>EMS</b>
We will ship within <b>3 business days</b> after confirmed your payment. (Excluding Saturdays, Sundays, and holidays at Japan time)
We only ship to the confirmed address provided by eBay.
Before you pay, please make sure your address in eBay matches the address you would like to be shipped to.
<br><br>
We will provide you with tracking information once item has been shipped.
<br><br>
Please be aware that delivery could be delayed depending on customs inspection, in case of natural disasters, and other reasons such as postal strikes.
</p>
      <hr><br><br>

      <h2 class="text-left"><i class="fa-solid fa-money-bill"></i>&nbsp;&middot; Payment Policy</h2><br>
      <br>
      <p class="text-left">
      Please make payment within <b>3 days</b> after the auction ends.
Orders will be cancelled if payment is not received
within 3 days after purchasing.
      </p>
      <br>
      <hr><br><br>

      <h2 class="text-left"><i class="fa-solid fa-rotate-left"></i>&nbsp;&middot; Return & Refund Policy</h2><br>
      <br>
      <p class="text-left">
      If the customer returns the item for the following reasons:<br><br>

      &middot; It does not fit<br>
      &middot; Changed mind<br>
      &middot; Ordered by mistake<br>
      &middot; Found a better price<br>
      &middot; Just didn't like it<br>

      <br>
<b>Customer will pay for return shipping fees</b>. Please return the item in the same condition as when it was delivered. After receiving the returned item, we will refund you after we do the inspection. If the item was damaged, a full refund can't be processed.

If the costumer reject the delivery of goods after shipping due to the custom taxes or any other reasons, we will deduct a one way shipping fee from the full amount when we refunded.
</p>
      
      <hr><br>

      Lux Fleek Japan - 2023, All Rights Reserved.

      <br><br>

      <hr><br><br>
      
      </card>

      <br><br><br>

    

</body>
</html>
        '''
        result_label.config(text=html_output)
    else:
        result_label.config(text=f'SKU {sku} not found in the Excel file')

    result_text.delete(1.0, tk.END)
    result_text.insert(tk.END, html_output)

# It's just the previous function without the onedrive part, so it wont load a auth and stuff, tbh i could just make an IF statement to not need to copy paste everything without the function, but wherevs.
def get_info_by_sku_nopic():

    sku = sku_entry.get().strip()
    excel_file_path = r"C:/Users/naohiro/OneDrive/Auction sales/Ebay/Ebay Item list.xlsx"
    
    try:
        df = pd.read_excel(excel_file_path)
    except pd.errors.ParserError:
        result_label.config(text="Error: Invalid Excel file")
        return
    
    sku_row = df[df['SKU'].str.lower() == sku.lower()]

    result_text = scrolledtext.ScrolledText(window, wrap=tk.WORD, width=40, height=10)
    result_text.pack()
    
    sku_row = df[df['SKU'] == sku]
    if not sku_row.empty:
        info = {
            'Brand': sku_row.iloc[0]['Brand'],
            'Item Name': sku_row.iloc[0]['Item Name'],
            'Condition': sku_row.iloc[0]['Condition'],
            'Model Number': sku_row.iloc[0]['Model Number'],
            'Serial No.': sku_row.iloc[0]['Serial No.'],
            'Manufacture': sku_row.iloc[0]['Manufacture'],
            'Dimensions': sku_row.iloc[0]['Dimentions'],
            'Material': sku_row.iloc[0]['Material'],
            'Accessory': sku_row.iloc[0]['Accessory'],
            'Comment': sku_row.iloc[0]['Comment']
        }

        comment_lines = info['Comment'].split('\n')
        formatted_comment = '<br>'.join(comment_lines)
        info['Comment'] = formatted_comment

        html_output = f'''
        <!DOCTYPE html>
<html>
<head>
    <!-- Meta tags, title, and links to CSS or other resources go here -->
    <!-- Specify the character encoding for Japanese -->
    <meta charset="Shift_JIS">
    
    <!-- Specify UTF-8 as the default encoding for the rest of the document -->
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    
    <!-- Other meta tags, title, and links to resources go here -->
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Lux Fleek Japan</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-9ndCyUaIbzAi2FUVXJi0CjmCapSmO7SnpJef0486qhLnuZ2cdeRhO02iuK6FUUVM" crossorigin="anonymous">
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js" integrity="sha384-geWF76RCwLtnZ8qwWowPQNguL3RmwHVBC9FhGdlKrxdiJJigb/j/68SIy3Te4Bkz" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.11.8/dist/umd/popper.min.js" integrity="sha384-I7E8VVD/ismYTF4hNIPjVp/Zjvgyol6VFvRkX/vR+Vc4jQkC+hVqc2pM8ODewa9r" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.min.js" integrity="sha384-fbbOQedDUMZZ5KreZpsbe1LCZPVmfTnH7ois6mU1QK+m14rQ1l2bGBq41eYeM/fS" crossorigin="anonymous"></script>
    <script src="https://kit.fontawesome.com/af3e0fd724.js" crossorigin="anonymous"></script>

</head>

<body >

  <br>
    <img src="https://onedrive.live.com/embed?resid=544986CF56E6B0E4%2118699&authkey=%21AOwyfEJ-yFkJT3U&width=1200&height=270" class="mx-auto d-flex justify-content-center align-items-center"/>
    <br><br>

    <card class="card w-75 text-left mx-auto border-0">

    <h2 class="text-left"><i class="fa-solid fa-magnifying-glass"></i>&middot; Product Details</h2><br>
    <table class="table table-bordered border-secondary text-left w-75 table-striped">
        <tbody>
          <tr>
            <th scope="row">Item Name</th>
            <td class="text-left">{info['Item Name']}</td>
          </tr>
          <tr>
            <th scope="row">Brand</th>
            <td>{info['Brand']}</td>
          </tr>
          <tr>
            <th scope="row">Condition</th>
            <td >{info['Condition']}</td>
          </tr>
          <tr>
            <th scope="row">Model Number</th>
            <td >{info['Model Number']}</td>
          </tr>
          <tr>
            <th scope="row">Serial No.</th>
            <td >{info['Serial No.']}</td>
          </tr>
          <tr>
            <th scope="row">Manufacture</th>
            <td >{info['Manufacture']}</td>
          </tr>
          <tr>
            <th scope="row">Dimensions</th>
            <td >{info['Dimensions']}</td>
          </tr>
          <tr>
            <th scope="row">Material</th>
            <td >{info['Material']}</td>
          </tr>
          <tr>
            <th scope="row">Accessory</th>
            <td >{info['Accessory']}</td>
          </tr>
          <tr>
            <th scope="row">Comment</th>
            <td >{info['Comment']}</td>
          </tr>
        </tbody>
      </table>

      <br>

    </card>

      
      <card class="card w-75 text-left mx-auto border-0">

      <hr>
      <br>

      <h2 class="text-left"><i class="fa-solid fa-star"></i>&middot; Item Ranking</h2><br>
      <table class="table rounded table-bordered border-secondary w-75">
        <tbody>
          <tr>
            <th class="text-center" scope="row" style="background: linear-gradient(45deg, #ffa155, #ffffff);">N</th>
            <td class="text-left"><b>New</b>, Perfect Condition or only without a tag.</td>
          </tr>
          <tr>
            <th class="text-center" scope="row" style="background: linear-gradient(45deg, #ff4747, #ffffff);">S</th>
            <td class="text-left"><b>Mint</b>. As if almost New.</td>
          </tr>
          <tr>
            <th class="text-center" scope="row" style="background: linear-gradient(45deg, #d64eff, #ffffff);">A</th>
            <td class="text-left"><b>Excellent</b>, near Mint, slightly used or new old stock.</td>
          </tr>
          <tr>
            <th class="text-center" scope="row" style="background: linear-gradient(45deg, #53ebff, #ffffff);">B</th>
            <td class="text-left"><b>Good</b>, some scratches, stains & used feeling.</td>
          </tr>
          <tr>
            <th class="text-center" scope="row" style="background: linear-gradient(45deg, #2e79be, #ffffff);">C</th>
            <td class="text-left"><b>Fair</b>, has obvious used feeling, but still acceptable.</td>
          </tr>
          <tr>
            <th class="text-center" scope="row" style="background: linear-gradient(45deg, #387e2f, #ffffff);">D</th>
            <td class="text-left"><b>Poor</b>, has damages, but still usable.</td>
          </tr>
            <th class="text-center" scope="row" style="background: linear-gradient(45deg, #9c9696, #ffffff);">E</th>
            <td class="text-left"><b>Junk</b> items.</td>
          </tr>
        </tbody>
      </table>
      <br><hr><br>

      <h2 class="text-left"><i class="fa-solid fa-circle-info"></i>&nbsp;&middot; Note for International Customers</h2><br>
      <br>
      <p class="text-left">
        &middot; Please check all the necessary policies.<br><br>

        &middot; Please check the product conditions and pictures very carefully before purchasing.<br>
                
        <p style="color: red;">Import duties, taxes, and charges are NOT INCLUDED in item prices or shipping costs. These charges are customer's responsibility. Please check with your country's customs office to determine what those additional charges will be prior to bidding or buying. Customs fees are normally charged by shipping companies or collected when you pick up the item. These fees will not be added to shipping charges.
        We don't under-value merchandise or mark the item as a gift on customs forms. Doing that is against international laws.</p>
</p>
    
<hr><br><br>

      <h2 class="text-left"><i class="fa-solid fa-truck-fast"></i>&nbsp;&middot; Shipping Policy</h2>
      <br><br>
<p class="text-left">
Delivery method: <b>FedEx</b>, <b>EMS</b>
We will ship within <b>3 business days</b> after confirmed your payment. (Excluding Saturdays, Sundays, and holidays at Japan time)
We only ship to the confirmed address provided by eBay.
Before you pay, please make sure your address in eBay matches the address you would like to be shipped to.
<br><br>
We will provide you with tracking information once item has been shipped.
<br><br>
Please be aware that delivery could be delayed depending on customs inspection, in case of natural disasters, and other reasons such as postal strikes.
</p>
      <hr><br><br>

      <h2 class="text-left"><i class="fa-solid fa-money-bill"></i>&nbsp;&middot; Payment Policy</h2><br>
      <br>
      <p class="text-left">
      Please make payment within <b>3 days</b> after the auction ends.
Orders will be cancelled if payment is not received
within 3 days after purchasing.
      </p>
      <br>
      <hr><br><br>

      <h2 class="text-left"><i class="fa-solid fa-rotate-left"></i>&nbsp;&middot; Return & Refund Policy</h2><br>
      <br>
      <p class="text-left">
      If the customer returns the item for the following reasons:<br><br>

      &middot; It does not fit<br>
      &middot; Changed mind<br>
      &middot; Ordered by mistake<br>
      &middot; Found a better price<br>
      &middot; Just didn't like it<br>

      <br>
<b>Customer will pay for return shipping fees</b>. Please return the item in the same condition as when it was delivered. After receiving the returned item, we will refund you after we do the inspection. If the item was damaged, a full refund can't be processed.

If the costumer reject the delivery of goods after shipping due to the custom taxes or any other reasons, we will deduct a one way shipping fee from the full amount when we refunded.
</p>
      
      <hr><br>

      Lux Fleek Japan - 2023, All Rights Reserved.

      <br><br>

      <hr><br><br>
      
      </card>

      <br><br><br>

    

</body>
</html>
        '''
        result_label.config(text=html_output)
        print('Excel worked!')
    else:
        result_label.config(text=f'SKU {sku} not found in the Excel file')

    result_text.delete(1.0, tk.END)
    result_text.insert(tk.END, html_output)

def get_label_text():
    sku = sku_entry.get()
    run_onedrive(sku)

# Create a function to save the HTML table to a file
def save_html_to_file():
    # Get the HTML table text from the result_label widget
    html_output = result_label.cget("text")
    
    # Ask the user to choose a file name and location to save the HTML file
    file_path = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML Files", "*.html")])
    
    # If the user cancels the file dialog, do nothing
    if not file_path:
        return
    
    # Save the HTML table to the selected file
    with open(file_path, "w") as html_file:
        html_file.write(html_output)
    
    # Inform the user that the file has been saved
    result_label.config(text=f'HTML file saved as "{file_path}"')

# Create the main window
window = tk.Tk()
window.title("OneDrive-Excel to HTML Tool (Naohiro)")

# Center the window on the screen
window.geometry("+%d+%d" % ((window.winfo_screenwidth() - window.winfo_reqwidth()) / 2,
                             (window.winfo_screenheight() - window.winfo_reqheight()) / 2))

# Set the window size to 450x300
window.geometry("400x300")

# Create an entry field for SKU input
sku_input_label = tk.Label(window, text="Input SKU")
sku_input_label.pack()
sku_entry = tk.Entry(window)
sku_entry.pack()


'''
USE THIS IF YOU WANT TO SELECT AN EXCEL FILE INSTEAD OF STATIC/FIXED PATH

# Create a label for file selection
file_label = tk.Label(window, text="Select Excel File:")
file_label.pack()


# Create a variable to store the selected file path
file_path_var = tk.StringVar()

# Create a button to open a file dialog for selecting the Excel file
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    file_path_var.set(file_path)


file_select_button = tk.Button(window, text="Browse", command=select_file)
file_select_button.pack(pady=5)
'''

# I put those "pady" just for the GUI to look better, everything was squished and terrible.

# No pictures procedure, this won't call the OD auth stuff.
fetch_button = tk.Button(window, text="HTML No Pictures", command=get_info_by_sku_nopic)
fetch_button.pack(pady=5)

# With pictures procedure, this WILL call the OD auth stuff.
fetch_button_pic = tk.Button(window, text="HTML With Pictures", command=get_info_by_sku)
fetch_button_pic.pack(pady=5)

# Create a button to save the HTML table to a file
save_button = tk.Button(window, text="Save HTML", command=save_html_to_file)
save_button.pack(pady=5)

# Create a label to display the result
result_label = tk.Label(window, text="")
result_label.pack()

watermark = tk.Label(window, text="Program made by gyuji")
watermark.pack()

# Start the GUI main loop
window.mainloop()



