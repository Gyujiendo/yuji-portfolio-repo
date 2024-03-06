# A script made in python for getting information from an excel file and making into a HTML table, simple and effective way to automate making HTML pages for ebay for example (which was the goal)
# Made by Gyuji.

import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import scrolledtext


def get_info_by_sku():
    sku = sku_entry.get().strip()  # It gets whats in there to make stuff work, see down bellow for the GUI.
    excel_file_path = file_path_var.get() # And this is the file path to use as ref, made some mods so it goes automatically if you got a static file, it saves time cause u need less clicks, its less annoying as well.

    # excel_file_path = file_path_var.get() -> Use this for selecting an excel file
    # excel_file_path = r"C:/Users/user/Desktop/Ebay Item list.xlsx" -> Use this for a defined excel file path
    # "C:/Users/naohiro/OneDrive/Auction sales/Ebay/Ebay Item list.xlsx" -> Hiro's PC excel file path.
    
    # This uses the pandas library to 'read' the excel file in the variable, if fails returns an error.
    try:
        df = pd.read_excel(excel_file_path)
    except pd.errors.ParserError:
        result_label.config(text="Error: Invalid Excel file")
        return
    
    # This should make the sku "Zsomething" NOT be case sensitive, it didn't work and only works with Uppercase.
    sku_row = df[df['SKU'].str.lower() == sku.lower()]

    # This will 'vomit' the HTML that was all found into the GUI it's good for knowing that the fetch worked, but it's empty, it will be filled after the fetch is complete.
    result_text = scrolledtext.ScrolledText(window, wrap=tk.WORD, width=40, height=10)
    result_text.pack()


    #And we are using the SKU row as ref. for fetching the info that we need.     
    sku_row = df[df['SKU'] == sku]
    if not sku_row.empty:
        info = {
            #These are the columns, you can custom this to any kind of info to your liking.
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

        # I tried everything to make a <br> appear into the comment part or else everything will be compressed into a single string, this was the only solution.
        comment_lines = info['Comment'].split('\n')
        
        # Format each linewith bullets
        formatted_comment = '<br>'.join(comment_lines)

        # Wrapping everything into a var.
        info['Comment'] = formatted_comment
        
        # A whole HTML page is created, and we are putting the info that we fetched from excel into those {info['xyz']}, scroll down a lot, this is BIG.
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
        # A label to act as a placeholder for the "save successful" message, nothing special.
        result_label.config(text=html_output)
    else:
        result_label.config(text=f'SKU {sku} not found in the Excel file')

    #this will populate the GUI with what was found, not the prettiest way, but it works.
    result_text.delete(1.0, tk.END)
    result_text.insert(tk.END, html_output)


# Create a function to save the HTML table to a file
def save_html_to_file():
    # Get the HTML table text from the result_label widget
    html_output = result_label.cget("text")
    
    # This will open where you wanna save the HTML.
    file_path = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML Files", "*.html")])
    
    # If the user cancels the file dialog, do nothing
    if not file_path:
        return
    
    # Save the HTML table to the selected file
    with open(file_path, "w") as html_file:
        html_file.write(html_output)
    
    # Inform the user that the file has been saved
    result_label.config(text=f'HTML file saved as "{file_path}"')


#All GUI related stuff are down here

# Main window
window = tk.Tk()
window.title("SKU Lookup")

# The label for entering the SKU, it's just a label.
sku_label = tk.Label(window, text="Enter SKU:")
sku_label.pack()

# This is the input strip, but the GET function is only up there.
sku_entry = tk.Entry(window)
sku_entry.pack()

# Label for file selection
file_label = tk.Label(window, text="Select Excel File:")
file_label.pack()

# Store the file path for easier handling.
file_path_var = tk.StringVar()

# Create a button to open a file dialog for selecting the Excel file
# These can go if the path is fixed/static
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    file_path_var.set(file_path)
#These can go if the path is fixed/static 
file_select_button = tk.Button(window, text="Browse", command=select_file)
file_select_button.pack()

# Create a button to fetch information
fetch_button = tk.Button(window, text="Fetch Information", command=get_info_by_sku)
fetch_button.pack()

# Create a button to save the HTML table to a file
save_button = tk.Button(window, text="Save as HTML", command=save_html_to_file)
save_button.pack()

# Create a label to display the result
result_label = tk.Label(window, text="")
result_label.pack()

# Start the GUI main loop
window.mainloop()
