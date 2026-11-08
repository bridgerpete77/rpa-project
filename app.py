# INSERT ANSWER FOR QUESTION 1 HERE
while True:
    temp_file = single_input('Please select the file path to the directory where your order files are located.')
    temp_file = temp_file.strip('"')
    temp_file = temp_file.strip("'")
    directory = f'{temp_file}\\'
    files = os.listdir(temp_file)
    if 'ProductData.xlsx' in files and 'MasterInvoice.pdf' in files and 'Order1.xlsx' in files and 'Order2.xlsx' in files and 'Order3.xlsx' in files and 'Order4.xlsx' in files:
        break
    else:
        show_message('Error',"The file path entered doesn't contain the necessary files in it's directory. Please try again.")

desc_df = pd.read_excel(f'{directory}ProductData.xlsx',sheet_name='Descriptions')
price_df = pd.read_excel(f'{directory}ProductData.xlsx',sheet_name='Prices')
product_descriptions = {}
for index,row in desc_df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']
product_prices = {}
for index,row in price_df.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']
print(product_descriptions)
print(product_prices)


# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num, filename, directory):
    datasheet = openpyxl.load_workbook(f'{directory}{filename}')['OrderForm']
    customername = datasheet['C14'].value
    customeremail = datasheet['C13'].value
    orderdate = datasheet['C9'].value.strftime('%m/%d/%Y')
    order_data = {}
    order_data['InvoiceNum'] = invoice_num
    order_data['Date'] = datasheet['C9'].value.strftime('%m/%d/%Y')
    order_data['CustomerName'] = datasheet['C14'].value
    order_data['CustomerAddress1'] = datasheet['C16'].value
    order_data['CustomerCompany'] = datasheet['C15'].value
    order_data['CustomerAddress2'] = f'{datasheet['C17'].value}, {datasheet['C18'].value} {datasheet['C19'].value}'

    order_data['Price0'] = str(product_prices[datasheet['B24'].value])
    order_data['Quantity0'] = str(datasheet['C24'].value)
    order_data['Amount0'] = str(float(order_data['Price0'])*float(order_data['Quantity0']))
    order_data['Total'] = str(float(order_data['Amount0']))
    if pd.isna(datasheet['B25'].value) == False:
        order_data['Price1'] = str(product_prices[datasheet['B25'].value])
        order_data['Quantity1'] = str(datasheet['C25'].value)
        order_data['Amount1'] = str(float(order_data['Price1'])*float(order_data['Quantity1']))
        order_data['Total'] = str(float(order_data['Total']) + float(order_data['Amount1']))

    if pd.isna(datasheet['B26'].value) == False:
        order_data['Price2'] = str(product_prices[datasheet['B26'].value])
        order_data['Quantity2'] = str(datasheet['C26'].value)
        order_data['Amount2'] = str(float(order_data['Price2'])*float(order_data['Quantity2']))
        order_data['Total'] = str(float(order_data['Total']) + float(order_data['Amount2']))

    if pd.isna(datasheet['B27'].value) == False:
        order_data['Price3'] = str(product_prices[datasheet['B27'].value])
        order_data['Quantity3'] = str(datasheet['C27'].value)
        order_data['Amount3'] = str(float(order_data['Price3'])*float(order_data['Quantity3']))
        order_data['Total'] = str(float(order_data['Total']) + float(order_data['Amount3']))

    if pd.isna(datasheet['B28'].value) == False:
        order_data['Price4'] = str(product_prices[datasheet['B28'].value])
        order_data['Quantity4'] = str(datasheet['C28'].value)
        order_data['Amount4'] = str(float(order_data['Price4'])*float(order_data['Quantity4']))
        order_data['Total'] = str(float(order_data['Total']) + float(order_data['Amount4']))

    order_data['Description0'] = str(product_descriptions[datasheet['B24'].value])

    if pd.isna(datasheet['B25'].value) == False:
        order_data['Description1'] = product_descriptions[datasheet['B25'].value]

    if pd.isna(datasheet['B26'].value) == False:
        order_data['Description2'] = product_descriptions[datasheet['B26'].value]

    if pd.isna(datasheet['B27'].value) == False:
        order_data['Description3'] = product_descriptions[datasheet['B27'].value]

    if pd.isna(datasheet['B28'].value) == False:
        order_data['Description4'] = product_descriptions[datasheet['B28'].value]

    return [order_data, customername, customeremail, orderdate]


# INSERT ANSWER FOR QUESTION 4 HERE
order_data = get_order_data(8515, 'Order1.xlsx', directory)[0]
print(order_data)
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory+'MasterInvoice.pdf', f"Invoice {invoice_num}.pdf", order_data)
create_invoice(8515, order_data, directory)



# INSERT ANSWER FOR QUESTION 5 Here
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()
def send_email(invoice_num, name, email, date, directory):
    
    gmail_user = email


    # List of recipients

    recipients = {email}

    # Email subect

    subject = f"Your Order from PythonTech Co. - Invoice {invoice_num}"

    # Body of the email

    body = f"""Dear {name},

    Attached is your invoice dated {date}.

    Thank you for your business.

    Python Tech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['CC'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file1 = f'{directory}Invoice {invoice_num}.pdf'

    message = add_attachment(path_to_file1, message)
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
inv_num = 8515
for file in filenames:
    if 'Order' in file and '.xlsx' in file:
        get_order_data(inv_num, file, directory)
        data = get_order_data(inv_num, file, directory)
        create_invoice(inv_num, data[0], directory)
        send_email(inv_num, data[1], data[2], data[3], directory)
        inv_num = inv_num + 1
        



