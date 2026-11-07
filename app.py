# Insert your python code here
# INSERT ANSWER FOR QUESTION 1 HERE
while True:  
    directory = single_input("Please input the directory.")
    directory = directory.strip('"')
    if directory[-1] != '\\':
        directory = directory + '\\'
    try:
        files = os.listdir(directory)
        if "ProductData.xlsx" in files and "MasterInvoice.pdf" in files and "Order1.xlsx" in files and "Order2.xlsx" in files and "Order3.xlsx" in files and "Order4.xlsx" in files:
            show_message("Success!", "Good job! The directory you entered includes the required files!")
            break
        else:
            show_message("Error", "The directory you entered does include the required files. Please try again.")
    except:
        show_message("Error", "The directory you entered is incorrect. Please try again.")

# INSERT ANSWER FOR QUESTION 2 HERE
file_path = directory + 'ProductData.xlsx'

product_descriptions = {}
description = pd.read_excel(file_path,'Descriptions')
for index,row in description.iterrows():
    item_desc = row['Description']
    unit_num = row['Unit number']
    product_descriptions[unit_num] = item_desc

product_prices = {}
price = pd.read_excel(file_path,'Prices')
for index,row in price.iterrows():
    unit_price = row['Price per unit']
    unit_num = row['Unit number']
    product_prices[unit_num] = unit_price

# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num, filename, directory):
    order_form = openpyxl.load_workbook(directory + filename, data_only=True)
    selected_sheet = order_form['OrderForm']

    order_data = {} 

    form = get_pdf(directory + 'MasterInvoice.pdf')

    while True: 
        total = 0
        if pd.isnull(selected_sheet['B24'].value) == False:
            unit0 = selected_sheet['B24'].value
            quantity0 = selected_sheet['C24'].value
            price0 = product_prices[unit0]
            price0str = f'${price0:,.2f}'
            amount0 = price0 * quantity0
            amount0str = f'${amount0:,.2f}'
            order_data['Price0'] = str(price0str)
            order_data['Quantity0'] = str(quantity0)
            order_data['Amount0'] = str(amount0str)
            order_data['Description0'] = str(product_descriptions[unit0])
            total = total + amount0
        else:
            break

        if pd.isnull(selected_sheet['B25'].value) == False:    
            unit1 = selected_sheet['B25'].value
            quantity1 = selected_sheet['C25'].value 
            price1 = product_prices[unit1]
            price1str = f'${price1:,.2f}'
            amount1 = price1 * quantity1
            amount1str = f'${amount1:,.2f}'
            order_data['Price1'] = str(price1str)
            order_data['Quantity1'] = str(quantity1)
            order_data['Amount1'] = str(amount1str)
            order_data['Description1'] = str(product_descriptions[unit1])
            total = total + amount1
        else:
            break 
            
        if pd.isnull(selected_sheet['B26'].value) == False:
            unit2 = selected_sheet['B26'].value
            quantity2 = selected_sheet['C26'].value
            price2 = product_prices[unit2]
            price2str = f'${price2:,.2f}'
            amount2 = price2 * quantity2
            amount2str = f'${amount2:,.2f}'
            order_data['Price2'] = str(price2str)
            order_data['Quantity2'] = str(quantity2)
            order_data['Amount2'] = str(amount2str)
            order_data['Description2'] = str(product_descriptions[unit2])   
            total = total + amount2
        else:
            break

        if pd.isnull(selected_sheet['B27'].value) == False:
            unit3 = selected_sheet['B27'].value
            quantity3 = selected_sheet['C27'].value
            price3 = product_prices[unit3]
            price3str = f'${price3:,.2f}'
            amount3 = price3 * quantity3
            amount3str = f'${amount3:,.2f}'
            order_data['Price3'] = str(price3str)
            order_data['Quantity3'] = str(quantity3)
            order_data['Amount3'] = str(amount3str)
            order_data['Description3'] = str(product_descriptions[unit3])   
            total = total + amount2
        else:
            break
    
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value
    date = selected_sheet['C9'].value
    datestr = date.strftime('%m/%d/%Y')

    street = selected_sheet['C16'].value
    city = selected_sheet['C17'].value
    state = selected_sheet['C18'].value
    zipcode = selected_sheet['C19'].value
    customer_address1 = str(street)
    customer_address2 = str(city) + ', ' + str(state) + ' ' + str(zipcode)
    customer_company = selected_sheet['C15'].value

    totalstr = f'${total:,.2f}'

    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = datestr
    order_data['Total'] = str(totalstr)
    order_data['CustomerAddress1'] = str(customer_address1)
    order_data['CustomerAddress2'] = str(customer_address2)
    order_data['CustomerName'] = str(name)
    order_data['CustomerCompany'] = str(customer_company)

    return order_data, name, email, datestr

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', f'Invoice {invoice_num}.pdf', order_data)

# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num, name, email, date, directory):
    # Identify user and password
    gmail_user = "ekei0480@gmail.com"
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()

    # List of recipients
    recipients = email

    # Email subect
    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email
    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.
    
    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    # Initialize the email message
    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment
    path_to_file = directory+ f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message)
            
    # Log in to the Gmail server
    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email
    mail_server.send_message(message)
    mail_server.quit()

# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)

invoice_num = 8515

for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data,name,email,date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num,order_data,directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1