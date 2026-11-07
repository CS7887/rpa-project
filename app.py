# Insert your python code here

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

# INSERT ANSWER FOR QUESTION 1 HERE
while True:    
    directory = single_input("Please input the directory where the file is located.")
    if directory[-1] == "/":
        None
    else:
        directory = directory + '/'
        
    if os.path.exists(directory) == True:
        files = os.listdir(directory)
    
        if "ProductData.xlsx" in files and "MasterInvoice.pdf" in files and "Order1.xlsx" in files and "Order2.xlsx" in files and "Order3.xlsx" in files and "Order4.xlsx" in files:
                break
        else:
            show_message("Error", "Please re-enter the directory with the appropriate files.")
    else:
        show_message("Error","This path does not exist.")


# INSERT ANSWER FOR QUESTION 2 HERE

product_descriptions = {}
product_prices = {}

descriptions_df = pd.read_excel(directory + "ProductData.xlsx" , sheet_name='Descriptions')
for i, row in descriptions_df.iterrows():
    product_descriptions[row["Unit number"]] = row.Description

price_df = pd.read_excel(directory + "ProductData.xlsx" , sheet_name='Prices')
for i, row in price_df.iterrows():
    product_prices[row["Unit number"]] = row["Price per unit"]

# INSERT ANSWER FOR QUESTION 3 HERE

def get_order_data(invoice_num, filename, directory):
    order_data = {}
    template = openpyxl.load_workbook(directory + filename)
    selected_sheet = template['OrderForm']
    invoice_num = str(invoice_num)
    CustomerName = selected_sheet['C14'].value
    CustomerEmail = selected_sheet['C13'].value
    OrderDate = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    order_data['InvoiceNum'] = invoice_num
    order_data['Date'] = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    order_data['CustomerAddress1'] = selected_sheet['C16'].value
    order_data['CustomerAddress2'] = f'{selected_sheet['C17'].value}, {selected_sheet['C18'].value} {selected_sheet['C19'].value}'
    order_data['CustomerName'] = selected_sheet['C14'].value
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    increment =24
    overall_total = 0
    initialize = 0
    while True:
        Unit_Number = selected_sheet[f'B{increment}'].value
        Quantity = selected_sheet[f'C{increment}'].value
        if Quantity is None:
            break
        else: 
            increment = increment + 1
            price = product_prices[Unit_Number]
            individual_total = price * Quantity
            overall_total = individual_total + overall_total
            product_name = product_descriptions [Unit_Number]
            order_data[f'Description{initialize}'] = product_name
            order_data[f'Price{initialize}'] = f'${price:,.2f}'
            order_data[f'Quantity{initialize}'] = Quantity
            order_data[f'Amount{initialize}'] = f'${individual_total:,.2f}'
            initialize = initialize +1 
    overall_total = f'${overall_total:,.2f}'
    order_data['Total'] = overall_total

    
    return  order_data, CustomerName,CustomerEmail,OrderDate
order_data, name, email,date = get_order_data(8515, "Order1.xlsx",directory)


# INSERT ANSWER FOR QUESTION 4 HERE

def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', f'Invoice {invoice_num}.pdf', order_data)



create_invoice(8515,order_data, directory)

# INSERT ANSWER FOR QUESTION 5 HERE.

gmail_password_path = select_file('Please select the gmail password file')

gmail_password = open(gmail_password_path).read()
def send_email(invoice_num, name, email, date, directory):
    invoice_num = str(invoice_num)
# Enter your gmail user name and password

    gmail_user = "sosa070101@gmail.com"
    

    # List of recipients

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""
    Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.

    PythonTech Co. 
    10 Python Drive
    Pleasant Grove, UT 84602
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file1 = directory + f'Invoice {invoice_num}.pdf'
    

    messsage = add_attachment(path_to_file1, message)
    
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()   

send_email(8515,'Thomas Holt','sosa7887@byu.edu','11/13/2023',directory) 

# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
filenames.sort()
filenames
starting_invoice = 8515
for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email,date = get_order_data(starting_invoice,file,directory)
        create_invoice(starting_invoice,order_data, directory)
        send_email(starting_invoice,name, email,date,directory)
        starting_invoice = starting_invoice + 1
    else:
        None