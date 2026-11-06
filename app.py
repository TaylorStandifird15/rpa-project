# Insert your python code here
# INSERT ANSWER FOR QUESTION 1 HERE
while True:
    before_directory = select_file("Select 'ProductData.xlsx' in the folder with 'MasterInvoice.pdf', 'Order.xlsx', etc")
    directory = before_directory.replace('ProductData.xlsx', '')
    try:
        count=0
        for file in os.listdir(directory):
            if file in ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx', 'Order4.xlsx']:
                count = count + 1
        if count ==6:
            break
        else:
            show_message('Error', 'Choose a working directory')
    except:
        show_message('Error', 'Invalid input; Choose a working directory')

# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}

df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')

for i, row in df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

df1 = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')
for i, row in df1.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

product_descriptions, product_prices

# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num : int, filename :str, directory:str):
    wb = openpyxl.load_workbook(directory + filename)
    selected_sheet = wb['OrderForm']
    name = selected_sheet['c14'].value
    company = selected_sheet['c15'].value
    date = selected_sheet['c9'].value
    email = selected_sheet['c8'].value
    date = date.strftime('%m/%d/%y')
    customer_name = selected_sheet['c14'].value
    address1 = selected_sheet['c16'].value
    address2 = (f"{selected_sheet['c17'].value},{selected_sheet['c18'].value} {selected_sheet['c19'].value}")
    order_data = {}
    order_data['InvoiceNum'] =  str(invoice_num)
    order_data['Date'] = date
    amounts = 0 
    order_data['CustomerAddress1'] = str(address1)
    order_data['CustomerName']  = str(name)
    order_data['CustomerCompany'] = str(company)  
    order_data['CustomerAddress2'] = str(address2)
    if selected_sheet['b24'].value is not None:
        product0 = selected_sheet['b24'].value
        if product0 in product_prices and product0 in product_descriptions:
            description0 = product_descriptions[product0]
            price0 = product_prices[product0]
            order_data['Price0']  = str(price0)
            order_data['Quantity0'] = selected_sheet['c24'].value 
            amount0 = (price0*selected_sheet['c24'].value)
            amounts = amount0
            order_data['Amount0'] = f'${amount0:,.2f}'
            order_data['Description0'] = str(description0)
    if selected_sheet['b25'].value is not None:
        product1 = selected_sheet['b25'].value
        if product1 in product_prices and product1 in product_descriptions:
            description1 = product_descriptions[product1]
            price1 = product_prices[product1]
            order_data['Price1']  = str(price1)   
            order_data['Quantity1'] = selected_sheet['c25'].value 
            amount1 = (price1*selected_sheet['c25'].value)
            amounts = amounts + amount1
            order_data['Amount1'] = f'${amount1:,.2f}'
            order_data['Description1'] = str(description1)
    if selected_sheet['b26'].value is not None:
        product2 = selected_sheet['b26'].value
        if product2 in product_prices and product2 in product_descriptions:
            description2 = product_descriptions[product2]
            price2 = product_prices[product2]
            order_data['Price2']  = str(price2)
            order_data['Quantity2'] = selected_sheet['c26'].value 
            amount2 = (price2*selected_sheet['c26'].value)
            amounts = amounts + amount2
            order_data['Amount2'] = f'${amount2:,.2f}'
            order_data['Description2'] = str(description2)
    if selected_sheet['b27'].value is not None:
        product3 = selected_sheet['b27'].value
        if product3 in product_prices and product3 in product_descriptions:
            description3 = product_descriptions[product3]
            price3 = product_prices[product3]
            order_data['Price3']  = str(price3)  
            order_data['Quantity3'] = selected_sheet['c27'].value 
            amount3 = (price3*selected_sheet['c27'].value)
            amount = amount + amount3
            order_data['Amount3'] = f'${amount3:,.2f}'
            order_data['Description3'] = str(description3)
    if selected_sheet['b28'].value is not None:
        product4 = selected_sheet['b28'].value
        if product4 in product_prices and product4 in product_descriptions:
            description4 = product_descriptions[product4]
            price4 = product_prices[product4]
            order_data['Price4']  = str(price4)   
            order_data['Quantity4'] = selected_sheet['c28'].value 
            amount4 = (price4*selected_sheet['c28'].value)
            amounts = amounts + amount4
            order_data['Amount4'] = f'${amount4:,.2f}'
            order_data['Description4'] = str(description4)
    if selected_sheet['b29'].value is not None:
        product5 = selected_sheet['b29'].value
        if product5 in product_prices and product5 in product_descriptions:
            description5 = product_descriptions[product5]
            price5 = product_prices[product5]
            order_data['Price5']  = str(price5)
            order_data['Quantity5'] = selected_sheet['c29'].value 
            amount5 = (price5*selected_sheet['c29'].value)
            amounts = amounts + amount5
            order_data['Amount5'] = f'${amount5:,.2f}'
            order_data['Description5'] = str(description5)
    order_data['Total'] =  f'${amounts:,.2f}'
   
    return (order_data, name, email, date)

order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
order_data

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num: int, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)

create_invoice(8515, order_data , directory)

# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num,name,email,date,directory):
    gmail_user = "standifird.taylor8@gmail.com"
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()


    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

# Body of the email


    body = f"""Dear {name},
    
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
    message['To'] = email
    message['Subject'] = subject
    message.set_content(body)

# Add attachment

    message = add_attachment(directory + f'Invoice {invoice_num}.pdf', message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515,name,'standifird.taylor8@gmail.com',date,directory)

# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
print(filenames)
for file in filenames:
    order = 8515
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email, date = get_order_data(order,file,directory)
        create_invoice(order,order_data,directory)
        send_email
        order = order + 1