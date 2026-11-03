import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

while True:
    directory = single_input("Input directory where the files are:")
    
    if directory[-1] != '/':
        directory = directory + '/'
        
    if os.path.exists(directory) == True:
        filenames = os.listdir(directory)
        
        if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
            break
            
    show_message("Error", "The directory doesn't exist or is missing files. Try again.")

descriptions_df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')
product_descriptions = {}
for index, row in descriptions_df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

prices_df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')
product_prices = {}
for index, row in prices_df.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

def get_order_data(invoice_num, filename, directory):
    df = pd.read_excel(directory + filename, header=None)
    name = ""
    email = ""
    date_str = ""
    company = ""
    address = ""
    city = ""
    state = ""
    zipcode = ""
    product_list = []
    
    for index, row in df.iterrows():
        if row[1] == 'Order Date':
            date_str = pd.to_datetime(row[2]).strftime('%m/%d/%Y')
        elif row[1] == 'Email':
            email = str(row[2])
        elif row[1] == 'Name':
            name = str(row[2])
        elif row[1] == 'Company Name':
            company = str(row[2])
        elif row[1] == 'Address':
            address = str(row[2])
        elif row[1] == 'City':
            city = str(row[2])
        elif row[1] == 'State':
            state = str(row[2])
        elif row[1] == 'Zipcode':
            zipcode = str(row[2])
        elif row[1] in product_descriptions:
            product_list.append([str(row[1]), float(row[2])])
    city_state_zip = f"{city}, {state} {zipcode}"
    order_dict = {
        'InvoiceNum': str(invoice_num),
        'Date': date_str,
        'CustomerName': name,
        'CustomerCompany': company,
        'CustomerAddress1': address,
        'CustomerAddress2': city_state_zip
    }
    total_amount = 0.0
    product_count = 0
    for item in product_list:
        if product_count < 3:
            unit_num = item[0]
            qty = item[1]
            desc = str(product_descriptions[unit_num])
            price = float(product_prices[unit_num])
            amt = qty * price
            total_amount = total_amount + amt
            order_dict[f'Description{product_count}'] = desc
            order_dict[f'Price{product_count}'] = f'${price:,.2f}'
            order_dict[f'Quantity{product_count}'] = str(int(qty))
            order_dict[f'Amount{product_count}'] = f'${amt:,.2f}'
            product_count = product_count + 1
            
    order_dict['Total'] = f'${total_amount:,.2f}'
    return order_dict, name, email, date_str


def create_invoice(invoice_num, order_data, directory):
    original_path = directory + 'MasterInvoice.pdf'
    new_path = directory + f'Invoice {invoice_num}.pdf'
    create_pdf(original_path, new_path, order_data)


def send_email(invoice_num, name, email, date, directory):
    gmail_user = "wilkinson.ryant@gmail.com"
    gmail_password = open(directory + 'gmail_password.txt').read().strip()

    recipients = [email]

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    body = f"""Dear {name},

Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    path_to_file = directory + f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message)
        
    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    mail_server.send_message(message)
    mail_server.quit()


filenames = os.listdir(directory)
filenames.sort()
invoice_num = 8515

for filename in filenames:
    
    if 'Order' in filename and 'xlsx' in filename:        
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)        
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1