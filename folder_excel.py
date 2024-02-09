import os
import tkinter as tk
from tkinter import filedialog
import pdfplumber
import xml.etree.ElementTree as ET
import pandas as pd
import re
from datetime import datetime


# Function to prompt the user to select a folder
def select_folder():
    root = tk.Tk()
    root.withdraw()  # we don't want a full GUI, so keep the root window from appearing
    folder_selected = filedialog.askdirectory()  # show the dialog to choose the directory
    return folder_selected

def read_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ''.join(page.extract_text() for page in pdf.pages)

    # Initialize dictionary to store extracted data
    info = {
        'First Name': '',
        'Last Name': '',
        'Address Line': '',
        'Address Line 2': '',
        'City': '',
        'State': '',
        'Prefix of MRN': '',
        'Medical Record Number': ''
    }

    # Split the text by lines
    lines = text.split('\n')

    # Extracting Name, which is on line 3
    name_line = lines[2]
    name_parts = name_line.split()

    # Name extraction logic
    info['First Name'] = name_parts[0]
    if len(name_parts) > 2 and len(name_parts[1]) == 1:
        info['Last Name'] = name_parts[2]
    else:
        info['Last Name'] = name_parts[1]

    # Address extraction logic
    address_line = lines[3].strip()

    # Regex for Address Line 2
    address_line_2_pattern = re.compile(r'\b(UNIT|unit|APT|Apt|Apt\.)\s.{1,7}')
    addr2_match = address_line_2_pattern.search(address_line)

    # Stop keywords for filtering address lines
    stop_keywords = ['Medical', 'COMPRADOR', 'Comprador', '醫', 'Health', 'Me','He',"Co"]

    if addr2_match:
        # Extract Address Line 2
        addr2_text = addr2_match.group()

        # Apply stop_keywords to Address Line 2
        for keyword in stop_keywords:
            index = addr2_text.find(keyword)
            if index != -1:
                addr2_text = addr2_text[:index].strip()
                break

        info['Address Line 2'] = addr2_text
        address_line = address_line[:addr2_match.start()].strip()

    # Apply stop_keywords to Address Line 1
    for keyword in stop_keywords:
        index = address_line.find(keyword)
        if index != -1:
            address_line = address_line[:index].strip()

    info['Address Line'] = address_line

    # City and State extraction logic
    city_state_line = lines[4]
    city_state_parts = city_state_line.split(',')
    if len(city_state_parts) == 2:
        city, state_zip = city_state_parts
        info['City'] = city.strip()
        state_zip_parts = state_zip.strip().split(' ')
        if state_zip_parts:
            info['State'] = state_zip_parts[0]

    # MRN extraction logic
    mrn_pattern = re.compile(
        r'(\b\d{2})-(\d{6,14})\b|'
        r'(?:Medical\s+Record\s+Number|Record\s+Number):\s*(\d+)'
    )
    mrn_match = mrn_pattern.search(text)
    if mrn_match:
        if mrn_match.group(1) and mrn_match.group(2):
            info['Prefix of MRN'] = mrn_match.group(1)
            info['Medical Record Number'] = mrn_match.group(2)
        elif mrn_match.group(3):
            info['Medical Record Number'] = mrn_match.group(3)
            info['Prefix of MRN'] = ''

    return info


def read_xml(xml_path, address_line_start):
    # Initialize dictionary to store extracted data
    xml_data = {
        'OrderID': '',
        'SKU': '',
        'Item Description': '',
        'DOCID': '',  # To store DOCID
        'Region': ''  # To store region code
    }

    # Parse the XML file
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # Extract OrderID from details
    details = root.find('.//details')
    order_id = details.find('orderId').text if details is not None and details.find('orderId') is not None else 'No ID'
    xml_data['OrderID'] = order_id

    # Find the recipient that contains the address line starting with 'address_line_start'
    for recipient in root.findall('.//recipient'):
        address = recipient.find('mailadr1')
        if address is not None and address.text.startswith(address_line_start[:4]):
            # Extract SKU
            sku_node = recipient.find('sku')
            if sku_node is not None:
                xml_data['SKU'] = sku_node.text

            # Extract DOCID
            docid_node = recipient.find('DOCID')  
            if docid_node is not None:
                xml_data['DOCID'] = docid_node.text

            # Extract Region Code
            region_cd_node = recipient.find('region_cd')  
            if region_cd_node is not None:
                xml_data['Region'] = region_cd_node.text

            # Assuming that both SKU and DOCID are mandatory for Item Description
            if xml_data['SKU'] and xml_data['DOCID']:
                xml_data['Item Description'] = get_item_description(xml_data['SKU'], xml_data['DOCID'])

            break  # Break the loop once the correct recipient is found

    return xml_data


# Function to determine item description based on SKU and DOCID
def get_item_description(sku, doc_id):
    if sku == 'AIA_0300':
        return f'Audio CD, {doc_id}'
    elif sku == 'AIB_0200':
        return f'Braille, {doc_id}'
    return 'Unknown Description'

   
def main():
    selected_folder_path = select_folder()
    data_list = []  # Prepare a list to hold all the data dictionaries

    # Get the current date
    current_date = datetime.now().strftime('%m/%d/%Y')

    # Function to process a single folder
    def process_folder(folder_path):
        nonlocal data_list
        pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
        xml_files = [f for f in os.listdir(folder_path) if f.endswith('.xml')]

        if not pdf_files or not xml_files:
            return

        xml_path = os.path.join(folder_path, xml_files[0])

        for pdf_file in pdf_files:
            pdf_path = os.path.join(folder_path, pdf_file)
            pdf_info = read_pdf(pdf_path)
            xml_data = read_xml(xml_path, pdf_info['Address Line'][:4])

            data_list.append({
                'Order ID': xml_data['OrderID'],
                'Invoice Number': '',
                'SKU': xml_data['SKU'],
                'Item Description': xml_data['Item Description'],
                'Vendor': 'AI',
                'Order Received': current_date,  # Set to the current date
                'IsKit': '1',
                'Qty': '1',
                'Unit Price': '',
                'Tax': '',
                'Prefix of MRN': pdf_info['Prefix of MRN'],
                'Medical Record Number': pdf_info['Medical Record Number'],
                'First Name': pdf_info['First Name'],
                'Last Name': pdf_info['Last Name'],
                'Region': xml_data['Region'],
                'Address Line': pdf_info['Address Line'],
                'Address Line 2': pdf_info['Address Line 2'],
                'City': pdf_info['City'],
                'State': pdf_info['State']
            })

    # Check if the selected folder is a parent folder or an individual subfolder
    if any(os.path.isdir(os.path.join(selected_folder_path, item)) for item in os.listdir(selected_folder_path)):
        # If it's a parent folder, iterate through each subfolder
        for folder_name in os.listdir(selected_folder_path):
            folder_path = os.path.join(selected_folder_path, folder_name)
            if os.path.isdir(folder_path):
                process_folder(folder_path)
    else:
        # If it's an individual subfolder, process it directly
        process_folder(selected_folder_path)

    # Export to Excel
    df = pd.DataFrame(data_list)
    output_path = 'extracted.xlsx'
    df.to_excel(output_path, index=False)
    print(f'Excel spreadsheet has been created at {output_path}')

if __name__ == "__main__":
    main()
