import os
import zipfile
import shutil
import datetime
from pathlib import Path
import pdfplumber
import xml.etree.ElementTree as ET
import pandas as pd
import re
from xml.dom import minidom
import xml.etree.ElementTree as ET
import csv



def copy_zips_to_complete(main_directory):
    # Ensure the 'Complete' folder exists
    complete_folder = os.path.join(main_directory, 'Complete')
    if not os.path.exists(complete_folder):
        os.makedirs(complete_folder)

    # Iterate over all files in the main directory
    for file in os.listdir(main_directory):
        if file.endswith('.zip'):
            # Construct full file path
            file_path = os.path.join(main_directory, file)

            # Copy the zip file to the 'Complete' folder
            shutil.copy(file_path, complete_folder)


def extract_and_combine_zips(main_directory):
    # Create a dictionary to track zip files with the same 8-digit prefix
    zip_groups = {}

    # Iterate over all files in the main directory
    for file in os.listdir(main_directory):
        if file.endswith('.zip'):
            # Extract the 8-digit prefix
            match = re.match(r'(\d{8})_.*\.zip$', file)
            if match:
                prefix = match.group(1)
                # Group files by their 8-digit prefix
                if prefix in zip_groups:
                    zip_groups[prefix].append(file)
                else:
                    zip_groups[prefix] = [file]

    # Process each group of files
    for prefix, files in zip_groups.items():
        # Create a directory for the group
        group_dir = os.path.join(main_directory, prefix)
        if not os.path.exists(group_dir):
            os.makedirs(group_dir)

        # Extract files from each zip in the group
        for zip_file in files:
            with zipfile.ZipFile(os.path.join(main_directory, zip_file), 'r') as zip_ref:
                zip_ref.extractall(group_dir)


def calculate_ship_date(received_date):
    # Calculate ship date, 5 business days from received date excluding weekends
    ship_date = received_date
    for _ in range(5):
        ship_date += datetime.timedelta(days=1)
        while ship_date.weekday() >= 5:  # Mon-Fri are 0-4
            ship_date += datetime.timedelta(days=1)
    return ship_date

def create_fulfillment_xml(main_directory):
    # Ensure the XML folder exists
    xml_folder = os.path.join(main_directory, 'XML')
    if not os.path.exists(xml_folder):
        os.makedirs(xml_folder)

    # Iterate over all directories in the main directory
    for folder_name in os.listdir(main_directory):
        folder_path = os.path.join(main_directory, folder_name)
        if os.path.isdir(folder_path) and folder_name.isdigit():
            # For each folder, create an XML file
            xml_file_path = os.path.join(xml_folder, f'order_{folder_name}.xml')
            root = ET.Element('fulfillment')
            
            # Add elements to the XML
            ET.SubElement(root, 'VendorID').text = 'N'
            ET.SubElement(root, 'OrderID').text = folder_name
            ET.SubElement(root, 'Status').text = 'Received'
            received_date = datetime.datetime.now().date()
            ET.SubElement(root, 'ReceivedDate').text = str(received_date)
            ship_date = calculate_ship_date(received_date)
            ET.SubElement(root, 'ShipDate').text = str(ship_date)
            ET.SubElement(root, 'ShippingMethod').text = 'USPS'
            ET.SubElement(root, 'ShippingCost').text = '10.00'
            ET.SubElement(root, 'Comments')
            ET.SubElement(root, 'PackagesCount')
            ET.SubElement(root, 'Tracking')

            # Write the XML file
            tree = ET.ElementTree(root)
            tree.write(xml_file_path)


def update_daily_status(main_directory):
    xml_folder = os.path.join(main_directory, 'XML')
    tsv_file_path = os.path.join(main_directory, 'daily_status.tsv')

    # Read existing data from TSV
    existing_data = []
    with open(tsv_file_path, 'r', newline='') as tsvfile:
        reader = csv.reader(tsvfile, delimiter='\t')
        existing_data.extend(list(reader))

    # Iterate over XML files and append new data
    new_data = []
    for xml_file in os.listdir(xml_folder):
        if xml_file.endswith('.xml'):
            xml_file_path = os.path.join(xml_folder, xml_file)
            tree = ET.parse(xml_file_path)
            root = tree.getroot()

            # Extract required fields
            vendor_id = root.find('VendorID').text
            order_id = root.find('OrderID').text
            status = root.find('Status').text
            received_date = root.find('ReceivedDate').text
            ship_date = root.find('ShipDate').text
            shipping_method = root.find('ShippingMethod').text
            shipping_cost = root.find('ShippingCost').text
            comments = root.find('Comments').text
            packages_count = root.find('PackagesCount').text
            tracking = root.find('Tracking').text

            # Create a row for the TSV
            row = [vendor_id, order_id, status, received_date, ship_date, shipping_method, shipping_cost, comments, packages_count, tracking]
            if row not in existing_data:
                new_data.append(row)

    # Append new data to the TSV
    with open(tsv_file_path, 'a', newline='') as tsvfile:
        writer = csv.writer(tsvfile, delimiter='\t')
        writer.writerows(new_data)

def separate_pdf_files(main_directory):
    for folder_name in os.listdir(main_directory):
        folder_path = os.path.join(main_directory, folder_name)
        if os.path.isdir(folder_path) and folder_name.isdigit():
            pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
            original_xml_files = [f for f in os.listdir(folder_path) if f.endswith('.xml')]

            if not original_xml_files:
                continue  # Skip if there are no XML files

            for index, pdf_file in enumerate(pdf_files, start=1):
                if index == 1:
                    # Keep the first PDF in the original folder
                    continue

                # Create a new folder for each subsequent PDF
                new_folder_name = f"{folder_name}.{index}"
                new_folder_path = os.path.join(main_directory, new_folder_name)
                if not os.path.exists(new_folder_path):
                    os.makedirs(new_folder_path)

                # Move the PDF file to the new folder
                src_pdf_path = os.path.join(folder_path, pdf_file)
                dst_pdf_path = os.path.join(new_folder_path, pdf_file)
                shutil.move(src_pdf_path, dst_pdf_path)

                # Copy the original XML file(s) to the new folder
                for xml_file in original_xml_files:
                    src_xml_path = os.path.join(folder_path, xml_file)
                    shutil.copy(src_xml_path, new_folder_path)


def move_folders_to_date_directory(main_directory):
    # Get current date
    current_date = datetime.datetime.now()
    date_folder_name = f"{current_date.month}.{current_date.day}"
    date_folder_path = os.path.join(main_directory, date_folder_name)

    # Create the date directory if it doesn't exist
    if not os.path.exists(date_folder_path):
        os.makedirs(date_folder_path)

    # Move the [8 digits] and suffixed folders to the date directory
    for folder_name in os.listdir(main_directory):
        folder_path = os.path.join(main_directory, folder_name)
        if os.path.isdir(folder_path) and (folder_name.isdigit() or re.match(r'\d{8}\.\d+', folder_name)):
            shutil.move(folder_path, date_folder_path)



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


def process_data_in_date_folder(date_folder_path):
    data_list = []  # List to hold all the data dictionaries
    current_date = datetime.datetime.now().strftime('%m/%d/%Y')

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

    # Iterate through each subfolder in the date folder
    for folder_name in os.listdir(date_folder_path):
        folder_path = os.path.join(date_folder_path, folder_name)
        if os.path.isdir(folder_path):
            process_folder(folder_path)

    # Export to Excel
    df = pd.DataFrame(data_list)
    output_path = os.path.join(date_folder_path, 'extracted.xlsx')
    df.to_excel(output_path, index=False)
    print(f'Excel spreadsheet has been created at {output_path}')


def main():
    # Set the main directory to the current script's directory
    main_directory = os.path.dirname(os.path.realpath(__file__))

    # Step 1: Copy zips to Complete folder
    copy_zips_to_complete(main_directory)

    # Step 2: Extract and combine zips
    extract_and_combine_zips(main_directory)

    # Step 3: Create fulfillment XML
    create_fulfillment_xml(main_directory)

    # Step 4: Update daily_status.tsv
    update_daily_status(main_directory)

    # Step 5: Separate PDF files
    separate_pdf_files(main_directory)

    # Step 6: Move folders to date directory
    move_folders_to_date_directory(main_directory)

     # Step 7: Process data in date folder
    date_folder_name = f"{datetime.datetime.now().month}.{datetime.datetime.now().day}"
    date_folder_path = os.path.join(main_directory, date_folder_name)
    process_data_in_date_folder(date_folder_path)

    # ... rest of the main function ...

if __name__ == "__main__":
    main()