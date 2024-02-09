import os
import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog
from datetime import datetime, timedelta

def select_directory():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    directory = filedialog.askdirectory(title="Select a Directory")
    return directory

def add_business_days(start_date, days_to_add):
    current_date = start_date
    while days_to_add > 0:
        current_date += timedelta(days=1)
        if current_date.weekday() < 5:  # Check if it's a weekday (0=Monday, 4=Friday)
            days_to_add -= 1
    return current_date

def create_daily_folder(parent_directory):
    today = datetime.now()
    daily_folder_name = today.strftime('%m.%d')
    daily_folder_path = os.path.join(parent_directory, daily_folder_name)
    os.makedirs(daily_folder_path, exist_ok=True)
    return daily_folder_path

def move_files_to_daily_folder(daily_folder_path, source_directory):
    # Move the 'XML' folder
    xml_folder_source = os.path.join(source_directory, 'XML')
    xml_folder_dest = os.path.join(daily_folder_path, 'XML')
    os.rename(xml_folder_source, xml_folder_dest)

    # Move the daily status report
    tsv_file_source = os.path.join(source_directory, 'daily_status.tsv')
    tsv_file_dest = os.path.join(daily_folder_path, 'daily_status.tsv')
    os.rename(tsv_file_source, tsv_file_dest)

def combine_and_move_zips(directory):
    zip_files = [filename for filename in os.listdir(directory) if filename.endswith('.zip')]
    combined_folders = {}

    for zip_filename in zip_files:
        # Extract the first 8 digits as the identifier
        identifier = zip_filename[:8]

        # If a folder for this identifier doesn't exist, create one
        if identifier not in combined_folders:
            combined_folders[identifier] = []

        # Add the zip file to the list for this identifier
        combined_folders[identifier].append(zip_filename)

    for identifier, zip_list in combined_folders.items():
        if len(zip_list) > 1:
            # Create a folder for the combined zips
            combined_folder_path = os.path.join(directory, identifier)
            os.makedirs(combined_folder_path, exist_ok=True)

            # Extract and move the zips to the combined folder
            for zip_filename in zip_list:
                zip_path = os.path.join(directory, zip_filename)
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(combined_folder_path)

                # Delete the extracted zip
                os.remove(zip_path)

    # Move the combined folders to the daily folder
    daily_folder_path = create_daily_folder(directory)
    for identifier in combined_folders.keys():
        combined_folder_path = os.path.join(directory, identifier)
        os.rename(combined_folder_path, os.path.join(daily_folder_path, identifier))

def process_zip_files(directory):
    cumulative_tsv = []
    
    for filename in os.listdir(directory):
        if filename.endswith('.zip') and 'done' in filename:
            zip_path = os.path.join(directory, filename)
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(directory)

                for file in zip_ref.namelist():
                    if file.endswith('.xml'):
                        xml_data = process_xml_file(directory, file)
                        cumulative_tsv.append(xml_data)
                    
                    # Delete the unzipped XML and DTL files
                    file_path = os.path.join(directory, file)
                    os.remove(file_path)
    
    # Write cumulative data to a .tsv file in the same directory as the zips
    tsv_file_path = os.path.join(directory, 'daily_status.tsv')
    with open(tsv_file_path, 'w') as tsv_file:
        for row in cumulative_tsv:
            tsv_file.write('\t'.join(row) + '\n')

    # Combine and move the zips
    combine_and_move_zips(directory)

    # Create a daily folder and move necessary files
    daily_folder_path = create_daily_folder(directory)
    move_files_to_daily_folder(daily_folder_path, directory)

def process_xml_file(directory, xml_file):
    xml_path = os.path.join(directory, xml_file)
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # Extract vendorIndicator
    vendor_indicator = root.find('.//vendorIndicator').text if root.find('.//vendorIndicator') is not None else 'Unknown'

    # Create a new XML element with the additional fields
    new_xml = ET.Element('fulfillment')
    ET.SubElement(new_xml, 'VendorID').text = vendor_indicator

    order_id = root.find('.//orderId').text if root.find('.//orderId') is not None else 'Unknown'
    ET.SubElement(new_xml, 'OrderID').text = order_id
    ET.SubElement(new_xml, 'Status').text = 'Received'
    ET.SubElement(new_xml, 'ReceivedDate').text = datetime.now().strftime('%Y-%m-%d')
    
    # Calculate ShipDate as 5 business days from ReceivedDate
    received_date = datetime.now()
    business_days = 5
    ship_date = add_business_days(received_date, business_days)
    ET.SubElement(new_xml, 'ShipDate').text = ship_date.strftime('%Y-%m-%d')
    
    ET.SubElement(new_xml, 'ShippingMethod').text = 'USPS'
    ET.SubElement(new_xml, 'ShippingCost').text = ''
    ET.SubElement(new_xml, 'Comments')
    ET.SubElement(new_xml, 'PackagesCount')
    ET.SubElement(new_xml, 'Tracking')

    # Collect the data and return it as a list of values
    xml_data = [
        vendor_indicator,
        order_id,
        'Received',
        datetime.now().strftime('%Y-%m-%d'),
        ship_date.strftime('%Y-%m-%d'),
        'USPS',
        '',
        '',  # Comments (empty for now)
        '',  # PackagesCount (empty for now)
        ''   # Tracking (empty for now)
    ]
    
    # Create the 'XML' folder if it doesn't exist
    xml_folder_path = os.path.join(directory, 'XML')
    os.makedirs(xml_folder_path, exist_ok=True)
    
    # Save the new XML in the 'XML' folder
    new_xml_path = os.path.join(xml_folder_path, f'order_{order_id}.xml')
    new_tree = ET.ElementTree(new_xml)
    new_tree.write(new_xml_path)
    
    return xml_data

if __name__ == "__main__":
    selected_directory = select_directory()
    if selected_directory:
        process_zip_files(selected_directory)
