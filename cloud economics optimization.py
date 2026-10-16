import pandas as pd
import numpy as np 
import chardet as ch
import logging
import csv
from pathlib import Path
from openpyxl import load_workbook
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, numbers, Font, Border, Side
import time
import datetime
import os
import tkinter as tk
from tkinter import filedialog
import json
import tempfile

# Configure logging
logging.basicConfig(format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', filename='logging_cloud_eco.log', filemode='w', level=logging.INFO)

def total_cost(dir):
    
    metering_data = data_source_normalization(dir)

    total_cost_df = metering_data.groupby('Meter Category').agg({
        'PreTaxCost':'sum'
    })
    print(total_cost_df)

    return total_cost_df

def calculate_min_max(row):
    values = [row['Saving PAYG'], row['Saving 1Y Reserv'], row['Saving 3Y Reserv']]
    non_zero_values = [value for value in values if value > 0.1]
    if non_zero_values:
        row['Min'] = min(non_zero_values)
        row['Max'] = max(non_zero_values)
    else:
        row['Min'] = 0
        row['Max'] = 0
    return row

def ensure_columns_exist(df, columns):
    for column in columns:
        if column not in df.columns:
            df[column] = 0

def adjust_column_widths(ws):
    column_widths = {}
    for row in ws.iter_rows(min_row=1, min_col=1):  # Adjust the range as necessary
        for cell in row:
            cell_value_length = len(str(cell.value)) if cell.value else 0
            column_letter = cell.column_letter
            if column_letter not in column_widths:
                column_widths[column_letter] = cell_value_length
            else:
                column_widths[column_letter] = max(column_widths[column_letter], cell_value_length)
    for column_letter, max_length in column_widths.items():
        adjusted_width = max_length + 2
        ws.column_dimensions[column_letter].width = adjusted_width

def format_headers(ws):
    header_font = Font(bold=False, size=12)
    header_alignment = Alignment(horizontal='center', vertical='center')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for cell in ws[1]:  # Assume headers are in the first row
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = thin_border

def remove_gridlines(ws):
    ws.sheet_view.showGridLines = True

def apply_number_format(ws, format_string):
    for row in ws.iter_rows(min_row=2, min_col=1):
        for cell in row:
            cell.number_format = format_string

def load_config(config_path):
    logging.info(f"Loading configuration from {config_path}")
    try:
        with open(config_path, 'r') as file:
            config = json.load(file)
        logging.info("Configuration loaded successfully")
        return config
    except Exception as e:
        logging.error(f"Failed to load configuration: {e}")
        raise

def convert_month(df, column='UsageDateTime'):
    logging.info(f"Converting column {column} to datetime and extracting month name")
    try:
        df[column] = pd.to_datetime(df[column])
        df[column] = df[column].dt.month_name(locale='Portuguese')
        logging.info("Month name extracted successfully")
        return df
    except Exception as e:
        logging.error(f"Error in converting month: {e}")
        raise

def detect_encoding(file_path):
    logging.info(f"Detecting encoding for file {file_path}")
    try:
        with open(file_path, 'rb') as f:
            result = ch.detect(f.read())
        logging.info(f"Detected encoding: {result['encoding']}")
        return result['encoding']
    except Exception as e:
        logging.error(f"Failed to detect encoding: {e}")
        raise

def extract_vcores(info):
    if info is None or not isinstance(info, str) or info.strip() == "":
        logging.warning("Empty or invalid JSON string passed to extract_vcores.")
        return 0
    
    try:
        info_dict = json.loads(info)
        vCores = info_dict.get('vCores', 0)
        return vCores
    except json.JSONDecodeError as e:
        logging.error(f"Error decoding JSON: {e}")
        return 0

def normalize_csv(df):
    logging.info("Normalizing CSV DataFrame")
    try:
        if 'UsageQuantity' in df.columns:
            logging.info("DataFrame already normalized")
            return df
        
        if 'Cost' in df.columns:
            logging.info("Normalizing based on Type 2")
            return df.rename(columns={
                'BillingCurrency': 'Currency', 
                'Cost': 'PreTaxCost',
                'BillingPeriodEndDate': 'UsageDateTime', 
                'MeterSubCategory': 'MeterSubcategory', 
                'Product': 'ProductName', 
                'UnitPrice': 'ResourceRate',
                'Quantity': 'UsageQuantity', 
                'ResourceId': 'InstanceId'
            })

        if 'BillingCurrencyCode' in df.columns:
            logging.info("Normalizing based on Type 3")
            return df.rename(columns={
                'BillingCurrencyCode': 'Currency', 
                'CostInBillingCurrency': 'PreTaxCost',
                'BillingPeriodEndDate': 'UsageDateTime', 
                'MeterSubCategory': 'MeterSubcategory', 
                'UnitPrice': 'ResourceRate',
                'Quantity': 'UsageQuantity', 
                'ResourceId': 'InstanceId'
            })

        if 'ProductId' in df.columns and 'exchangeRatePricingToBilling' in df.columns:
            logging.info("Normalizing based on Type 4")
            df['ResourceRate'] = df['costInBillingCurrency'] / df['quantity']
            return df.rename(columns={
                'additionalInfo': 'AdditionalInfo',
                'billingCurrency': 'Currency', 
                'billingPeriodEndDate': 'UsageDateTime',
                'costInBillingCurrency': 'PreTaxCost', 
                'meterCategory': 'MeterCategory', 
                'meterSubCategory': 'MeterSubcategory', 
                'quantity': 'UsageQuantity', 
                'ResourceId': 'InstanceId', 
                'resourceLocation': 'ResourceLocation', 
                'subscriptionName': 'SubscriptionName',
                'meterName': 'MeterName', 
            })
        
        if 'customerTenantId' in df.columns: 
            logging.info("Normalizing based on Type 5")
            return df.rename(columns={'productOrderName':'ProductName',
                                      'additionalInfo':'AdditionalInfo',
                                      'billingPeriodEndDate':'UsageDateTime',
                                      'meterCategory':'MeterCategory',
                                      'meterName':'MeterName',
                                      'meterSubCategory':'MeterSubcategory',
                                      'quantity':'UsageQuantity',
                                      'resourceId':'InstanceId',
                                      'resourceLocation':'ResourceLocation',
                                      'unitPrice':'ResourceRate',
                                      'billingCurrency':'Currency',
                                      'costInBillingCurrency':'PreTaxCost',
                                      'subscriptionName':'SubscriptionName'})

        logging.error("No matching columns found for normalization.")
        return None
    except Exception as e:
        logging.error(f"Normalization failed: {e}")
        raise

def normalize_xlsx(df):
    logging.info("Normalizing XLSX DataFrame")
    try:
        df = df.rename(columns={
            'Amount': 'PreTaxCost', 
            'Category': 'MeterCategory', 
            'Location': 'ResourceLocation', 
            'Meter Name': 'MeterName', 
            'Product Name': 'ProductName', 
            'Publisher Subscription Name': 'SubscriptionName', 
            'Quantity': 'UsageQuantity', 
            'Subcategory': 'MeterSubcategory', 
            'Unit Price': 'ResourceRate', 
            'Usage End Date': 'UsageDateTime'
        })
        logging.info("Normalization successful")
        return df
    except Exception as e:
        logging.error(f"Normalization failed: {e}")
        raise

def data_source_normalization(dir: Path):
    if not dir:
        logging.error("No file selected.")
        return None

    try:
        logging.info(f"Processing file: {dir}")
        config = load_config("config.json")
        _, file_extension = os.path.splitext(dir)
        logging.info(f"File extension detected: {file_extension}")

        if file_extension == '.csv':
            encoding = 'ISO-8859-1'
            df = pd.read_csv(dir, encoding=encoding)
            normalized_df = normalize_csv(df)
            if normalized_df is None:
                logging.error("Normalization failed for CSV file.")
                return None
            logging.info("CSV file normalized successfully")
            return normalized_df

        elif file_extension == '.xlsx':
            df = pd.read_excel(dir, sheet_name='Metering Data')
            normalized_df = normalize_xlsx(df)
            if normalized_df is None:
                logging.error("Normalization failed for XLSX file.")
                return None
            logging.info("XLSX file normalized successfully")
            return normalized_df

        else:
            logging.error('Unsupported file type. Please select a .xlsx or .csv file.')
            return None

    except Exception as e:
        logging.error(f"Error processing file: {e}")
        return None

def virtual_machines(dir: Path, config):

    """
    Processes Azure Export or CloudIQ data source for virtual machines.
    Parameters:
    dir (Path): The directory path of the CSV file to be processed.
    config (dict): Configuration parameters including month_hours, threshold_virtual_machines_1y_reserv,
                   threshold_virtual_machines_3y_reserv, reservation_hours, virtual_machines_1y_reserv,
                   virtual_machines_3y_reserv.
    Returns:
    DataFrame: The processed DataFrame with calculated fields for virtual machines.
    """
    
    logging.info("Starting virtual machines processing")

    try:
        # Load the CSV file into a DataFrame
        metering_data = data_source_normalization(dir)
        
        if metering_data is None:
            logging.error("Failed to load and normalize data source")
            return None
        
        # Filter out rows where 'MeterCategory' is 'Virtual Machines'
        logging.info("Filtering data for 'Virtual Machines' category")
        virtual_machines = metering_data[metering_data['MeterCategory'] == 'Virtual Machines'].copy()
        
        if virtual_machines.empty:
            logging.warning("No data found for 'Virtual Machines'")
            # Return an empty DataFrame with the necessary columns
            columns = [
                'MeterName', 'UsageDateTime', 'SubscriptionName', 'ProductName',
                'MeterCategory', 'MeterSubcategory', 'UsageQuantity', 'ResourceRate', 'PreTaxCost',
                'Currency', 'ResourceLocation', 'Qty Instances', 'Reserves 1Y', 'Reserves 3Y',
                'Reserv Hours 1Y', 'Reserv Hours 3Y', 'PAYG Hours 1Y', 'PAYG Hours 3Y',
                'UP 1Year Reserv', 'UP 3Years Reserv', 'Amount 1Year Reserv', 'Amount 3Years Reserv',
                'Saving 1Y Reserv', 'Saving 3Y Reserv', 'Final Saving'
            ]
            return pd.DataFrame(columns=columns)

        # Convert 'UsageQuantity' column to float data type
        logging.info("Converting 'UsageQuantity' column to float")
        virtual_machines['UsageQuantity'] = virtual_machines['UsageQuantity'].astype(float)
        
        # Group the DataFrame by 'MeterName' and aggregate other columns
        logging.info("Grouping and aggregating data by 'MeterName'")
        virtual_machines = virtual_machines.groupby('MeterName').agg({
            'UsageDateTime': 'first',
            'SubscriptionName': 'first',
            'ProductName': 'first',
            'MeterCategory': 'first',
            'MeterSubcategory': 'first',
            'UsageQuantity': 'sum',
            'ResourceRate': 'first',
            'PreTaxCost': 'sum',
            'Currency': 'first',
            'ResourceLocation': 'first'
        }).reset_index()

        # Calculate the quantity of instances by dividing 'UsageQuantity' by the number of hours in the current month
        logging.info("Calculating quantity of instances")
        virtual_machines['Qty Instances'] = round(virtual_machines['UsageQuantity'] / config['month_hours'], 2)

        # Calculate the quantity of reservations for 1 year and 3 years
        logging.info("Calculating 1-year and 3-year reservations")
        virtual_machines['Reserves 1Y'] = np.where(
            np.mod(virtual_machines['Qty Instances'], 1) >= config['threshold_virtual_machines_1y_reserv'], 
            np.ceil(virtual_machines['Qty Instances']), 
            np.floor(virtual_machines['Qty Instances'])
        )
        virtual_machines['Reserves 3Y'] = np.where(
            np.mod(virtual_machines['Qty Instances'], 1) >= config['threshold_virtual_machines_3y_reserv'], 
            np.ceil(virtual_machines['Qty Instances']), 
            np.floor(virtual_machines['Qty Instances'])
        )
        
        # Calculate the number of reservation hours for 1 year and 3 years
        logging.info("Calculating reservation hours for 1 year and 3 years")
        virtual_machines['Reserv Hours 1Y'] = virtual_machines['Reserves 1Y'] * config['reservation_hours']
        virtual_machines['Reserv Hours 3Y'] = virtual_machines['Reserves 3Y'] * config['reservation_hours']

        # Calculate the residual value of the reserves that go to pay as you go 
        logging.info("Calculating PAYG hours for 1 year and 3 years")
        virtual_machines['PAYG Hours 1Y'] = virtual_machines['UsageQuantity'] - virtual_machines['Reserv Hours 1Y']
        virtual_machines['PAYG Hours 3Y'] = virtual_machines['UsageQuantity'] - virtual_machines['Reserv Hours 3Y']

        # Calculate the unit price for a 1 year and 3 years reservation with respective discounts
        logging.info("Calculating unit prices for 1 year and 3 years reservations")
        virtual_machines['UP 1Year Reserv'] = virtual_machines['ResourceRate'] * config['virtual_machines_1y_reserv']
        virtual_machines['UP 3Years Reserv'] = virtual_machines['ResourceRate'] * config['virtual_machines_3y_reserv']

        # Calculate the total amount for the 1 and 3 year reservation
        logging.info("Calculating total amounts for 1 year and 3 years reservations")
        virtual_machines['Amount 1Year Reserv'] = np.where(
            virtual_machines['PAYG Hours 1Y'] < 0, 
            virtual_machines['Reserv Hours 1Y'] * virtual_machines['UP 1Year Reserv'], 
            virtual_machines['Reserv Hours 1Y'] * virtual_machines['UP 1Year Reserv'] + virtual_machines['PAYG Hours 1Y'] * virtual_machines['ResourceRate']
        )
        virtual_machines['Amount 3Years Reserv'] = np.where(
            virtual_machines['PAYG Hours 3Y'] < 0, 
            virtual_machines['Reserv Hours 3Y'] * virtual_machines['UP 3Years Reserv'], 
            virtual_machines['Reserv Hours 3Y'] * virtual_machines['UP 3Years Reserv'] + virtual_machines['PAYG Hours 3Y'] * virtual_machines['ResourceRate']
        )

        # Calculate savings for 1 and 3 year reserves
        logging.info("Calculating savings for 1 year and 3 years reservations")
        virtual_machines['Saving 1Y Reserv'] = np.where(
            virtual_machines['PreTaxCost'] - virtual_machines['Amount 1Year Reserv'] > 0, 
            virtual_machines['PreTaxCost'] - virtual_machines['Amount 1Year Reserv'], 
            0
        )
        virtual_machines['Saving 3Y Reserv'] = np.where(
            virtual_machines['PreTaxCost'] - virtual_machines['Amount 3Years Reserv'] > 0, 
            virtual_machines['PreTaxCost'] - virtual_machines['Amount 3Years Reserv'], 
            0
        )

        # Calculate total savings based on conditions
        logging.info("Calculating final savings")
        conditions = [
            (virtual_machines['Saving 1Y Reserv'] > 0) & (virtual_machines['Saving 3Y Reserv'] > 0),
            (virtual_machines['Saving 1Y Reserv'] > 0) & (virtual_machines['Saving 3Y Reserv'] <= 0),
            (virtual_machines['Saving 1Y Reserv'] <= 0) & (virtual_machines['Saving 3Y Reserv'] > 0)
        ]
        results = [
            virtual_machines[['Saving 1Y Reserv', 'Saving 3Y Reserv']].mean(axis=1),
            virtual_machines['Saving 1Y Reserv'],
            virtual_machines['Saving 3Y Reserv']
        ]
        virtual_machines['Final Saving'] = np.select(conditions, results, default=0)

        logging.info("Virtual machines processing completed successfully")
        return virtual_machines

    except Exception as e:
        logging.error(f"Error processing virtual machines: {e}")
        return None

def storage(dir: Path, config):
    
    """
    Processes Azure API data source for storage.
    Parameters:
    dir (Path): The directory path of the CSV file to be processed.
    config (dict): Configuration parameters including storage_lrs, storage_zrs, storage_data_write_zrs,
                   storage_data_write_grs, storage_data_write_ra-grs.
    Returns:
    DataFrame: The processed DataFrame with calculated fields for storage.
    """
    
    logging.info("Starting storage processing")

    try:
        # Load the CSV file into a DataFrame
        logging.info("Loading and normalizing data source")
        metering_data = data_source_normalization(dir)
        
        if metering_data is None:
            logging.error("Failed to load and normalize data source")
            return None
        
        # Filter out rows where 'MeterCategory' is not 'Storage'
        logging.info("Filtering data for 'Storage' category")
        storage = metering_data[metering_data['MeterCategory'] == 'Storage'].copy()
        
        if storage.empty:
            logging.warning("No data found for 'Storage'")
            # Return an empty DataFrame with the necessary columns
            columns = [
                'MeterName', 'InstanceId', 'UsageDateTime', 'SubscriptionName', 'ProductName',
                'MeterCategory', 'MeterSubcategory', 'UsageQuantity', 'ResourceRate', 'PreTaxCost',
                'Currency', 'ResourceLocation', 'New UP PAYG', 'New PAYG Amount', 'Saving PAYG',
                'Saving 1Y Reserv', 'Saving 3Y Reserv', 'Final Saving'
            ]
            return pd.DataFrame(columns=columns)

        # Add 'InstanceId' if not present
        if 'InstanceId' not in storage.columns:
            logging.info("Adding 'InstanceId' column")
            storage['InstanceId'] = storage['Publisher Subscription ID'].astype(str) + storage['Resource Group Name'].astype(str) + storage['Resource Name'].astype(str)

        # Group the DataFrame by 'MeterName' and 'InstanceId' and aggregate other columns
        logging.info("Grouping and aggregating data by 'MeterName' and 'InstanceId'")
        storage = storage.groupby(['MeterName', 'InstanceId']).agg({
            'UsageDateTime': 'first',
            'SubscriptionName': 'first',
            'ProductName': 'first',
            'MeterCategory': 'first',
            'MeterSubcategory': 'first',
            'UsageQuantity': 'sum',
            'ResourceRate': 'first',
            'PreTaxCost': 'sum',
            'Currency': 'first',
            'ResourceLocation': 'first'
        }).reset_index()

        # Define more conditions to calculate the new unit price for PAYG
        logging.info("Defining conditions and calculating new unit prices for PAYG")
        conditions = [
            (storage['MeterName'].str.contains('LRS')) & (storage['ProductName'].str.contains('Premium')) & (storage['ProductName'].str.contains('Disks')),
            (storage['MeterName'].str.contains('ZRS')) & (storage['ProductName'].str.contains('Premium')) & (storage['ProductName'].str.contains('Disks'))
        ]
        results = [
            storage['ResourceRate'] * config['storage_lrs'],
            storage['ResourceRate'] * config['storage_zrs']
        ]
        storage['New UP PAYG'] = np.select(conditions, results, default=np.nan)

        # Replace null values with an empty string
        logging.info("Replacing null values with empty strings")
        storage['MeterName'] = storage['MeterName'].fillna('')
        storage['ProductName'] = storage['ProductName'].fillna('')

        # Define more conditions to calculate the new unit price for PAYG
        logging.info("Defining additional conditions and calculating new unit prices for PAYG")
        conditions = [
            (storage['MeterName'].str.contains('LRS')) & ((storage['MeterName'].str.contains('Data')) | (storage['MeterName'].str.contains('Write'))) & ~(storage['ProductName'].str.contains('Premium')) & ~(storage['ProductName'].str.contains('Disks')),
            (storage['MeterName'].str.contains('ZRS')) & ((storage['MeterName'].str.contains('Data')) | (storage['MeterName'].str.contains('Write'))) & ~(storage['ProductName'].str.contains('Premium')) & ~(storage['ProductName'].str.contains('Disks')),
            (storage['MeterName'].str.contains('GRS')) & ((storage['MeterName'].str.contains('Data')) | (storage['MeterName'].str.contains('Write'))) & ~(storage['ProductName'].str.contains('Premium')) & ~(storage['ProductName'].str.contains('Disks')),
            (storage['MeterName'].str.contains('RA-GRS')) & ((storage['MeterName'].str.contains('Data')) | (storage['MeterName'].str.contains('Write'))) & ~(storage['ProductName'].str.contains('Premium')) & ~(storage['ProductName'].str.contains('Disks'))
        ]
        results = [
            storage['ResourceRate'],
            storage['ResourceRate'] * config['storage_data_write_zrs'],
            storage['ResourceRate'] * config['storage_data_write_grs'],
            storage['ResourceRate'] * config['storage_data_write_ra-grs']
        ]
        storage['New UP PAYG'] = np.select(conditions, results, default=storage['New UP PAYG'])

        # Calculate the new PAYG amount
        logging.info("Calculating new PAYG amounts")
        storage['New PAYG Amount'] = storage['UsageQuantity'] * storage['New UP PAYG'] 

        # Calculate savings with PAYG
        logging.info("Calculating savings with PAYG")
        storage['Saving PAYG'] = np.where(storage['PreTaxCost'] - storage['New PAYG Amount'] > 0, storage['PreTaxCost'] - storage['New PAYG Amount'], 0)

        # Initialize savings columns for 1 and 3 year reservations with 0
        logging.info("Initializing savings columns for 1 and 3 year reservations")
        storage['Saving 1Y Reserv'] = 0
        storage['Saving 3Y Reserv'] = 0

        # Calculate final savings
        logging.info("Calculating final savings")
        storage['Final Saving'] = storage['Saving PAYG'].fillna(0)

        logging.info("Storage processing completed successfully")
        return storage

    except Exception as e:
        logging.error(f"Error processing storage: {e}")
        return None

def app_services(dir: Path, config):
    
    """
    Processes Azure API data source for app services.
    Parameters:
    dir (Path): The directory path of the CSV file to be processed.
    config (dict): Configuration parameters.
    Returns:
    DataFrame: The processed DataFrame with calculated fields for app services.
    """
    
    logging.info("Starting app services processing")

    try:
        # Load the CSV file into a DataFrame
        logging.info("Loading and normalizing data source")
        metering_data = data_source_normalization(dir)
        
        if metering_data is None:
            logging.error("Failed to load and normalize data source")
            return None
        
        # Filter the dataframe to only include rows where 'MeterCategory' is 'Azure App Service'
        logging.info("Filtering data for 'Azure App Service' category")
        app_services = metering_data[metering_data['MeterCategory'] == 'Azure App Service'].copy()
        
        if 'Azure App Service' not in app_services['MeterCategory'].values:
            logging.warning("No data found for 'Azure App Service'")
            # Return an empty DataFrame with the necessary columns
            columns = [
                'MeterName', 'InstanceId', 'UsageDateTime', 'SubscriptionName', 'ProductName',
                'MeterCategory', 'MeterSubcategory', 'UsageQuantity', 'ResourceRate', 'PreTaxCost',
                'Currency', 'ResourceLocation', 'Qty Instances', 'New Meter Name', 'Reserve Eligible',
                'Reserves 1Y', 'Reserves 3Y', 'Reserv Hours PAYG', 'Reserv Hours 1Y', 'Reserv Hours 3Y',
                'New UP PAYG', 'UP 1Y Reserv', 'UP 3Y Reserv', 'New Amount PAYG', 'Amount 1Year Reserv',
                'Amount 3Years Reserv', 'Saving PAYG', 'Saving 1Y Reserv', 'Saving 3Y Reserv', 'Final Saving'
            ]
            return pd.DataFrame(columns=columns)

        # Add 'InstanceId' if not present
        if 'InstanceId' not in app_services.columns:
            logging.info("Adding 'InstanceId' column")
            app_services['InstanceId'] = app_services['Publisher Subscription ID'].astype(str) + app_services['Resource Group Name'].astype(str) + app_services['Resource Name'].astype(str)

        # Group the data by 'MeterName' and 'InstanceId', and aggregate the other columns as specified
        logging.info("Grouping and aggregating data by 'MeterName' and 'InstanceId'")
        app_services = app_services.groupby(['MeterName', 'InstanceId']).agg({
            'UsageDateTime': 'first',
            'SubscriptionName': 'first',
            'ProductName': 'first',
            'MeterCategory': 'first',
            'MeterSubcategory': 'first',
            'UsageQuantity': 'sum',
            'ResourceRate': 'first',
            'PreTaxCost': 'sum',
            'Currency': 'first',
            'ResourceLocation': 'first'
        }).reset_index()

        # Calculate the quantity of instances by dividing 'UsageQuantity' by number of hours in a month
        logging.info("Calculating quantity of instances")
        app_services['Qty Instances'] = round(app_services['UsageQuantity'] / config['month_hours'], 2)

        # Define conditions to calculate the new meter name
        logging.info("Calculating new meter names based on conditions")
        conditions = [
            app_services['MeterName'] == "P1 v2 App",
            app_services['MeterName'] == "P2 v2 App",
            app_services['MeterName'] == "P3 v2 App",
            app_services['MeterName'] == "I2",
            app_services['MeterName'] == "I3"
        ]
        results = [
            "P0v3 App",
            "P1v3 App",
            "P2v3 App",
            "I1 v2",
            "I2 v2" 
        ]
        app_services['New Meter Name'] = np.select(conditions, results, default=app_services['MeterName'])

        # Define conditions to check if it is eligible for reservation
        logging.info("Checking reservation eligibility based on new meter names")
        conditions = [
            app_services['New Meter Name'] == "P0v3 App",
            app_services['New Meter Name'] == "P1v3 App",
            app_services['New Meter Name'] == "P2v3 App",
            app_services['New Meter Name'] == "I1 v2",
            app_services['New Meter Name'] == "I2 v2"
        ]
        results = ['Yes', 'Yes', 'Yes', 'Yes', 'Yes']
        app_services['Reserve Eligible'] = np.select(conditions, results, default='No')

        # Define conditions to calculate 1-year reserves
        logging.info("Calculating 1-year reserves based on conditions")
        conditions = [
            (app_services['New Meter Name'] == "P0v3 App") & (app_services['Qty Instances'] >= config['threshold_app_services_1y_reserv_p0v3']),   
            (app_services['New Meter Name'] == "P0v3 App") & (app_services['Qty Instances'] < config['threshold_app_services_1y_reserv_p0v3']),
            (app_services['New Meter Name'] == "P1v3 App") & (app_services['Qty Instances'] >= config['threshold_app_services_1y_reserv_p1v3']),
            (app_services['New Meter Name'] == "P1v3 App") & (app_services['Qty Instances'] < config['threshold_app_services_1y_reserv_p1v3']),
            (app_services['New Meter Name'] == "P2v3 App") & (app_services['Qty Instances'] >= config['threshold_app_services_1y_reserv_p2v3']), 
            (app_services['New Meter Name'] == "P2v3 App") & (app_services['Qty Instances'] < config['threshold_app_services_1y_reserv_p2v3']),
            (app_services['New Meter Name'] == "I1 v2") & (app_services['Qty Instances'] >= config['threshold_app_services_1y_reserv_i1v2']),
            (app_services['New Meter Name'] == "I1 v2") & (app_services['Qty Instances'] < config['threshold_app_services_1y_reserv_i1v2']),
            (app_services['New Meter Name'] == "I2 v2") & (app_services['Qty Instances'] >= config['threshold_app_services_1y_reserv_i2v2']),
            (app_services['New Meter Name'] == "I2 v2") & (app_services['Qty Instances'] < config['threshold_app_services_1y_reserv_i2v2'])
        ]
        results = [
            np.ceil(app_services['Qty Instances']),
            np.floor(app_services['Qty Instances']),
            np.ceil(app_services['Qty Instances']),
            np.floor(app_services['Qty Instances']),
            np.ceil(app_services['Qty Instances']),
            np.floor(app_services['Qty Instances']),
            np.ceil(app_services['Qty Instances']),
            np.floor(app_services['Qty Instances']),
            np.ceil(app_services['Qty Instances']),
            np.floor(app_services['Qty Instances']),
        ]
        app_services['Reserves 1Y'] = np.select(conditions, results, default=app_services['Qty Instances'])

        # Define conditions to calculate 3-year reserves
        logging.info("Calculating 3-year reserves based on conditions")
        conditions = [
            (app_services['New Meter Name'] == "P0v3 App") & (app_services['Qty Instances'] >= config['threshold_app_services_3y_reserv_p0v3']),   
            (app_services['New Meter Name'] == "P0v3 App") & (app_services['Qty Instances'] < config['threshold_app_services_3y_reserv_p0v3']),
            (app_services['New Meter Name'] == "P1v3 App") & (app_services['Qty Instances'] >= config['threshold_app_services_3y_reserv_p1v3']),
            (app_services['New Meter Name'] == "P1v3 App") & (app_services['Qty Instances'] < config['threshold_app_services_3y_reserv_p1v3']),
            (app_services['New Meter Name'] == "P2v3 App") & (app_services['Qty Instances'] >= config['threshold_app_services_3y_reserv_p2v3']), 
            (app_services['New Meter Name'] == "P2v3 App") & (app_services['Qty Instances'] < config['threshold_app_services_3y_reserv_p2v3']),
            (app_services['New Meter Name'] == "I1 v2") & (app_services['Qty Instances'] >= config['threshold_app_services_3y_reserv_i1v2']),
            (app_services['New Meter Name'] == "I1 v2") & (app_services['Qty Instances'] < config['threshold_app_services_3y_reserv_i1v2']),
            (app_services['New Meter Name'] == "I2 v2") & (app_services['Qty Instances'] >= config['threshold_app_services_3y_reserv_i2v2']),
            (app_services['New Meter Name'] == "I2 v2") & (app_services['Qty Instances'] < config['threshold_app_services_3y_reserv_i2v2'])
        ]
        results = [
            np.ceil(app_services['Qty Instances']),
            np.floor(app_services['Qty Instances']),
            np.ceil(app_services['Qty Instances']),
            np.floor(app_services['Qty Instances']),
            np.ceil(app_services['Qty Instances']),
            np.floor(app_services['Qty Instances']),
            np.ceil(app_services['Qty Instances']),
            np.floor(app_services['Qty Instances']),
            np.ceil(app_services['Qty Instances']),
            np.floor(app_services['Qty Instances']),
        ]
        app_services['Reserves 3Y'] = np.select(conditions, results, default=app_services['Qty Instances'])

        # Calculate the reserve hours for PAYG, 1 year, and 3 years
        logging.info("Calculating reserve hours for PAYG, 1 year, and 3 years")
        app_services['Reserv Hours PAYG'] = app_services['Qty Instances'] * config['month_hours']
        app_services['Reserv Hours 1Y'] = app_services['Reserves 1Y'] * config['reservation_hours']
        app_services['Reserv Hours 3Y'] = app_services['Reserves 3Y'] * config['reservation_hours']

        # Define conditions to calculate the new unit price for PAYG
        logging.info("Calculating new unit prices for PAYG")
        conditions = [
            app_services['New Meter Name'] == "P0v3 App",
            app_services['New Meter Name'] == "P1v3 App",
            app_services['New Meter Name'] == "P2v3 App",
            app_services['New Meter Name'] == "I1 v2",
            app_services['New Meter Name'] == "I2 v2"
        ]
        results = [
            app_services['ResourceRate'],
            app_services['ResourceRate'] * config['app_services_payg_p1v3'],
            app_services['ResourceRate'] * config['app_services_payg_p2v3'],
            app_services['ResourceRate'] * config['app_services_payg_i1v2'],
            app_services['ResourceRate'] * config['app_services_payg_i2v2']
        ]
        app_services['New UP PAYG'] = np.select(conditions, results, default=app_services['ResourceRate'])

        # Define conditions to calculate the unit price for 1-year reserve
        logging.info("Calculating unit prices for 1-year reserves")
        conditions = [
            app_services['New Meter Name'] == "P0v3 App",
            app_services['New Meter Name'] == "P1v3 App",
            app_services['New Meter Name'] == "P2v3 App",
            app_services['New Meter Name'] == "I1 v2",
            app_services['New Meter Name'] == "I2 v2"
        ]
        results = [
            app_services['ResourceRate'] * config['app_services_1y_reserv_p0v3'],
            app_services['ResourceRate'] * config['app_services_1y_reserv_p1v3'],
            app_services['ResourceRate'] * config['app_services_1y_reserv_p2v3'],
            app_services['ResourceRate'] * config['app_services_1y_reserv_i1v2'],
            app_services['ResourceRate'] * config['app_services_1y_reserv_i2v2']
        ]
        app_services['UP 1Y Reserv'] = np.select(conditions, results, default=app_services['ResourceRate'])

        # Define conditions to calculate the unit price for 3-year reserve
        logging.info("Calculating unit prices for 3-year reserves")
        conditions = [
            app_services['New Meter Name'] == "P0v3 App",
            app_services['New Meter Name'] == "P1v3 App",
            app_services['New Meter Name'] == "P2v3 App",
            app_services['New Meter Name'] == "I1 v2",
            app_services['New Meter Name'] == "I2 v2"
        ]
        results = [
            app_services['ResourceRate'] * config['app_services_3y_reserv_p0v3'],
            app_services['ResourceRate'] * config['app_services_3y_reserv_p1v3'],
            app_services['ResourceRate'] * config['app_services_3y_reserv_p2v3'],
            app_services['ResourceRate'] * config['app_services_3y_reserv_i1v2'],
            app_services['ResourceRate'] * config['app_services_3y_reserv_i2v2']
        ]
        app_services['UP 3Y Reserv'] = np.select(conditions, results, default=app_services['ResourceRate'])

        # Calculate the new amount for PAYG
        logging.info("Calculating new amounts for PAYG")
        app_services['New Amount PAYG'] = app_services['New UP PAYG'] * app_services['UsageQuantity']
        
        # Calculate the amount for 1-year reserve
        logging.info("Calculating amounts for 1-year reserves")
        app_services['Amount 1Year Reserv'] = np.where(app_services['Reserve Eligible'] == 'Yes', app_services['Reserv Hours 1Y'] * app_services['UP 1Y Reserv'], app_services['PreTaxCost'])

        # Calculate the amount for 3-year reserve
        logging.info("Calculating amounts for 3-year reserves")
        app_services['Amount 3Years Reserv'] = np.where(app_services['Reserve Eligible'] == 'Yes', app_services['Reserv Hours 3Y'] * app_services['UP 3Y Reserv'], app_services['PreTaxCost'])

        # Calculate the savings for PAYG
        logging.info("Calculating savings for PAYG")
        app_services['Saving PAYG'] = np.where(app_services['PreTaxCost'] - app_services['New Amount PAYG'] > 0, app_services['PreTaxCost'] - app_services['New Amount PAYG'], 0)
        
        # Calculate the savings for 1-year reserve
        logging.info("Calculating savings for 1-year reserves")
        app_services['Saving 1Y Reserv'] = np.where(app_services['PreTaxCost'] - app_services['Amount 1Year Reserv'] > 0, app_services['PreTaxCost'] - app_services['Amount 1Year Reserv'], 0)
        
        # Calculate the savings for 3-year reserve
        logging.info("Calculating savings for 3-year reserves")
        app_services['Saving 3Y Reserv'] = np.where(app_services['PreTaxCost'] - app_services['Amount 3Years Reserv'] > 0, app_services['PreTaxCost'] - app_services['Amount 3Years Reserv'], 0)

        # Calculate the final saving as the average of the savings for PAYG, 1-year reserve, and 3-year reserve
        logging.info("Calculating final savings")
        app_services['Final Saving'] = app_services[['Saving PAYG', 'Saving 1Y Reserv', 'Saving 3Y Reserv']].mean(axis=1)
         
        logging.info("App services processing completed successfully")
        return app_services

    except Exception as e:
        logging.error(f"Error processing app services: {e}")
        return None

def backup(dir: Path, config):
    """
    Processes Azure API data source for backups.
    Parameters:
    dir (Path): The directory path of the CSV file to be processed.
    config (dict): Configuration parameters.
    Returns:
    DataFrame: The processed DataFrame with calculated fields for backups.
    """
    
    logging.info("Starting backup processing")

    try:
        # Load the CSV file into a DataFrame
        logging.info("Loading and normalizing data source")
        metering_data = data_source_normalization(dir)
        
        if metering_data is None:
            logging.error("Failed to load and normalize data source")
            return None
        
        # Filter the dataframe to only include rows where 'MeterCategory' is 'Backup'
        logging.info("Filtering data for 'Backup' category")
        backup = metering_data[metering_data['MeterCategory'] == 'Backup'].copy()
        
        if backup.empty:
            logging.warning("No data found for 'Backup'")
            # Return an empty DataFrame with the necessary columns
            columns = [
                'MeterName', 'InstanceId', 'UsageDateTime', 'SubscriptionName', 'ProductName',
                'MeterCategory', 'MeterSubcategory', 'UsageQuantity', 'ResourceRate', 'PreTaxCost',
                'Currency', 'ResourceLocation', 'New UP PAYG', 'New PAYG Amount', 'Saving PAYG',
                'Saving 1Y Reserv', 'Saving 3Y Reserv', 'Final Saving'
            ]
            return pd.DataFrame(columns=columns)

        # Create a mask to filter out resource names that do not contain "Prod"
        logging.info("Filtering out non-production resources")
        mask = ~backup['SubscriptionName'].str.contains('PRD|Prod|Production', case=False, na=False)

        # Apply the mask to the dataframe
        final_backup = backup[mask]

        if final_backup.empty:
            logging.warning("No production data found for 'Backup'")
            return pd.DataFrame(columns=columns)

        # Create a copy of the filtered dataframe
        backup_df = final_backup.copy()

        # Add 'InstanceId' if not present
        if 'InstanceId' not in backup_df.columns:
            logging.info("Adding 'InstanceId' column")
            backup_df['InstanceId'] = backup_df['Publisher Subscription ID'].astype(str) + backup_df['Resource Group Name'].astype(str) + backup_df['Resource Name'].astype(str)

        # Group the dataframe by 'MeterName' and 'InstanceId', and aggregate the other columns
        logging.info("Grouping and aggregating data by 'MeterName' and 'InstanceId'")
        backup_df = backup_df.groupby(['MeterName', 'InstanceId']).agg({
            'UsageDateTime': 'first',
            'SubscriptionName': 'first',
            'ProductName': 'first',
            'MeterCategory': 'first',
            'MeterSubcategory': 'first',
            'UsageQuantity': 'sum',
            'ResourceRate': 'first',
            'PreTaxCost': 'sum',
            'Currency': 'first',
            'ResourceLocation': 'first'
        }).reset_index()

        # Define the conditions for calculating the new unit price
        logging.info("Calculating new unit prices based on conditions")
        conditions = [
            backup_df['MeterName'].str.contains('LRS'),
            backup_df['MeterName'].str.contains('ZRS'),
            backup_df['MeterName'].str.contains('GRS'),
            backup_df['MeterName'].str.contains('RA-GRS'),
        ]

        # Define the results for each condition
        results = [
            backup_df['ResourceRate'],
            backup_df['ResourceRate'] * config['backups_payg_zrs'],
            backup_df['ResourceRate'] * config['backups_payg_grs'],
            backup_df['ResourceRate'] * config['backups_payg_ra-grs'],
        ]

        # Calculate the new unit price based on the defined conditions and results
        backup_df['New UP PAYG'] = np.select(conditions, results, default=backup_df['ResourceRate'])

        # Calculate the new PAYG amount
        logging.info("Calculating new PAYG amounts")
        backup_df['New PAYG Amount'] = backup_df['UsageQuantity'] * backup_df['New UP PAYG'] 

        # Calculate the PAYG savings
        logging.info("Calculating PAYG savings")
        backup_df['Saving PAYG'] = np.where(backup_df['PreTaxCost'] - backup_df['New PAYG Amount'] > 0, backup_df['PreTaxCost'] - backup_df['New PAYG Amount'], 0)

        # Initialize the 'Saving 1Y Reserv' and 'Saving 3Y Reserv' columns with 0
        logging.info("Initializing savings columns for 1-year and 3-year reservations")
        backup_df['Saving 1Y Reserv'] = 0
        backup_df['Saving 3Y Reserv'] = 0

        # Calculate the final saving
        logging.info("Calculating final savings")
        backup_df['Final Saving'] = backup_df['Saving PAYG']
        
        logging.info("Backup processing completed successfully")
        return backup_df

    except Exception as e:
        logging.error(f"Error processing backup: {e}")
        return None

def application_gateway(dir: Path, config):
    """
    Processes Azure or API data source for Application Gateway.
    Parameters:
    dir (Path): The directory path of the Excel file to be processed.
    config (dict): A dictionary containing configuration parameters.
    Returns:
    DataFrame: The processed DataFrame with calculated fields for Application Gateway.
    """
    
    logging.info("Starting application gateway processing")

    try:
        # Load the CSV file into a DataFrame
        logging.info("Loading and normalizing data source")
        metering_data = data_source_normalization(dir)
        
        if metering_data is None:
            logging.error("Failed to load and normalize data source")
            return None
        
        # Filter the dataframe to only include rows where 'MeterCategory' is 'Application Gateway'
        logging.info("Filtering data for 'Application Gateway' category")
        application_gateway = metering_data[metering_data['MeterCategory'] == 'Application Gateway'].copy()
        
        if application_gateway.empty:
            logging.warning("No data found for 'Application Gateway'")
            # Return an empty DataFrame with the necessary columns
            columns = [
                'MeterName', 'InstanceId', 'UsageDateTime', 'SubscriptionName', 'ProductName',
                'MeterCategory', 'MeterSubcategory', 'UsageQuantity', 'ResourceRate', 'PreTaxCost',
                'Currency', 'ResourceLocation', 'Qty Instances', 'Total Hours per Id', 'Usage %',
                'New PAYG Amount', 'Saving PAYG', 'Final Saving'
            ]
            return pd.DataFrame(columns=columns)

        # Create a mask to filter out resource names that do not contain "Prod"
        logging.info("Filtering out non-production resources")
        mask = ~application_gateway['SubscriptionName'].str.contains('PRD|Prod|Production', case=False, na=False)

        # Apply the mask to the dataframe
        final_application_gateway = application_gateway[mask]

        if final_application_gateway.empty:
            logging.warning("No production data found for 'Application Gateway'")
            columns = [
                'MeterName', 'InstanceId', 'UsageDateTime', 'SubscriptionName', 'ProductName',
                'MeterCategory', 'MeterSubcategory', 'UsageQuantity', 'ResourceRate', 'PreTaxCost',
                'Currency', 'ResourceLocation', 'Qty Instances', 'Total Hours per Id', 'Usage %',
                'New PAYG Amount', 'Saving PAYG', 'Final Saving'
            ]
            return pd.DataFrame(columns=columns)

        # Create a copy of the filtered dataframe
        app_gateway_df = final_application_gateway.copy()

        # Add 'InstanceId' if not present
        if 'InstanceId' not in app_gateway_df.columns:
            logging.info("Adding 'InstanceId' column")
            app_gateway_df['InstanceId'] = app_gateway_df['Publisher Subscription ID'].astype(str) + app_gateway_df['Resource Group Name'].astype(str) + app_gateway_df['Resource Name'].astype(str)

        # Group the dataframe by 'MeterName' and 'InstanceId', and aggregate the other columns
        logging.info("Grouping and aggregating data by 'MeterName' and 'InstanceId'")
        app_gateway_df = app_gateway_df.groupby(['MeterName', 'InstanceId']).agg({
            'UsageDateTime': 'first',
            'SubscriptionName': 'first',
            'ProductName': 'first',
            'MeterCategory': 'first',
            'MeterSubcategory': 'first',
            'UsageQuantity': 'sum',
            'ResourceRate': 'first',
            'PreTaxCost': 'sum',
            'Currency': 'first',
            'ResourceLocation': 'first'
        }).reset_index()

        # Calculate the quantity of instances by dividing 'UsageQuantity' by 720 hours (the number of hours in the current month)
        logging.info("Calculating quantity of instances")
        app_gateway_df['Qty Instances'] = np.round(app_gateway_df['UsageQuantity'] / config['month_hours'], 0)

        # Calculate the total hours per ID
        logging.info("Calculating total hours per ID")
        app_gateway_df['Total Hours per Id'] = app_gateway_df['Qty Instances'] * config['month_hours']

        # Calculate the usage percentage
        logging.info("Calculating usage percentage")
        app_gateway_df['Usage %'] = np.round(app_gateway_df['UsageQuantity'] / app_gateway_df['Total Hours per Id'], 2)

        # Calculate the new PAYG amount
        logging.info("Calculating new PAYG amounts")
        app_gateway_df['New PAYG Amount'] = np.where(
            app_gateway_df['Usage %'] >= config['threshold_application_gateway'], 
            app_gateway_df['PreTaxCost'] * config['threshold_application_gateway'] / app_gateway_df['Usage %'], 
            app_gateway_df['PreTaxCost']
        )

        # Calculate the PAYG savings
        logging.info("Calculating PAYG savings")
        app_gateway_df['Saving PAYG'] = app_gateway_df['PreTaxCost'] - app_gateway_df['New PAYG Amount']

        # Calculate the final saving
        logging.info("Calculating final savings")
        app_gateway_df['Final Saving'] = app_gateway_df['Saving PAYG']

        logging.info("Application gateway processing completed successfully")
        return app_gateway_df

    except Exception as e:
        logging.error(f"Error processing application gateway: {e}")
        return None

def sql_database(dir: Path, config):
    """
    Processes Azure or API data source for SQL Database.
    Parameters:
    dir (Path): The directory path of the Excel file to be processed.
    config (dict): A dictionary containing configuration parameters.
    Returns:
    DataFrame: The processed DataFrame with calculated fields for SQL Database.
    """
    
    logging.info("Starting SQL Database processing")

    try:
        # Load the CSV file into a DataFrame
        logging.info("Loading and normalizing data source")
        metering_data = data_source_normalization(dir)
        
        if metering_data is None:
            logging.error("Failed to load and normalize data source")
            return None
        
        # Filter the dataframe to only include rows where 'MeterCategory' is 'SQL Database'
        logging.info("Filtering data for 'SQL Database' category")
        sql_database = metering_data[metering_data['MeterCategory'] == 'SQL Database'].copy()
        
        if sql_database.empty:
            logging.warning("No data found for 'SQL Database'")
            # Return an empty DataFrame with the necessary columns
            columns = [
                'MeterName', 'InstanceId', 'UsageDateTime', 'SubscriptionName', 'ProductName',
                'MeterCategory', 'MeterSubcategory', 'UsageQuantity', 'ResourceRate', 'PreTaxCost',
                'Currency', 'vCores', 'ResourceLocation', 'UsageHours', 'PAYG Hours 1Y', 'PAYG Hours 3Y',
                'UP 1Year Reserv', 'UP 3Years Reserv', 'Amount 1Year Reserv', 'Amount 3Years Reserv',
                'Saving 1Y Reserv', 'Saving 3Y Reserv', 'Final Saving'
            ]
            return pd.DataFrame(columns=columns)

        # Filter the dataframe to only include rows where 'MeterName' is 'vCore'
        logging.info("Filtering data for 'vCore' meter name")
        sql_database = sql_database[sql_database['MeterName'] == 'vCore']

        # Filter the dataframe to only include rows where 'MeterSubcategory' matches specific patterns
        logging.info("Filtering data for specific 'MeterSubcategory' patterns")
        filter_values = [
            'Single/Elastic Pool General Purpose - Compute Gen5', 
            'Single/Elastic Pool Business Critical - Compute Gen5', 
            'Single/Elastic Pool Hyperscale - Compute Gen5'
        ]
        pattern = '|'.join([f'({value})' for value in filter_values])
        sql_database = sql_database[sql_database['MeterSubcategory'].str.contains(pattern)]

        # Create a copy of the filtered dataframe
        sql_database_copy = sql_database.copy()

        logging.info("Extracting 'vCores' from 'AdditionalInfo'")
        sql_database_copy['vCores'] = sql_database_copy['AdditionalInfo'].apply(extract_vcores)

        logging.info("Standardizing 'InstanceId' column to uppercase")
        sql_database_copy['InstanceId'] = sql_database_copy['InstanceId'].str.upper()

        # Group the dataframe by 'MeterName' and 'InstanceId', and aggregate the other columns
        logging.info("Grouping and aggregating data by 'MeterName' and 'InstanceId'")
        sql_database_df = sql_database_copy.groupby(['MeterName', 'InstanceId', 'AdditionalInfo']).agg({
            'UsageDateTime': 'first',
            'SubscriptionName': 'first',
            'ProductName': 'first',
            'MeterCategory': 'first',
            'MeterSubcategory': 'first',
            'UsageQuantity': 'sum',
            'ResourceRate': 'first',
            'PreTaxCost': 'sum',
            'Currency': 'first',
            'vCores': 'first',
            'ResourceLocation': 'first'
        }).reset_index()

        logging.info("Calculating usage hours")
        sql_database_df['UsageHours'] = sql_database_df['UsageQuantity'] / sql_database_df['vCores']

        # Calculate the residual value of the reserves that go to pay as you go 
        logging.info("Calculating PAYG hours for 1 year and 3 years")
        sql_database_df['PAYG Hours 1Y'] = sql_database_df['UsageHours'] - config['reservation_hours']
        sql_database_df['PAYG Hours 3Y'] = sql_database_df['UsageHours'] - config['reservation_hours']

        # Define the conditions and choices for 1-year reserves
        logging.info("Calculating unit prices for 1-year reservations")
        conditions_1y = [
            sql_database_df['UsageHours'] >= config['treshhold_sql_database_1y_reserv_80'] * config['month_hours'],
            (sql_database_df['UsageHours'] >= config['treshhold_sql_database_1y_reserv_70'] * config['month_hours']) & (sql_database_df['UsageHours'] < config['treshhold_sql_database_1y_reserv_80'] * config['month_hours']),
            sql_database_df['UsageHours'] < config['treshhold_sql_database_1y_reserv_70'] * config['month_hours']
        ]
        choices_1y = [
            sql_database_df['ResourceRate'] * config['sql_database_1y_reserv_80'],
            sql_database_df['ResourceRate'],
            sql_database_df['ResourceRate']
        ]
        sql_database_df['UP 1Year Reserv'] = np.select(conditions_1y, choices_1y, default=np.nan)

        # Define the conditions and choices for 3-year reserves
        logging.info("Calculating unit prices for 3-year reservations")
        conditions_3y = [
            sql_database_df['UsageHours'] >= config['treshhold_sql_database_3y_reserv_80'] * config['month_hours'],
            (sql_database_df['UsageHours'] >= config['treshhold_sql_database_3y_reserv_70'] * config['month_hours']) & (sql_database_df['UsageHours'] < config['treshhold_sql_database_3y_reserv_80'] * config['month_hours']),
            sql_database_df['UsageHours'] < config['treshhold_sql_database_3y_reserv_70'] * config['month_hours']
        ]
        choices_3y = [
            sql_database_df['ResourceRate'] * config['sql_database_3y_reserv_80'],
            sql_database_df['ResourceRate'] * config['sql_database_3y_reserv_70'],
            sql_database_df['ResourceRate']
        ]
        sql_database_df['UP 3Years Reserv'] = np.select(conditions_3y, choices_3y, default=np.nan)

        # Calculate the total amount for the 1 and 3 year reservation
        logging.info("Calculating total amounts for 1-year and 3-year reservations")
        sql_database_df['Amount 1Year Reserv'] = np.where(
            sql_database_df['PAYG Hours 1Y'] < 0, 
            config['reservation_hours'] * sql_database_df['UP 1Year Reserv'] * sql_database_df['vCores'], 
            config['reservation_hours'] * sql_database_df['UP 1Year Reserv'] * sql_database_df['vCores'] + sql_database_df['PAYG Hours 1Y'] * sql_database_df['ResourceRate'] * sql_database_df['vCores']
        )
        sql_database_df['Amount 3Years Reserv'] = np.where(
            sql_database_df['PAYG Hours 3Y'] < 0, 
            config['reservation_hours'] * sql_database_df['UP 3Years Reserv'] * sql_database_df['vCores'], 
            config['reservation_hours'] * sql_database_df['UP 3Years Reserv'] * sql_database_df['vCores'] + sql_database_df['PAYG Hours 3Y'] * sql_database_df['ResourceRate'] * sql_database_df['vCores']
        )

        # Calculate savings for 1 and 3 year reserves
        logging.info("Calculating savings for 1-year and 3-year reservations")
        sql_database_df['Saving 1Y Reserv'] = np.where(
            sql_database_df['PreTaxCost'] - sql_database_df['Amount 1Year Reserv'] > 0, 
            sql_database_df['PreTaxCost'] - sql_database_df['Amount 1Year Reserv'], 
            0
        )
        sql_database_df['Saving 3Y Reserv'] = np.where(
            sql_database_df['PreTaxCost'] - sql_database_df['Amount 3Years Reserv'] > 0, 
            sql_database_df['PreTaxCost'] - sql_database_df['Amount 3Years Reserv'], 
            0
        )

        # Calculate total savings based on conditions
        logging.info("Calculating final savings")
        conditions = [
            (sql_database_df['Saving 1Y Reserv'] > 0) & (sql_database_df['Saving 3Y Reserv'] > 0),
            (sql_database_df['Saving 1Y Reserv'] > 0) & (sql_database_df['Saving 3Y Reserv'] <= 0),
            (sql_database_df['Saving 1Y Reserv'] <= 0) & (sql_database_df['Saving 3Y Reserv'] > 0)
        ]
        results = [
            sql_database_df[['Saving 1Y Reserv', 'Saving 3Y Reserv']].mean(axis=1),
            sql_database_df['Saving 1Y Reserv'],
            sql_database_df['Saving 3Y Reserv']
        ]
        sql_database_df['Final Saving'] = np.select(conditions, results, default=0)

        logging.info("SQL Database processing completed successfully")
        return sql_database_df

    except Exception as e:
        logging.error(f"Error processing SQL Database: {e}")
        return None

def summary(vm: pd.DataFrame, storage: pd.DataFrame, app_services: pd.DataFrame, backups: pd.DataFrame, application_gateway: pd.DataFrame, sql_database: pd.DataFrame):
    try:
        logging.info("Initializing 'Saving PAYG' column in 'vm' dataframe.")
        vm['Saving PAYG'] = 0
    except Exception as e:
        logging.error("Error while initializing 'Saving PAYG' column in 'vm' dataframe: " + str(e))

    try:
        logging.info("Grouping 'vm' dataframe by 'MeterCategory'.")
        vm_summary = vm.groupby('MeterCategory').agg({
            'Saving PAYG': 'sum',
            'Saving 1Y Reserv': 'sum',
            'Saving 3Y Reserv': 'sum',
            'Final Saving': 'sum'
        }).reset_index()

        vm_summary = vm_summary.rename(columns={'MeterCategory':'Category'})

    except Exception as e:
        logging.error("Error while grouping 'vm' dataframe: " + str(e))

    try:
        logging.info("Grouping 'storage' dataframe by 'MeterCategory'.")
        storage_summary = storage.groupby('MeterCategory').agg({
            'Saving PAYG': 'sum',
            'Saving 1Y Reserv': 'sum',
            'Saving 3Y Reserv': 'sum',
            'Final Saving': 'sum'
        }).reset_index()
        storage_summary = storage_summary.rename(columns={'MeterCategory':'Category'})

    except Exception as e:
        logging.error("Error while grouping 'storage' dataframe: " + str(e))

    try:
        logging.info("Grouping 'app_services' dataframe by 'MeterCategory'.")
        app_services_summary = app_services.groupby('MeterCategory').agg({
            'Saving PAYG': 'sum',
            'Saving 1Y Reserv': 'sum',
            'Saving 3Y Reserv': 'sum',
            'Final Saving': 'sum'
        }).reset_index()
        app_services_summary = app_services_summary.rename(columns={'MeterCategory':'Category'})
        
    except Exception as e:
        logging.error("Error while grouping 'app_services' dataframe: " + str(e))

    try:
        logging.info("Grouping 'backups' dataframe by 'MeterCategory'.")
        backup_summary = backups.groupby('MeterCategory').agg({
            'Saving PAYG': 'sum',
            'Saving 1Y Reserv': 'sum',
            'Saving 3Y Reserv': 'sum',
            'Final Saving': 'sum'
        }).reset_index()
        backup_summary = backup_summary.rename(columns={'MeterCategory':'Category'})

    except Exception as e:
        logging.error("Error while grouping 'backups' dataframe: " + str(e))

    try:
        logging.info("Initializing 'Saving 1Y Reserv' and 'Saving 3Y Reserv' columns in 'application_gateway' dataframe.")
        application_gateway['Saving 1Y Reserv'] = 0
        application_gateway['Saving 3Y Reserv'] = 0
    except Exception as e:
        logging.error("Error while initializing 'Saving 1Y Reserv' and 'Saving 3Y Reserv' columns in 'application_gateway' dataframe: " + str(e))

    try:
        logging.info("Grouping 'application_gateway' dataframe by 'MeterCategory'.")
        application_gateway_summary = application_gateway.groupby('MeterCategory').agg({
            'Saving PAYG': 'sum',
            'Saving 1Y Reserv': 'sum',
            'Saving 3Y Reserv': 'sum',
            'Final Saving': 'sum'
        }).reset_index()
        application_gateway_summary = application_gateway_summary.rename(columns={'MeterCategory':'Category'})

    except Exception as e:
        logging.error("Error while grouping 'application_gateway' dataframe: " + str(e))

    try:
        if sql_database is None:
            logging.info("No 'sql_database' dataframe provided.")
            sql_database_summary = None
        else:
            logging.info("Initializing 'Saving PAYG' column in 'sql_database' dataframe.")
            sql_database['Saving PAYG'] = 0
            logging.info("Grouping 'sql_database' dataframe by 'MeterCategory'.")
            sql_database_summary = sql_database.groupby('MeterCategory').agg({
                'Saving PAYG': 'sum',
                'Saving 1Y Reserv': 'sum',
                'Saving 3Y Reserv': 'sum',
                'Final Saving': 'sum'
            }).reset_index()
            sql_database_summary = sql_database_summary.rename(columns={'MeterCategory':'Category'})

    except Exception as e:
        logging.error("Error while grouping 'sql_database' dataframe: " + str(e))

    try:
        # Aplicando a função calculate_min_max a todos os summaries
        vm_summary = vm_summary.apply(calculate_min_max, axis=1)
        storage_summary = storage_summary.apply(calculate_min_max, axis=1)
        app_services_summary = app_services_summary.apply(calculate_min_max, axis=1)
        backup_summary = backup_summary.apply(calculate_min_max, axis=1)
        application_gateway_summary = application_gateway_summary.apply(calculate_min_max, axis=1)
        if sql_database_summary is not None:
            sql_database_summary = sql_database_summary.apply(calculate_min_max, axis=1)

    except Exception as e:
        logging.error("Error while calculating Min and Max values: " + str(e))

    try:
        logging.info("Concatenating summary dataframes to create final summary.")
        summaries = [vm_summary, storage_summary, app_services_summary, backup_summary, application_gateway_summary]
        if sql_database_summary is not None:
            summaries.append(sql_database_summary)
        final_summary = pd.concat(summaries)
        final_summary = final_summary.rename(columns={'Min':'Min Estimated Savings (€)', 'Max':'Max Estimated Savings (€)'})
    except Exception as e:
        logging.error("Error while concatenating summary dataframes: " + str(e))
        return None

    logging.info("Summary processing completed successfully")
    return final_summary

def summary_overview(vm: pd.DataFrame, storage: pd.DataFrame, app_services: pd.DataFrame, backups: pd.DataFrame, application_gateway: pd.DataFrame, sql_database: pd.DataFrame):
    try:
        logging.info("Calculating 'Optimized Cost' for 'vm' dataframe.")
        #vm['Optimized Monthly Cost (€)'] = vm['PreTaxCost'] - vm['Final Saving']
        
        logging.info("Applying min and max calculations to 'vm' dataframe.")
        vm = vm.apply(calculate_min_max, axis=1)
        
        logging.info("Grouping 'vm' dataframe by 'MeterCategory'.")
        vm_summary = vm.groupby('MeterCategory').agg({
                'PreTaxCost': 'sum',
                'Min': 'sum',
                'Max': 'sum'
        }).reset_index()
        vm_summary = vm_summary.rename(columns={'PreTaxCost': 'Current Monthly Cost (€)', 'MeterCategory':'Category'})
    except Exception as e:
        logging.error("Error while grouping 'vm' dataframe: " + str(e))

    try:
        logging.info("Calculating 'Optimized Cost' for 'storage' dataframe.")
        #storage['Optimized Monthly Cost (€)'] = storage['PreTaxCost'] - storage['Final Saving']
        
        logging.info("Applying min and max calculations to 'storage' dataframe.")
        storage = storage.apply(calculate_min_max, axis=1)

        logging.info("Grouping 'storage' dataframe by 'MeterCategory'.")
        storage_summary = storage.groupby('MeterCategory').agg({
                'PreTaxCost': 'sum',
                'Min': 'sum',
                'Max': 'sum'
        }).reset_index()
        storage_summary = storage_summary.rename(columns={'PreTaxCost': 'Current Monthly Cost (€)', 'MeterCategory':'Category'})        
    except Exception as e:
        logging.error("Error while grouping 'storage' dataframe: " + str(e))

    try:
        logging.info("Calculating 'Optimized Cost' for 'app_services' dataframe.")
        #app_services['Optimized Monthly Cost (€)'] = app_services['PreTaxCost'] - app_services['Final Saving']
        
        logging.info("Applying min and max calculations to 'app_services' dataframe.")
        app_services = app_services.apply(calculate_min_max, axis=1)

        logging.info("Grouping 'app_services' dataframe by 'MeterCategory'.")
        app_services_summary = app_services.groupby('MeterCategory').agg({
                'PreTaxCost': 'sum',
                'Min': 'sum',
                'Max': 'sum'
        }).reset_index()
        app_services_summary = app_services_summary.rename(columns={'PreTaxCost': 'Current Monthly Cost (€)', 'MeterCategory':'Category'})   
    except Exception as e:
        logging.error("Error while grouping 'app_services' dataframe: " + str(e))

    try:
        logging.info("Calculating 'Optimized Cost' for 'backups' dataframe.")
        #backups['Optimized Monthly Cost (€)'] = backups['PreTaxCost'] - backups['Final Saving']
        
        logging.info("Applying min and max calculations to 'backups' dataframe.")
        backups = backups.apply(calculate_min_max, axis=1)

        logging.info("Grouping 'backups' dataframe by 'MeterCategory'.")
        backup_summary = backups.groupby('MeterCategory').agg({
                'PreTaxCost': 'sum',
                'Min': 'sum',
                'Max': 'sum'
        }).reset_index()
        backup_summary = backup_summary.rename(columns={'PreTaxCost': 'Current Monthly Cost (€)', 'MeterCategory':'Category'})   
    except Exception as e:
        logging.error("Error while grouping 'backups' dataframe: " + str(e))

    try:
        logging.info("Calculating 'Optimized Cost' for 'application_gateway' dataframe.")
        if application_gateway is None or application_gateway.empty:
            logging.info("No 'application gateway' dataframe provided.")
            application_gateway_summary = pd.DataFrame(columns=['Category', 'Current Monthly Cost (€)', 'Min', 'Max'])
        else:
            logging.info("Applying min and max calculations to 'application_gateway' dataframe.")
            application_gateway = application_gateway.apply(calculate_min_max, axis=1)

            logging.info("Grouping 'application_gateway' dataframe by 'MeterCategory'.")
            application_gateway_summary = application_gateway.groupby('MeterCategory').agg({
                    'PreTaxCost': 'sum',
                    'Min': 'sum',
                    'Max': 'sum'
            }).reset_index()
            application_gateway_summary = application_gateway_summary.rename(columns={'PreTaxCost': 'Current Monthly Cost (€)', 'MeterCategory':'Category'})   
    except Exception as e:
        logging.error("Error while grouping 'application_gateway' dataframe: " + str(e))

    try:
        if sql_database is None or sql_database.empty:
            logging.info("No 'sql_database' dataframe provided.")
            sql_database_summary = pd.DataFrame(columns=['Category', 'Current Monthly Cost (€)', 'Min', 'Max'])
        else:
            logging.info("Calculating 'Optimized Cost' for 'sql_database' dataframe.")
            #sql_database['Optimized Monthly Cost (€)'] = sql_database['PreTaxCost'] - sql_database['Final Saving']
            
            logging.info("Applying min and max calculations to 'sql_database' dataframe.")
            sql_database = sql_database.apply(calculate_min_max, axis=1)

            logging.info("Grouping 'sql_database' dataframe by 'MeterCategory'.")
            sql_database_summary = sql_database.groupby('MeterCategory').agg({
                'PreTaxCost': 'sum',
                'Min': 'sum',
                'Max': 'sum'
            }).reset_index()
            sql_database_summary = sql_database_summary.rename(columns={'PreTaxCost': 'Current Monthly Cost (€)', 'MeterCategory':'Category'})
    except Exception as e:
        logging.error("Error while grouping 'sql_database' dataframe: " + str(e), exc_info=True)

    try:
        logging.info("Concatenating summary dataframes to create final summary.")
        summaries = [vm_summary, storage_summary, app_services_summary, backup_summary, application_gateway_summary]
        if not sql_database_summary.empty:
            summaries.append(sql_database_summary)
        final_summary = pd.concat(summaries)
        final_summary = final_summary.rename(columns={'Min':'Min Estimated Savings (€)', 'Max':'Max Estimated Savings (€)'})
        final_summary['Min Estimated Savings (%)'] = final_summary['Min Estimated Savings (€)'] / final_summary['Current Monthly Cost (€)']
        final_summary['Max Estimated Savings (%)'] = final_summary['Max Estimated Savings (€)'] / final_summary['Current Monthly Cost (€)']
        final_summary['Min Optimized Cost (€)'] = final_summary['Current Monthly Cost (€)'] - final_summary['Max Estimated Savings (€)']
        final_summary['Max Optimized Cost (€)'] = final_summary['Current Monthly Cost (€)'] - final_summary['Min Estimated Savings (€)']

    except Exception as e:
        logging.error("Error while concatenating summary dataframes: " + str(e))
        return None

    logging.info("Summary overview processing completed successfully")
    return final_summary

def summary_assessment(vm: pd.DataFrame = None, storage: pd.DataFrame = None, app_services: pd.DataFrame = None, backups: pd.DataFrame = None, application_gateway: pd.DataFrame = None, sql_database: pd.DataFrame = None):
    try:
        # Inicializando variáveis de custos e economias
        total_current_cost = 0
        total_savings = 0
        min_estimated_savings = 0
        max_estimated_savings = 0

        # Função auxiliar para calcular os valores de um dataframe, se ele não for None
        def calculate_costs_and_savings(df):
            nonlocal total_current_cost, total_savings, min_estimated_savings, max_estimated_savings
            
            if df is not None:
                ensure_columns_exist(df, ['PreTaxCost', 'Final Saving', 'Min', 'Max'])
                total_current_cost += df['PreTaxCost'].sum()
                total_savings += df['Final Saving'].sum()
                df = df.apply(calculate_min_max, axis=1)
                min_estimated_savings += df['Min'].sum()
                max_estimated_savings += df['Max'].sum()
            else:
                df = pd.DataFrame(columns=['PreTaxCost', 'Final Saving', 'Min', 'Max'])
            return df

        # Calculando os valores para cada DataFrame
        logging.info("Calculando custos e economias para 'vm'.")
        vm = calculate_costs_and_savings(vm)

        logging.info("Calculando custos e economias para 'storage'.")
        storage = calculate_costs_and_savings(storage)

        logging.info("Calculando custos e economias para 'app_services'.")
        app_services = calculate_costs_and_savings(app_services)

        logging.info("Calculando custos e economias para 'backups'.")
        backups = calculate_costs_and_savings(backups)

        logging.info("Calculando custos e economias para 'application_gateway'.")
        application_gateway = calculate_costs_and_savings(application_gateway)

        logging.info("Calculando custos e economias para 'sql_database'.")
        sql_database = calculate_costs_and_savings(sql_database)

        # Criando o DataFrame de resumo
        logging.info("Criando summary data.")
        summary_data = {
            'Category': ['Virtual Machines', 'Storage', 'Azure App Service', 'Backup', 'Application Gateway', 'SQL Database'],
            'Current Monthly Cost (€)': [total_current_cost, None, None, None, None, None],
            'Min Estimated Savings (€)': [min_estimated_savings, None, None, None, None, None],
            'Max Estimated Savings (€)': [max_estimated_savings, None, None, None, None, None]
        }

        summary_df = pd.DataFrame(summary_data)

    except Exception as e:
        logging.error("Erro ao processar o resumo: " + str(e))
        summary_df = pd.DataFrame()

    logging.info("Resumo do assessment processado com sucesso")
    return summary_df

def summary_total(dir: Path, vm: pd.DataFrame = None, storage: pd.DataFrame = None, app_services: pd.DataFrame = None, backups: pd.DataFrame = None, application_gateway: pd.DataFrame = None, sql_database: pd.DataFrame = None):
    try:
        df = data_source_normalization(dir)
        
        total_azure_cost = df['PreTaxCost'].sum()

        # Função auxiliar para aplicar calculate_min_max e calcular savings se o DataFrame não for None
        def process_dataframe(df):
            if df is not None:
                ensure_columns_exist(df, ['PreTaxCost', 'Final Saving', 'Min', 'Max'])
                df = df.apply(calculate_min_max, axis=1)
                return df['Final Saving'].sum(), df['Min'].sum(), df['Max'].sum()
            else:
                return 0, 0, 0

        # Calculando savings e economias mínimas e máximas para cada DataFrame
        logging.info("Calculando custos e economias para 'vm'.")
        vm_savings, vm_min_savings, vm_max_savings = process_dataframe(vm)

        logging.info("Calculando custos e economias para 'storage'.")
        storage_savings, storage_min_savings, storage_max_savings = process_dataframe(storage)

        logging.info("Calculando custos e economias para 'app_services'.")
        app_services_savings, app_services_min_savings, app_services_max_savings = process_dataframe(app_services)

        logging.info("Calculando custos e economias para 'backups'.")
        backups_savings, backups_min_savings, backups_max_savings = process_dataframe(backups)

        logging.info("Calculando custos e economias para 'application_gateway'.")
        application_gateway_savings, application_gateway_min_savings, application_gateway_max_savings = process_dataframe(application_gateway)

        logging.info("Calculando custos e economias para 'sql_database'.")
        sql_database_savings, sql_database_min_savings, sql_database_max_savings = process_dataframe(sql_database)

        # Calculando totais
        total_savings = (vm_savings + storage_savings + app_services_savings + backups_savings + application_gateway_savings + sql_database_savings)
        total_optimized_cost = total_azure_cost - total_savings

        min_estimated_savings = (vm_min_savings + storage_min_savings + app_services_min_savings + backups_min_savings + application_gateway_min_savings + sql_database_min_savings)
        max_estimated_savings = (vm_max_savings + storage_max_savings + app_services_max_savings + backups_max_savings + application_gateway_max_savings + sql_database_max_savings)

        summary_data = {
            'Current Monthly Cost (€)': [total_azure_cost],
            'Min Estimated Savings (€)': [min_estimated_savings],
            'Max Estimated Savings (€)': [max_estimated_savings],
            'Min Estimated Savings (%)': [min_estimated_savings / total_azure_cost],
            'Max Estimated Savings (%)': [max_estimated_savings / total_azure_cost],
            'Minimum Optimized Cost (€)': [total_azure_cost - max_estimated_savings],
            'Maximum Optimized Cost (€)': [total_azure_cost - min_estimated_savings]
        }

        summary_df = pd.DataFrame(summary_data)
        return summary_df

    except Exception as e:
        logging.error(f"Error in summary_total: {e}")
        return None

def save_and_open_excel(vm, storage, app_services, backup, application_gateway, sql_database, summary_overview, summary_assessment, summary_total):
    try:
        # Create a temporary file
        logging.info("Creating a temporary Excel file.")
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            temp_path = tmp.name

        with pd.ExcelWriter(temp_path, engine='openpyxl') as writer:
            logging.info("Writing dataframes to Excel sheets.")
            vm.to_excel(writer, sheet_name='Virtual Machines', index=False)
            storage.to_excel(writer, sheet_name='Storage', index=False)
            app_services.to_excel(writer, sheet_name='App Services', index=False)
            backup.to_excel(writer, sheet_name='Backups', index=False)
            application_gateway.to_excel(writer, sheet_name='Application Gateway', index=False)
            #summary.to_excel(writer, sheet_name='Optimizations by category', index=False)
            summary_total.to_excel(writer, sheet_name='Summary', index=False)
            if sql_database is not None:
                sql_database.to_excel(writer, sheet_name='SQL Database', index=False)
            if summary_overview is not None:
                summary_overview.to_excel(writer, sheet_name='Optimizations by category', index=False)

        wb = load_workbook(temp_path)

        # Apply formatting to Summary Assessment
        if summary_assessment is not None:
            logging.info("Adding 'Total optimizations' sheet.")
            ws_assessment = wb.create_sheet('Total optimizations')
            for r in dataframe_to_rows(summary_assessment, index=False, header=True):
                ws_assessment.append(r)
            format_headers(ws_assessment)
            adjust_column_widths(ws_assessment)
            remove_gridlines(ws_assessment)

            # Merge cells for Current Cost, Savings, and Optimized Cost
            logging.info("Merging cells in 'Total optimizations' sheet.")
            ws_assessment.merge_cells('B2:B7')
            ws_assessment.merge_cells('C2:C7')
            ws_assessment.merge_cells('D2:D7')

            # Centralize alignment in merged cells and format as currency
            alignment = Alignment(horizontal='center', vertical='center')
            euro_format = '#,##0€'
            percentage_format = '0.00%'

            for cell in ['B2', 'C2', 'D2', 'E2', 'F2', 'G2']:
                ws_assessment[cell].alignment = alignment
                if cell != 'F2':
                    ws_assessment[cell].number_format = euro_format
                else:
                    ws_assessment[cell].number_format = percentage_format

        # Apply formatting to Summary Overview
        if summary_overview is not None:
            logging.info("Adding 'Optimizations by category' sheet.")
            ws_overview = wb['Optimizations by category']
            format_headers(ws_overview)
            remove_gridlines(ws_overview)
            apply_number_format(ws_overview, '#,##0€')
            adjust_column_widths(ws_overview)

            # Centralize alignment em todas as células e formatar como moeda
            alignment = Alignment(horizontal='center', vertical='center')
            euro_format = '#,##0€'
            percentage_format = '0.00%'

            for row in ws_overview.iter_rows(min_row=2, min_col=2, max_col=8):
                for cell in row:
                    cell.alignment = alignment
                    if cell.column_letter not in ['E', 'F']:
                        cell.number_format = euro_format
                    else:
                        cell.number_format = percentage_format

        # Apply formatting to Summary
        #if summary is not None:
        #    logging.info("Adding 'Optimizations by category' sheet.")
        #    ws_summary = wb['Optimizations by category']
        #    format_headers(ws_summary)
        #    remove_gridlines(ws_summary)
        #    apply_number_format(ws_summary, '#,##0.00€')
        #    adjust_column_widths(ws_summary)

        # Apply formatting to Summary Total
        if summary_total is not None:
            logging.info("Adding 'Summary' sheet.")
            ws_summary_total = wb['Summary']
            format_headers(ws_summary_total)
            remove_gridlines(ws_summary_total)
            apply_number_format(ws_summary_total, '#,##0.00€')
            adjust_column_widths(ws_summary_total)

            # Centralize alignment in merged cells and format as currency
            alignment = Alignment(horizontal='center', vertical='center')
            euro_format = '#,##0€'
            percentage_format = '0.00%'

            for cell in ['B2', 'C2', 'D2', 'E2', 'F2', 'G2']:
                ws_summary_total[cell].alignment = alignment
                if cell != 'D2' and cell != 'E2':
                    ws_summary_total[cell].number_format = euro_format
                else:
                    ws_summary_total[cell].number_format = percentage_format

        # Reorder sheets
        logging.info("Reordering sheets.")
        sheet_order = ['Total optimizations', 'Optimizations by category', 'Summary', 'Virtual Machines', 'Storage', 'App Services', 'Backups', 'Application Gateway', 'SQL Database']
        sheets = {sheet.title: sheet for sheet in wb.worksheets}
        new_order_sheets = [sheets[name] for name in sheet_order if name in sheets]

        wb._sheets = new_order_sheets

        # Set the active sheet to "Summary Assessment"
        wb.active = wb['Total optimizations']
        
        wb.save(temp_path)

        logging.info(f"Data saved to temporary Excel file: {temp_path}")

        # Open the file with the default application
        logging.info("Opening the Excel file with the default application.")
        if os.name == 'nt':  # Windows
            os.startfile(temp_path)
        elif os.name == 'posix':  # macOS or Linux
            os.system(f'open "{temp_path}"' if sys.platform == 'darwin' else f'xdg-open "{temp_path}"')

        logging.info("Excel file opened successfully.")
    except Exception as e:
        logging.error(f"Error in save_and_open_excel: {str(e)}")

def select_file():
    try:
        logging.info("Opening file dialog to select a file.")
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename()
        if file_path:
            logging.info(f"File selected: {file_path}")
        else:
            logging.warning("No file selected.")
        return file_path
    except Exception as e:
        logging.error(f"Error in select_file: {str(e)}")
        return None

def process_file(file_path):
    if not file_path:
        print("No file selected.")
        logging.error("No file selected.")
        return None, None, None, None, None, None, None, None, None

    try:
        logging.info(f"Loading configuration from config.json")
        config = load_config("config.json")
        _, file_extension = os.path.splitext(file_path)
        logging.info(f"File extension detected: {file_extension}")

        logging.info("Processing virtual machines data.")
        virtual_machines_df = virtual_machines(file_path, config)
        
        logging.info("Processing storage data.")
        storage_df = storage(file_path, config)
        
        logging.info("Processing app services data.")
        app_services_df = app_services(file_path, config)
        
        logging.info("Processing backup data.")
        backup_df = backup(file_path, config)
        
        logging.info("Processing application gateway data.")
        application_gateway_df = application_gateway(file_path, config)
        
        if file_extension == '.xlsx':
            logging.info("Skipping SQL database processing for Excel file.")
            sql_database_df = None
        else:
            logging.info("Processing SQL database data.")
            sql_database_df = sql_database(file_path, config)
        
        logging.info("Creating summary.")
        summary_df = summary(virtual_machines_df, storage_df, app_services_df, backup_df, application_gateway_df, sql_database_df)
        
        logging.info("Creating summary overview.")
        summary_overview_df = summary_overview(virtual_machines_df, storage_df, app_services_df, backup_df, application_gateway_df, sql_database_df)
        
        logging.info("Creating summary assessment.")
        summary_assessment_df = summary_assessment(virtual_machines_df, storage_df, app_services_df, backup_df, application_gateway_df, sql_database_df)

        logging.info("Creating summary assessment.")
        summary_total_df = summary_total(file_path, virtual_machines_df, storage_df, app_services_df, backup_df, application_gateway_df, sql_database_df)

        logging.info("Saving and opening Excel file with all dataframes.")
        save_and_open_excel(virtual_machines_df, storage_df, app_services_df, backup_df, application_gateway_df, sql_database_df, summary_overview_df, summary_assessment_df, summary_total_df)

        logging.info("File processing completed successfully.")
        return virtual_machines_df, storage_df, app_services_df, backup_df, application_gateway_df, sql_database_df, summary_overview_df, summary_assessment_df, summary_total_df

    except Exception as e:
        logging.error(f"An error occurred during file processing: {e}")
        return None, None, None, None, None, None, None, None

# A função 'select_file' é chamada para selecionar o arquivo1
selected_file = select_file()
if selected_file:
    # A função 'process_file' é chamada para processar o arquivo selecionado, retornando DataFrames
    df, df2, df3, df4, df5, df6, df7, df8, df9 = process_file(selected_file)
    if df is not None and df2 is not None and df3 is not None and df4 is not None and df5 is not None and df6 is not None:
        print("Processing complete. Output file opened.")
    else:
        print("File processing failed.")
else:
    df, df2, df3, df4, df5, df6, df7, df8, df9 = None, None, None, None, None, None, None, None, None
    print("No file selected.")

