import os
import re
import xml.etree.ElementTree as ET
import pandas as pd
import json

# Function to parse an XML file and return the root element
def parse_xml_file(file_path):
    try:
        tree = ET.parse(file_path)  # Parse the XML file
        root = tree.getroot()  # Get the root element
        return root
    except Exception as e:
        raise Exception(f"Failed to parse the XML file: {e}")

# Function to extract connection manager data from an XML root element
def get_cnx_mngr_data(root):
    try:
        # Extract connection manager ID and name
        cnx_mngr_id = root.attrib.get('{www.microsoft.com/SqlServer/Dts}DTSID', '')
        cnx_mngr_name = root.attrib.get('{www.microsoft.com/SqlServer/Dts}ObjectName', '')
        namespace = {"DTS": "www.microsoft.com/SqlServer/Dts"}

        # Find the ConnectionManager element in the XML
        cnx_mgr = root.find("DTS:ObjectData/DTS:ConnectionManager", namespace)

        if cnx_mgr is not None:
            cnx_str = cnx_mgr.attrib.get("{www.microsoft.com/SqlServer/Dts}ConnectionString", '')
        else:
            # If the ConnectionManager element is not found, extract connection string from other elements
            cnx_str = root.find(".//*[@ConnectionString]", namespace).attrib.get("ConnectionString", '')

        # Return cleaned connection manager ID and a tuple of name and connection string
        return cnx_mngr_id.strip('{}'), (cnx_mngr_name, cnx_str)
    except Exception as e:
        raise Exception(f"Failed to extract connection manager data: {e}")

# Function to get connection managers data from a folder of XML files
def get_cnx_mngrs_data(ssis_package_path):
    try:
        cnx_mngrs_dict = {}

        # Iterate over files in the specified directory
        for filename in os.listdir(ssis_package_path):
            if filename.endswith('.conmgr'):  # Check if the file is a connection manager file
                file_path = os.path.join(ssis_package_path, filename)  # Create the full path to the file
                root = parse_xml_file(file_path)  # Parse the XML file
                cnx_mngr_data = get_cnx_mngr_data(root)  # Extract connection manager data
                cnx_mngrs_dict[cnx_mngr_data[0]] = cnx_mngr_data[1]  # Store the data in the dictionary with the cleaned ID as the key

        # Return the dictionary of connection manager data
        return cnx_mngrs_dict
    except Exception as e:
        raise Exception(f"Failed to get connection managers data: {e}")

# Function to extract SSIS package data from an XML root element
def get_ssis_package_data(root, cnx_mngrs_dict):
    try:
        cnx_mngrs = []
        tables_lst = []
        procedures_lst = []
        tables_pattern = r'\[dbo\]\.\[(\w+)\]'  # Regular expression pattern to extract table names
        procedures_pattern = r'exec (\S+)'  # Regular expression pattern to extract procedure names

        # Iterate over elements in the XML
        for element in root.iter():
            if 'connectionManagerID' in element.attrib or 'Connection' in element.attrib:
                cnx_mngrs.append(cnx_mngrs_dict.get(element.get('connectionManagerID' or 'Connection').split(":")[0].strip('{}')))
            if element.text is not None and '[dbo]' in element.text:
                tables_lst.extend(re.findall(tables_pattern, element.text))
            if element.text is not None and 'exec ' in element.text:
                procedures_lst.extend(re.findall(procedures_pattern, element.text))
            for attr in element.attrib:
                if '[dbo]' in element.attrib[attr]:
                    tables_lst.extend(re.findall(tables_pattern, element.attrib[attr]))
                if 'exec ' in element.attrib[attr]:
                    procedures_lst.extend(re.findall(procedures_pattern, element.attrib[attr]))
                if 'Connection' in attr:
                    cnx_mngrs.append(cnx_mngrs_dict.get(element.attrib[attr].strip('{}')))
                if 'ConnectionString' in attr:
                    cnx_mngrs.append(element.attrib[attr])

        # Remove duplicates from the lists of related connection managers, tables, and procedures
        cnx_mngrs_lst = list(set(cnx_mngrs))
        tables_flat_lst = list(set(tables_lst))
        procedures_flat_lst = list(set(procedures_lst))

        # Return lists of related connection managers, tables, and procedures
        return (cnx_mngrs_lst, tables_flat_lst, procedures_flat_lst)
    except Exception as e:
        raise Exception(f"Failed to extract SSIS package data: {e}")

# Function to get SSIS package data from a folder of XML files
def get_ssis_packages_data(ssis_folder_path, cnx_mngrs_dict):
    try:
        ssis_packages_dict = {}

        # Iterate over files in the specified directory
        for filename in os.listdir(ssis_folder_path):
            if filename.endswith('.dtsx'):  # Check if the file is an SSIS package file
                file_path = os.path.join(ssis_folder_path, filename)  # Create the full path to the file
                root = parse_xml_file(file_path)  # Parse the XML file
                ssis_packages_dict[filename] = get_ssis_package_data(root, cnx_mngrs_dict)  # Extract and store SSIS package data in the dictionary

        # Return the dictionary of SSIS package data
        return ssis_packages_dict
    except Exception as e:
        raise Exception(f"Failed to get SSIS packages data: {e}")

# Function to save data to a JSON file
def save_to_json_file(my_dict, result_file_path):
    try:
        json_object = json.dumps(my_dict, indent=4)
        with open(result_file_path, "w") as outfile:
            outfile.write(json_object)
        return True
    except Exception as e:
        raise Exception(f"Failed to save data to JSON file: {e}")

# Function to save SSIS package data to an Excel file
def save_to_excel(ssis_packages_dict, result_file_path):
    try:
        df = pd.DataFrame.from_dict(ssis_packages_dict, orient='index',
                                 columns=['related connections',
                                           'related tables', 'related procedures'])
        df.reset_index(inplace=True)
        df.rename(columns={'index': 'package name'}, inplace=True)
        df.to_excel(result_file_path, index=False)
    except Exception as e:
        raise Exception(f"Failed to save data to Excel file: {e}")

# Main function that orchestrates the program
def main():
    try:
        # Prompt the user for input
        ssis_package_path = input("Enter the path to the SSIS package directory: ")
        result_json_path = input("Enter the path to save the JSON file: ")
        result_excel_path = input("Enter the path to save the Excel file: ")

        cnx_mngrs_dict = get_cnx_mngrs_data(ssis_package_path)  # Get connection manager data
        save_to_json_file(cnx_mngrs_dict, result_json_path)  # Save connection manager data to a JSON file

        ssis_packages_dict = get_ssis_packages_data(ssis_package_path, cnx_mngrs_dict)  # Get SSIS package data
        save_to_excel(ssis_packages_dict, result_excel_path)  # Save SSIS package data to an Excel file

        print("Data extraction and saving completed successfully.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
