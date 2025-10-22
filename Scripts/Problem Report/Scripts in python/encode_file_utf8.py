
import os
import xml.etree.ElementTree as ET

def save_xml_with_utf8_encoding(file_path):
    # Load the XML file
    tree = ET.parse(file_path)
    
    # Get the root element of the XML
    root = tree.getroot()
    
    # Save the XML file with UTF-8 encoding (overwrite the current file)
    tree.write(file_path, encoding="utf-8", xml_declaration=True)
    
    print(f"XML file saved with UTF-8 encoding: {file_path}")

# Provide the relative path of the XML file as the parameter
file_path = r".\EventTraces\Logging\LoggingExport.xml"

# Call the save_xml_with_utf8_encoding function
save_xml_with_utf8_encoding(file_path)
