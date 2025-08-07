# processor/logic/docx_writer.py

import zipfile
import xml.etree.ElementTree as ET
import shutil
from anoymizer import anonymize_text

def modify_docx_inplace(file_path, output_path):
    """
    Modify the original docx file in-place by parsing its XML and replacing entities
    """
    # Create a copy of the original file
    shutil.copy2(file_path, output_path)
    
    # Open the docx file (it's a zip file)
    with zipfile.ZipFile(output_path, 'a') as docx_zip:
        # Get the document.xml file
        try:
            # Read the main document XML
            xml_content = docx_zip.read('word/document.xml')
            root = ET.fromstring(xml_content)
            
            # Define the namespace
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            # Find all text elements and anonymize them
            modified = False
            for text_elem in root.findall('.//w:t', ns):
                if text_elem.text and text_elem.text.strip():
                    original_text = text_elem.text
                    anonymized_text = anonymize_text(original_text)
                    if original_text != anonymized_text:
                        text_elem.text = anonymized_text
                        modified = True
                  
            if modified:
                # Write the modified XML back to the zip
                modified_xml = ET.tostring(root, encoding='UTF-8', xml_declaration=True)
                docx_zip.writestr('word/document.xml', modified_xml)
                print(f"Document modified and saved to: {output_path}")
            else:
                print("No entities found to anonymize")
                
        except KeyError:
            print("Could not find document.xml in the docx file")
        except Exception as e:
            print(f"Error processing document: {e}")
