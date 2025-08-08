# processor/logic/docx_writer.py

import zipfile
import xml.etree.ElementTree as ET
import shutil
import random
import string
from faker import Faker

# Initialize Faker for generating random content
fake = Faker()

def generate_random_text(original_text):
    """
    Generate random text to replace the original content while preserving structure
    """
    if not original_text or not original_text.strip():
        return original_text
    
    # Preserve whitespace and special characters
    if original_text.isspace():
        return original_text
    
    # Handle numbers - replace with random numbers of similar magnitude
    clean_text = original_text.strip()
    
    # Debug: Print what we're processing
    print(f"Processing text: '{original_text}'")
    
    # First, try to extract just the digits and decimal point
    import re
    
    digits_only = re.sub(r'[^\d\.]', '', clean_text)
    print(f"Digits only: '{digits_only}'")
    
    # Check if it's a valid number
    if digits_only and (digits_only.isdigit() or ('.' in digits_only and digits_only.replace('.', '').isdigit())):
        print(f"Detected as number: '{digits_only}'")
        
        try:
            random_str = ""  # Initialize random_str
            
            if '.' in digits_only:
                # It's a decimal number
                original_float = float(digits_only)
                if original_float.is_integer():
                    # Integer with decimal point
                    # Count digits in the original number
                    digit_count = len(digits_only.replace('.', ''))
                    # Generate random number with same digit count
                    min_val = 10 ** (digit_count - 1)
                    max_val = (10 ** digit_count) - 1
                    random_num = fake.random_int(min=min_val, max=max_val)
                    random_str = str(random_num)
                else:
                    # Real decimal
                    # Count digits before and after decimal
                    parts = digits_only.split('.')
                    before_decimal = len(parts[0])
                    after_decimal = len(parts[1])
                    
                    # Generate random number with same digit structure
                    min_val = 10 ** (before_decimal - 1) if before_decimal > 1 else 1
                    max_val = (10 ** before_decimal) - 1
                    random_whole = fake.random_int(min=min_val, max=max_val)
                    
                    # Generate decimal part with same number of digits
                    random_decimal = fake.random_int(min=0, max=(10 ** after_decimal) - 1)
                    random_str = f"{random_whole}.{random_decimal:0{after_decimal}d}"
            else:
                # It's an integer
                # Count digits in the original number
                digit_count = len(digits_only)
                # Generate random number with same digit count
                min_val = 10 ** (digit_count - 1)
                max_val = (10 ** digit_count) - 1
                random_num = fake.random_int(min=min_val, max=max_val)
                random_str = str(random_num)
            
            # Preserve the original formatting by replacing digits with random digits
            result = clean_text
            # Replace the digits part with our random number
            if '.' in digits_only:
                # For decimals, replace the whole number part
                result = re.sub(r'\d+\.\d+', random_str, clean_text)
            else:
                # For integers, replace the digits
                result = re.sub(r'\d+', random_str, clean_text)
            
            print(f"Replaced '{original_text}' with '{result}'")
            return result
            
        except ValueError as e:
            print(f"Error processing number: {e}")
            pass
    
    # Generate random text with similar characteristics for non-numbers
    words = original_text.split()
    if not words:
        return original_text
    
    # Generate random words with similar length
    random_words = []
    for word in words:
        # Preserve punctuation and special characters
        if word.isalpha():
            # Generate random word of similar length
            length = len(word)
            if length <= 3:
                random_word = fake.word()[:length]
            else:
                random_word = fake.word()
                if len(random_word) < length:
                    random_word += ''.join(random.choices(string.ascii_lowercase, k=length - len(random_word)))
                else:
                    random_word = random_word[:length]
        else:
            # Keep non-alphabetic content (punctuation) as is
            random_word = word
        
        random_words.append(random_word)
    
    return ' '.join(random_words)

def modify_docx_inplace(file_path, output_path):
    """
    Modify the original docx file in-place by parsing its XML and replacing all text content
    with random text while preserving document structure
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
            
            # Find all text elements and replace their content
            modified = False
            for text_elem in root.findall('.//w:t', ns):
                if text_elem.text and text_elem.text.strip():
                    original_text = text_elem.text
                    random_text = generate_random_text(original_text)
                    text_elem.text = random_text
                    modified = True
                  
            if modified:
                # Write the modified XML back to the zip
                modified_xml = ET.tostring(root, encoding='UTF-8', xml_declaration=True)
                docx_zip.writestr('word/document.xml', modified_xml)
                print(f"Document modified and saved to: {output_path}")
            else:
                print("No text content found to replace")
                
        except KeyError:
            print("Could not find document.xml in the docx file")
        except Exception as e:
            print(f"Error processing document: {e}")
