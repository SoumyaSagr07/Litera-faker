import docx
import re
from faker import Faker

# Initialize Faker for generating random content
fake = Faker()

# Your generate_random_text function does not need any changes.
# It's working correctly.
def generate_random_text(original_text):
    """
    Generate random text to replace the original content while preserving structure.
    (This function is the same as your original and is correct)
    """
    if not original_text or not original_text.strip():
        return original_text
    
    if original_text.isspace():
        return original_text
    
    clean_text = original_text.strip()
    digits_only = re.sub(r'[^\d\.]', '', clean_text)
    
    if digits_only and (digits_only.isdigit() or ('.' in digits_only and digits_only.replace('.', '', 1).isdigit())):
        try:
            if '.' in digits_only:
                parts = digits_only.split('.')
                before_decimal = len(parts[0])
                after_decimal = len(parts[1])
                min_val_whole = 10**(before_decimal - 1) if before_decimal > 0 else 0
                max_val_whole = (10**before_decimal) - 1
                random_whole = fake.random_int(min=min_val_whole, max=max_val_whole)
                random_decimal = fake.random_int(min=0, max=(10**after_decimal) - 1)
                random_str = f"{random_whole}.{random_decimal:0{after_decimal}d}"
            else:
                digit_count = len(digits_only)
                min_val = 10**(digit_count - 1) if digit_count > 0 else 0
                max_val = (10**digit_count) - 1
                random_num = fake.random_int(min=min_val, max=max_val)
                random_str = str(random_num)
            
            # Use a more careful replacement to keep surrounding text
            return original_text.replace(clean_text, random_str)
        except ValueError:
            pass
    
    words = original_text.split()
    if not words:
        return original_text
    
    random_words = [fake.word() for _ in words]
    return ' '.join(random_words)


def modify_docx_final(file_path, output_path):
    """
    A robust and correct function to anonymize a .docx file using python-docx.
    This version correctly replaces text in runs and also processes headers,
    footers, and tables.
    """
    try:
        # Open the source document. We will save it to the new output_path later.
        document = docx.Document(file_path)

        # --- Anonymize content in the main body (paragraphs and tables) ---
        for para in document.paragraphs:
            for run in para.runs:
                if run.text.strip():
                    run.text = generate_random_text(run.text)

        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if run.text.strip():
                                run.text = generate_random_text(run.text)

        # --- Anonymize content in headers and footers ---
        for section in document.sections:
            # Anonymize Header
            for para in section.header.paragraphs:
                for run in para.runs:
                    if run.text.strip():
                        run.text = generate_random_text(run.text)
            for table in section.header.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            for run in para.runs:
                                if run.text.strip():
                                    run.text = generate_random_text(run.text)
            
            # Anonymize Footer
            for para in section.footer.paragraphs:
                for run in para.runs:
                    if run.text.strip():
                        run.text = generate_random_text(run.text)
            for table in section.footer.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            for run in para.runs:
                                if run.text.strip():
                                    run.text = generate_random_text(run.text)

        # Save the fully modified document to the new output path
        document.save(output_path)
        print(f"Document successfully and correctly anonymized to: {output_path}")

    except Exception as e:
        print(f"An error occurred while processing with python-docx: {e}")
        # Re-raising the exception can help with debugging in Django
        raise