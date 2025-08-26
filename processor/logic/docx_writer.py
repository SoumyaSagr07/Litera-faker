import re
import random
import string
import docx
from faker import Faker
from docx.oxml.ns import qn

# Initialize Faker for generating random content
fake = Faker()

def categorize_token(token):
    """
    Categorize a token to determine how it should be processed.
    Returns: 'word', 'number', 'punctuation', or 'mixed'
    """
    if not token.strip():
        return 'empty'
    
    alpha_count = len([c for c in token if c.isalpha()])
    digit_count = len([c for c in token if c.isdigit()])
    other_count = len([c for c in token if not c.isalpha() and not c.isdigit() and not c.isspace()])
    
    if alpha_count > 0 and digit_count == 0 and other_count == 0:
        return 'word'
    elif digit_count > 0 and alpha_count == 0 and other_count == 0:
        return 'number'
    elif alpha_count == 0 and digit_count == 0 and other_count > 0:
        return 'punctuation'
    elif alpha_count > 0 and (digit_count > 0 or other_count > 0):
        return 'mixed'
    else:
        return 'other'

def generate_length_preserving_text(original_text, padding_char='y'):
    """
    Generates random text that preserves length and casing by truncating or
    padding a single fake word. This is a fast, no-map approach.
    """
    if not original_text or not original_text.strip():
        return original_text

    new_text_parts = []
    # This regex splits the text into words and the spaces/punctuation between them.
    tokens = re.split(r'(\s+)', original_text)

    for token in tokens:
        # Use the categorization function to better understand the token
        token_type = categorize_token(token)
        
        # Debug logging for mixed tokens (like company names)
        if token_type == 'mixed':
            print(f"Debug: Mixed token '{token}' - type: {token_type}, alpha_chars: {len([c for c in token if c.isalpha()])}, non_alpha_chars: {len([c for c in token if not c.isalpha()])}")
        
        # Process based on token type
        if token_type in ['word', 'mixed']:
            # Treat both pure words and mixed tokens (like AT&T) as words to anonymize
            length = len(token)
            fake_word = fake.word()
            
            if len(fake_word) > length:
                new_word = fake_word[:length]
            elif len(fake_word) < length:
                # Use a more intelligent padding strategy
                if length <= 3:
                    # For very short words, just use the fake word as-is
                    new_word = fake_word
                else:
                    # For longer words, try to create a more natural-looking word
                    # by repeating parts of the fake word instead of single characters
                    if len(fake_word) >= 2:
                        # Repeat parts of the fake word to reach desired length
                        repeat_part = fake_word[-2:] if len(fake_word) >= 2 else fake_word
                        while len(fake_word) < length:
                            fake_word += repeat_part
                        new_word = fake_word[:length]
                    else:
                        # Fallback to minimal padding
                        new_word = fake_word.ljust(length, padding_char)
            else:
                new_word = fake_word
            
            if token.isupper():
                new_text_parts.append(new_word.upper())
            elif token.istitle():
                new_text_parts.append(new_word.title())
            else:
                new_text_parts.append(new_word.lower())
        elif token_type == 'number':
            length = len(token)
            
            # To avoid leading zeros in multi-digit numbers, handle first digit separately
            if length > 1 and token[0] != '0':
                first_digit = str(random.randint(1, 9))
                rest_of_digits = ''.join(random.choices(string.digits, k=length - 1))
                new_number_str = first_digit + rest_of_digits
            else: # For single digits or numbers that start with '0'
                new_number_str = ''.join(random.choices(string.digits, k=length))
            
            new_text_parts.append(new_number_str)

        else:
            # It's punctuation, numbers, or whitespace; keep it as is.
            new_text_parts.append(token)
            
    return "".join(new_text_parts)


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
    The main anonymizer function. It anonymizes comments, tracked changes,
    body text, tables, headers, and footers.
    """
    try:
        document = docx.Document(file_path)

        # --- Comments ---
        def anonymize_all_comments(document):
            """
            Anonymize all comments in the DOCX file.
            """
            if hasattr(document.part, "comments") and document.part.comments:
                for comment in document.part.comments:
                    # Replace comment text
                    for para in comment.paragraphs:
                        if para.text.strip():
                            para.text = generate_length_preserving_text(para.text)

                    # Replace author details
                    comment.author = fake.name()
                    comment.initials = "".join(part[0] for part in comment.author.split() if part)

        anonymize_all_comments(document)

        # --- Authors mapping for tracked changes ---
        author_map = {}
        def get_fake_author(original_author):
            if original_author not in author_map:
                author_map[original_author] = fake.name()
            return author_map[original_author]

        # --- Recursive anonymizer for all content ---
        def process_element_runs(element):
            # Handle tracked changes (insertions/deletions)
            for tag_name in ['w:ins', 'w:del','w:moveFrom', 'w:moveTo']:
                for change_tag in element._element.findall('.//' + qn(tag_name)):
                    original_author = change_tag.get(qn('w:author'))
                    if original_author:
                        change_tag.set(qn('w:author'), get_fake_author(original_author))
                    if change_tag.get(qn('w:date')):
                        fake_date = fake.past_datetime(start_date="-30d")
                        change_tag.set(qn('w:date'), fake_date.strftime('%Y-%m-%dT%H:%M:%SZ'))

                    for t in change_tag.findall('.//' + qn('w:t')):
                     if t.text and t.text.strip():
                      t.text = generate_length_preserving_text(t.text)

            # deleted text often lives in w:delText (not w:t)
                    for dt in change_tag.findall('.//' + qn('w:delText')):
                     if dt.text and dt.text.strip():
                      dt.text = generate_length_preserving_text(dt.text)


            # Handle paragraphs and runs
            if hasattr(element, "paragraphs"):
                for para in element.paragraphs:
                    for run in para.runs:
                        if run.text.strip():
                            run.text = generate_length_preserving_text(run.text)

            # Handle tables recursively
            if hasattr(element, "tables"):
                for table in element.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            process_element_runs(cell)

        # --- Process the document body ---
        process_element_runs(document)

        # --- Process headers and footers ---
        for section in document.sections:
            process_element_runs(section.header)
            process_element_runs(section.footer)

        document.save(output_path)
        print(f"Document fully anonymized (truncate/pad method): {output_path}")

    except Exception as e:
        print(f"An error occurred while processing the DOCX file: {e}")
        raise
