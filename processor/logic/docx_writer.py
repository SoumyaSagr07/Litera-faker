import docx
import re
from faker import Faker
from docx.oxml.ns import qn

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
    Anonymizes a .docx file, including comments, by inspecting the underlying XML
    to find comment references within text runs.
    """
    try:
        document = docx.Document(file_path)

        # --- Step 1: Create a dictionary mapping comment IDs to comment objects ---
        comments_map = {}
        # The correct way to access the comments part
        if hasattr(document.part, 'comments') and document.part.comments:
            for comment in document.part.comments:
                comments_map[comment.comment_id] = comment

        # If there are no comments, no need to proceed with comment logic
        if not comments_map:
            print("No comments found in the document.")
            # Fallback to simple anonymization without comment handling if needed
            # For now, we'll just proceed as the main logic handles text anyway
            pass

        processed_comment_ids = set()

        def anonymize_comment(comment_obj):
            """Anonymizes the text and metadata of a single comment object."""
            # Replace all paragraphs in the comment with a single fake sentence
            for para in comment_obj.paragraphs:
                # Clear existing content
                for run in para.runs:
                    run.text = ''
                # Add new fake content
                if para.runs:
                    para.runs[0].text = fake.sentence()
                else:
                    para.add_run(fake.sentence())
            
            # Anonymize comment metadata
            comment_obj.author = fake.name()
            comment_obj.initials = "".join(part[0] for part in comment_obj.author.split() if part)

        def process_element_runs(element):
            """Processes all runs in an element, faking text and finding comment links."""
            for para in element.paragraphs:
                for run in para.runs:
                    # Anonymize the run's text
                    if run.text.strip():
                        run.text = generate_random_text(run.text)

                    # --- Step 2 & 3: Inspect run's XML for comment reference ---
                    # The qn() function gets the qualified tag name
                    comment_ref_tag = qn('w:commentReference')
                    # Find all comment reference tags in the run's XML element (_r)
                    comment_references = run._r.findall(comment_ref_tag)
                    
                    for comment_ref in comment_references:
                        # --- Step 4: Extract comment ID ---
                        comment_id = int(comment_ref.get(qn('w:id')))
                        
                        # --- Step 5 & 6: Link to comment, process if not already done ---
                        if comment_id in comments_map and comment_id not in processed_comment_ids:
                            comment_to_anonymize = comments_map[comment_id]
                            anonymize_comment(comment_to_anonymize)
                            processed_comment_ids.add(comment_id)
            
            # Recursively process tables
            for table in element.tables:
                for row in table.rows:
                    for cell in row.cells:
                        process_element_runs(cell)

        # --- Process the entire document ---
        process_element_runs(document)
        for section in document.sections:
            process_element_runs(section.header)
            process_element_runs(section.footer)

        document.save(output_path)
        print(f"Document, including comments, successfully anonymized to: {output_path}")

    except Exception as e:
        print(f"An error occurred while processing the DOCX file: {e}")
        # Re-raising is good for debugging in a server environment
        raise