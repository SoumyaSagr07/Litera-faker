# processor/logic/anonymizer.py

import re
from faker import Faker
from collections import defaultdict
import spacy

# Initialize Faker
fake = Faker()

# Map detected entities to fake replacements
entity_replacements = defaultdict(str)

# Map spaCy labels to fake generators
ENTITY_LABEL_TO_FAKE = {
    "PERSON": lambda: fake.name(),
    "ORG": lambda: fake.company(),
    "GPE": lambda: fake.city(),  # Geo-political entity (city, country)
    "DATE": lambda: fake.date(),
    "MONEY": lambda: f"${fake.random_int(min=100, max=10000)}",
    "CARDINAL": lambda: str(fake.random_int(min=1, max=1000)),
    "NORP": lambda: fake.country(),  # Nationality, religious or political groups
}

# Try to load spacy model, fallback to basic text processing if not available
try:
    nlp = spacy.load("en_core_web_sm")
    SPACY_AVAILABLE = True
except OSError:
    print("Warning: spaCy model 'en_core_web_sm' not found. Using basic text anonymization.")
    SPACY_AVAILABLE = False
    nlp = None  # type: ignore

def anonymize_text(text):
    """
    Anonymize text by replacing named entities with fake data
    """
    if not text or not text.strip():
        return text

    # If spacy is not available, use basic pattern matching
    if not SPACY_AVAILABLE:
        return anonymize_text_basic(text)
    
    try:
        doc = nlp(text)
        
        # Replace each entity only once (consistent replacement)
        new_text = text
        for ent in doc.ents:
            if ent.label_ in ENTITY_LABEL_TO_FAKE:
                original = ent.text
                if original not in entity_replacements:
                    entity_replacements[original] = ENTITY_LABEL_TO_FAKE[ent.label_]()
                # Use regex to replace exact match
                new_text = re.sub(rf'\b{re.escape(original)}\b', entity_replacements[original], new_text)
        
        return new_text
    except Exception as e:
        print(f"Error in spacy processing: {e}")
        return anonymize_text_basic(text)

def anonymize_text_basic(text):
    """
    Basic text anonymization using regex patterns when spacy is not available
    """
    # Common patterns for names, emails, phone numbers, etc.
    patterns = [
        # Email addresses
        (r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', lambda: fake.email()),
        # Phone numbers
        (r'\b\d{3}[-.]?\d{3}[-.]?\d{4}\b', lambda: fake.phone_number()),
        # Dates (MM/DD/YYYY or YYYY-MM-DD)
        (r'\b\d{1,2}/\d{1,2}/\d{4}\b', lambda: fake.date()),
        (r'\b\d{4}-\d{1,2}-\d{1,2}\b', lambda: fake.date()),
        # Credit card numbers
        (r'\b\d{4}[- ]?\d{4}[- ]?\d{4}[- ]?\d{4}\b', lambda: fake.credit_card_number()),
    ]
    
    new_text = text
    for pattern, replacement_func in patterns:
        matches = re.findall(pattern, new_text)
        for match in matches:
            if match not in entity_replacements:
                entity_replacements[match] = replacement_func()
            new_text = re.sub(re.escape(match), entity_replacements[match], new_text)
    
    return new_text

def anonymize_xml_tree(tree):
    """
    Anonymize XML tree elements
    """
    for elem in tree.iter():
        if elem.text:
            elem.text = anonymize_text(elem.text)
