import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# File paths
#existing_excel_file = "نور الإيضاح في الفقه علي مذهب الإمام أبي حنيفة النعمان.xlsx"  # Existing Excel file
new_input_file = "الوحي والقرآن.txt"  # New text file
output_excel_file = "الوحي والقرآن.xlsx"  # Output will overwrite the existing Excel file

# Function to read and split text into paragraphs, ignoring PAGE_SEPARATOR
def read_and_split_paragraphs(file_path):
    try:
        # Read the file with UTF-8 encoding to handle Arabic text
        with open(file_path, 'r', encoding='utf-8') as file:
            text = file.read()
        
        # Remove PAGE_SEPARATOR to avoid splitting on it
        text = text.replace("PAGE_SEPARATOR", "")
        
        # Split text into paragraphs based on multiple newlines
        paragraphs = re.split(r'\n{2,}', text)
        
        # Clean up paragraphs: remove empty strings and strip whitespace
        paragraphs = [para.strip() for para in paragraphs if para.strip()]
        
        return paragraphs
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return []

# Function to load existing Excel file (if it exists)
def load_existing_excel(excel_file):
    try:
        df = pd.read_excel(excel_file, engine='openpyxl')
        return df['Paragraph'].tolist() if 'Paragraph' in df.columns else []
    except FileNotFoundError:
        print(f"Excel file {excel_file} not found. Starting with an empty list.")
        return []
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

# Read existing paragraphs from Excel (if any)
#existing_paragraphs = load_existing_excel(existing_excel_file)

# Read new paragraphs from the new text file
new_paragraphs = read_and_split_paragraphs(new_input_file)

# Combine paragraphs (existing + new)
all_paragraphs = new_paragraphs #existing_paragraphs + 

if all_paragraphs:
    # Create a DataFrame with all paragraphs in the first column
    df = pd.DataFrame(all_paragraphs, columns=["Paragraph"])
    
    # Write to Excel file
    try:
        df.to_excel(output_excel_file, index=False, engine='openpyxl')
        print(f"Excel file '{output_excel_file}' updated successfully with {len(all_paragraphs)} paragraphs.")
        
        # Apply right-to-left text alignment for Arabic text
        workbook = load_workbook(output_excel_file)
        worksheet = workbook.active
        for row in worksheet['A']:
            row.alignment = Alignment(horizontal='right')
        workbook.save(output_excel_file)
        print("Right-to-left alignment applied to Excel file.")
    except Exception as e:
        print(f"Error writing to Excel: {e}")
else:
    print("No paragraphs found or error occurred during processing.")