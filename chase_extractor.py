import PyPDF2
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
import pdfplumber

def extract_text_from_pdf(pdf_path):
    """
    Extracts text from a PDF file using pdfplumber.
    Args:
        pdf_path (str): The path to the PDF file.
    Returns:
        str: The extracted text from the PDF.
    """
    text = ""
    print(f"Attempting to extract text from '{pdf_path}' using pdfplumber...")
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Debug: Check if PDF is openable and number of pages
            print(f"  PDF opened successfully. Number of pages: {len(pdf.pages)}")
            for page_num, page in enumerate(pdf.pages):
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n" # Add newline to separate pages
                    # Debug: Print a snippet of text from each page to see if it's working
                    # print(f"    Page {page_num + 1} extracted text (first 200 chars):\n    {page_text[:200]}...")
                else:
                    print(f"  Warning: Page {page_num + 1} of '{pdf_path}' yielded no text with pdfplumber. It might be a complex layout or scanned.")
        
        if not text.strip(): # If no text was extracted at all from any page
            print(f"  CRITICAL WARNING: No text extracted from '{pdf_path}' using pdfplumber. This PDF might be problematic.")
    
    except pdfplumber.pdf.PDFError as e: # Catch pdfplumber's specific PDF errors
        print(f"  Error reading PDF '{pdf_path}' (pdfplumber PDFError): {e}. This might be a corrupted or non-standard PDF internal structure.")
        return ""
    except Exception as e:
        print(f"  An unexpected error occurred while extracting text from '{pdf_path}': {e}")
        return ""
    
    print(f"  Finished text extraction. Total characters extracted: {len(text)}")
    # Debug: Print the full extracted text for one PDF for manual inspection
    # if len(text) > 0 and len(text) < 5000: # Don't print huge texts
    #     print("\n--- FULL EXTRACTED TEXT FOR DEBUGGING ---\n", text, "\n------------------------------------------\n")
    return text

def normalize_text(text):
    """
    Normalizes text by compressing repeated characters (e.g., AAA to A, AA to A).
    Useful for PDFs where text extraction doubles or triples characters.
    """
    if not text:
        return ""
    
    # New regex: replace 2 or more occurrences of a character with a single occurrence.
    # This will turn AA -> A, AAA -> A, AAAA -> A, etc.
    normalized = re.sub(r'(.)\1+', r'\1', text) 
    
    # Replace non-standard spaces and form feeds
    normalized = normalized.replace('\xa0', ' ').replace('\x0c', '').strip()
    return normalized

def parse_chase_statement(statement_text):
    transactions = []
    
    #normalized_statement_text = normalize_text(statement_text)
    
    #print("\n--- NORMALIZED EXTRACTED TEXT FOR DEBUGGING ---")
    #print(normalized_statement_text)
    #print("-----------------------------------------------\n")
    
    lines = statement_text.split('\n')
    in_transactions_section = False
    
    for i, line in enumerate(lines):
        line_strip = line.strip()
        # print(f"PARSING LINE {i}: '{line_strip}'") # Uncomment for extreme debugging
        
        # Use 'in' for marker check, it's more flexible than exact match or startswith
        if "Merchant Name" in line_strip:
            in_transactions_section = True
            print(f"PARSER DEBUG: Found start marker '' at line {i}.")
            continue 

        # --- IMPORTANT: Re-check end markers and be flexible ---
        if in_transactions_section:
            if ("FEES CHARGED" in line_strip or             #
                #"INTEREST CHARGED" in line_strip or         #
                "TOTAL FEES FOR THIS PERIOD" in line_strip or #
                "TOTAL INTEREST FOR THIS PERIOD" in line_strip or #
                "Totals Year-to-Date" in line_strip or       #
                "Total Balance" in line_strip or              # General Chase marker
                "Previous Balance" in line_strip or           # General Chase marker
                "Account Summary" in line_strip):             # General Chase marker
                
                print(f"PARSER DEBUG: Found end marker at line {i}. Stopping transaction parsing.")
                in_transactions_section = False
                break 
        
        if in_transactions_section:
            # Try re.search instead of re.match if lines might have leading junk
            # Original: r'(\d{2}\/\d{2})\s+(.+?)\s+(-?\d{1,3}(?:,\d{3})*\.\d{2})\s*$'
            transaction_match = re.search(
                r'(\d{2}\/\d{2})\s+' # Date (MM/DD)
                r'(.+?)\s+'        # Description (non-greedy, at least one char)
                r'(-?\d{1,3}(?:,\d{3})*\.\d{2})\s*$' # Amount (optional negative, commas, two decimals) to end of line
                , line_strip)
            
            if transaction_match:
                date = transaction_match.group(1)
                description = transaction_match.group(2).strip()
                amount = float(transaction_match.group(3).replace(',', ''))
                
                transactions.append({
                    "Date": date,
                    "Description": description,
                    "Amount": amount
                })
                print(f"PARSER DEBUG: MATCHED transaction at line {i}: {date}, {description}, {amount}")
            # Handling "EURO" lines or follow-up lines
            elif "EURO" in line_strip and transactions and re.search(r'^\s*\d{1,3}(?:,\d{3})*\.\d{2}\s+X\s+\d+\.\d+\s+\(EXCHG RATE\)', line_strip):
                # This specifically targets lines like "281.51 X 1.044794145 (EXCHG RATE)"
                transactions[-1]["Description"] += " " + line_strip
                print(f"PARSER DEBUG: Appended EXCHG RATE to prev transaction at line {i}.")
            elif "EURO" in line_strip and transactions and re.search(r'^\d{3}\.\d{2}\s+EURO\s*$', line_strip):
                # This catches "281.51 EURO" style lines
                transactions[-1]["Description"] += " " + line_strip
                print(f"PARSER DEBUG: Appended simple EURO to prev transaction at line {i}.")
            elif line_strip in ("PURCHASE", "PAYMENTS AND OTHER CREDITS"): #
                print(f"PARSER DEBUG: Skipping known section header: '{line_strip}' at line {i}.")
            else:
                print(f"PARSER DEBUG: NO MATCH (in transaction section) at line {i}: '{line_strip}'")
    
    print(f"PARSER DEBUG: Finished. Extracted {len(transactions)} transactions.")
    return transactions

def categorize_transaction(description, amount):
    """
    Categorizes a transaction based on its description and amount.
    This is a basic categorization and needs to be expanded and refined.
    """
    description = description.lower()
    
    # Payments and Credits
    if "payment thank you" in description or amount < 0: # Payments often have negative amounts
        return "Payment/Credit"
    
    # Fees Charged
    if "annual membership fee" in description or "plan fee" in description:
        return "Fees"
    
    # Interest Charged
    if "purchase interest charge" in description:
        return "Interest Charged"

    # Specific merchants from the sample
    if "amazon" in description:
        return "Online Shopping - Amazon"
    if "helium" in description:
        return "Phone services"
    if "google" in description:
        return "Digital Services / Google"
    if "playstation" in description:
        return "Digital Services/ Sony"
    if "sun fresh produce" in description or "js produce" in description or "trader joe" in description or "fresh market" in description or "wegmans" in description or "haris teeter" in description or "aldi" in description or "global store" in description or "giant" in description or "farm" in description or "lidl" in description:
        return "Groceries"
    if "sq *smart energy pros" in description:
        return "Utilities" # Placeholder, assuming it's a utility provider
    if "lansing bp lanh" in description:
        return "Home improvement" # Placeholder, assuming it's gas station
    if "prime videos" in description:
        return "Entertainment / Subscriptions"
    if "hetzner online" in description or "contabo" in description or "travchis" in description:
        return "Crypto/Servers" # Based on the name
    if "telegram" in description:
        return "Communication / App" # Based on the name
    if "ikea" in description in description or "lowes" in description:
        return "Home inventory"
    if "filling station" in description or "black eyed susan" in description:
        return "Coffee"
    if "american eagle" in description or "tjmax" in description or "ross store" in description or "j crew" in description:
        return "Clothing"
    if "californiapizzakithen" in description or "gongcha" in description or "glyndongrill" in description or "royal" in description:
        return "Dining Out"
    if "costco" in description:
        return "essentials"
    



    # General keywords (expand extensively for real use)
    if "restaurant" in description or "cafe" in description or "starbucks" in description or "pizza" in description or "taco" in description or "panera" in description or "grill" in description or "7-eleven" in description:
        return "Dining Out"
    if "wines" in description or "beer" in description or "liquor" in description or "spirit" in description:
        return "Alcohol"
    if "uber" in description or "lyft" in description or "taxi" in description:
        return "Transportation"
    if "oil" in description or "sunoco" in description or "gas" in description or "exxon" in description:
        return "Gas"
    if "auto parts" in description:
        return "Car maintenance"
    if "ezmd" in description or "mva" in description:
        return "Car fees"
    if "walmart" in description or "target" in description or "marshalls" in description:
        return "General Merchandise"
    if "pharmacy" in description or "cvs" in description or "walgreens" in description:
        return "Health"
    if "home" in description:
        return "Home inventory"
    if "spothero" in description:
        return "Parking"
    if "plan fee" in description:
        return "Credit card fees"
    if "linkedin" in description or "codecademy" in description or "discord" in description:
        return "subscriptions"
    
    return "Miscellaneous" # Default category

def main(pdf_folder_path, output_excel_path):
    """
    Main function to orchestrate PDF parsing, data categorization, and Excel export
    for multiple PDF files in a folder.
    """
    all_transactions_data = []

    if not os.path.isdir(pdf_folder_path):
        print(f"Error: Folder '{pdf_folder_path}' not found.")
        return

    for filename in os.listdir(pdf_folder_path):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder_path, filename)
            print(f"Processing {filename}...")

            # Extract last 4 digits of credit card number from filename
            # Example: "20250504-statements-1335-.pdf" or "20250504-statements-1335.pdf"
            card_match = re.search(r'-(\d{4})(?:-|\.pdf$)', filename)
            credit_card_last_4 = "N/A"
            if card_match:
                credit_card_last_4 = card_match.group(1)
            else:
                print(f"Warning: Could not extract last 4 digits from filename: {filename}")

            statement_text = extract_text_from_pdf(pdf_path)
            transactions_from_pdf = parse_chase_statement(statement_text)

            if transactions_from_pdf:
                # Add credit card number to each transaction
                for transaction in transactions_from_pdf:
                    transaction['Credit Card Last 4 Digits'] = credit_card_last_4
                all_transactions_data.extend(transactions_from_pdf)
            else:
                print(f"No transactions found in {filename}.")

    if not all_transactions_data:
        print("No transactions were extracted from any PDF. Exiting.")
        return

    df = pd.DataFrame(all_transactions_data)
    
    print("Categorizing all transactions...")
    df['Category'] = df.apply(lambda row: categorize_transaction(row['Description'], row['Amount']), axis=1)
    
    print(f"Saving categorized data to {output_excel_path}...")
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Chase Statement Analysis"
    
    # Write the DataFrame to the worksheet
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)
            
    # Auto-adjust column widths
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter 
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

    wb.save(output_excel_path)
    print("Analysis complete!")
    print("\nFirst 10 rows of combined and processed data:")
    print(df.head(10))

if __name__ == "__main__":
    # --- IMPORTANT: Setup for testing ---
    # 1. Create a folder named 'chase_statements' in the same directory as your script.
    # 2. Inside 'chase_statements', place your Chase PDF statement files.
    #    Rename them to include the last 4 digits of the card number, e.g.:
    #    "202501-statements-1234-.pdf"
    #    "202502-statements-5678.pdf"
    #    (The script uses a regex to find -XXXX- or -XXXX.pdf)

    # Set the path to your folder containing Chase PDF statements
    pdf_statements_folder =  "chase_statements"
    output_excel_file = "combined_chase_spending.xlsx"

    print("\n--- Starting Chase Statement Analysis ---")
    main(pdf_statements_folder, output_excel_file)
    print("--- Analysis Finished ---")