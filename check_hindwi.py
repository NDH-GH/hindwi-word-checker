import os
import sys
import time
import urllib.parse
import re
import html
import csv
import openpyxl
import requests

# Fix console encoding on Windows
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

INPUT_EXCEL = "Validated word list.xlsx"
PROGRESS_CSV = "hindwi_validation_results.csv"
OUTPUT_EXCEL = "Validated_word_list_with_hindwi.xlsx"
SHEET_NAME = "Validated_words"
SAVE_INTERVAL = 25  # Save to disk every 25 words
REQUEST_DELAY = 0.35  # Courtesy delay between requests (seconds)

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9,hi;q=0.8',
}


def load_checkpoint():
    """Load already processed words from CSV to avoid redundant requests."""
    processed = {}
    if os.path.exists(PROGRESS_CSV):
        try:
            with open(PROGRESS_CSV, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    word = row.get('Word', '').strip()
                    if word:
                        processed[word] = row
            print(f"[CHECKPOINT] Loaded {len(processed)} previously checked words from '{PROGRESS_CSV}'.")
        except Exception as e:
            print(f"[CHECKPOINT WARNING] Error reading {PROGRESS_CSV}: {e}")
    return processed


def append_results_to_csv(new_rows):
    """Appends batch of results to the progress CSV."""
    file_exists = os.path.exists(PROGRESS_CSV) and os.path.getsize(PROGRESS_CSV) > 0
    fieldnames = ['Index', 'Word', 'Hindwi_Status', 'Card_Word', 'Meaning', 'Meaning_URL']
    with open(PROGRESS_CSV, 'a', encoding='utf-8', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if not file_exists:
            writer.writeheader()
        for r in new_rows:
            writer.writerow(r)


def query_hindwi(session, word, retries=3):
    """Queries hindwidictionary.com and extracts the first card word & meaning."""
    encoded_word = urllib.parse.quote(word)
    url = f"https://www.hindwidictionary.com/search?keyword={encoded_word}"
    
    for attempt in range(1, retries + 1):
        try:
            resp = session.get(url, timeout=15)
            if resp.status_code == 200:
                resp_html = resp.text
                
                # Check for first card
                card_match = re.search(
                    r'<div class=["\']WordRelatedLists saveWord listingCardBg overflowAuto["\']>(.*?)</div>',
                    resp_html,
                    re.DOTALL
                )
                if card_match:
                    card_html = card_match.group(1)
                    li_match = re.search(
                        r'<li>\s*([^<\n\r]+?)\s*<p class=["\']nobullet["\']>(.*?)</p>',
                        card_html,
                        re.DOTALL
                    )
                    if li_match:
                        card_word = li_match.group(1).strip()
                        raw_meaning = li_match.group(2).strip()
                        meaning = html.unescape(re.sub(r'\s+', ' ', raw_meaning))
                        
                        # Extract data-url
                        url_match = re.search(r'data-url=["\']([^"\']*)["\']', card_html)
                        meaning_url = ""
                        if url_match:
                            slug = url_match.group(1).strip()
                            meaning_url = f"https://www.hindwidictionary.com{slug}" if slug.startswith('/') else slug
                        
                        status = "OKAY" if word.strip() == card_word.strip() else "MISMATCH"
                        return {
                            'Hindwi_Status': status,
                            'Card_Word': card_word,
                            'Meaning': meaning,
                            'Meaning_URL': meaning_url
                        }
                    else:
                        return {
                            'Hindwi_Status': 'PARSE_ERROR',
                            'Card_Word': '',
                            'Meaning': '',
                            'Meaning_URL': ''
                        }
                else:
                    return {
                        'Hindwi_Status': 'NOT_FOUND',
                        'Card_Word': '',
                        'Meaning': '',
                        'Meaning_URL': ''
                    }
            elif resp.status_code == 429:
                print(f"    [429 Rate Limit] Backing off for {5 * attempt}s...")
                time.sleep(5 * attempt)
            else:
                time.sleep(1)
        except Exception as e:
            if attempt == retries:
                return {
                    'Hindwi_Status': 'ERROR',
                    'Card_Word': '',
                    'Meaning': str(e),
                    'Meaning_URL': ''
                }
            time.sleep(1.5)
            
    return {
        'Hindwi_Status': 'ERROR',
        'Card_Word': '',
        'Meaning': 'Max retries exceeded',
        'Meaning_URL': ''
    }


def main():
    if not os.path.exists(INPUT_EXCEL):
        print(f"Error: Could not find '{INPUT_EXCEL}' in current directory.")
        sys.exit(1)

    print(f"Opening '{INPUT_EXCEL}' (Sheet: '{SHEET_NAME}')...")
    wb = openpyxl.load_workbook(INPUT_EXCEL, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        print(f"Error: Sheet '{SHEET_NAME}' not found in {wb.sheetnames}")
        sys.exit(1)

    ws = wb[SHEET_NAME]
    
    # Read words from column 1
    words_to_process = []
    for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if r_idx == 1:
            continue  # Skip header
        if not row or not row[0]:
            continue
        word = str(row[0]).strip()
        words_to_process.append((r_idx, word))

    total = len(words_to_process)
    print(f"Total words found in '{SHEET_NAME}': {total}")

    checkpoint = load_checkpoint()
    session = requests.Session()
    session.headers.update(HEADERS)

    batch_to_save = []
    processed_count = len(checkpoint)
    
    print(f"Starting verification... ({processed_count}/{total} already done)")

    for idx, (row_num, word) in enumerate(words_to_process, start=1):
        if word in checkpoint:
            continue

        result = query_hindwi(session, word)
        record = {
            'Index': row_num,
            'Word': word,
            'Hindwi_Status': result['Hindwi_Status'],
            'Card_Word': result['Card_Word'],
            'Meaning': result['Meaning'],
            'Meaning_URL': result['Meaning_URL']
        }
        batch_to_save.append(record)
        checkpoint[word] = record
        processed_count += 1

        print(f"[{processed_count}/{total}] Row {row_num:5d}: {word:15s} -> {result['Hindwi_Status']:10s} (Card: {result['Card_Word'] or '-'})")

        if len(batch_to_save) >= SAVE_INTERVAL:
            append_results_to_csv(batch_to_save)
            batch_to_save.clear()

        time.sleep(REQUEST_DELAY)

    if batch_to_save:
        append_results_to_csv(batch_to_save)
        batch_to_save.clear()

    print("\nAll words processed! Creating final enriched Excel file...")
    # Export full results back into an enriched Excel file
    try:
        wb_full = openpyxl.load_workbook(INPUT_EXCEL)
        ws_full = wb_full[SHEET_NAME]
        
        # Add headers in columns D, E, F, G
        ws_full.cell(row=1, column=4, value="Hindwi_Status")
        ws_full.cell(row=1, column=5, value="Hindwi_Card_Word")
        ws_full.cell(row=1, column=6, value="Hindwi_Meaning")
        ws_full.cell(row=1, column=7, value="Hindwi_Meaning_URL")
        
        for r_idx, word in words_to_process:
            data = checkpoint.get(word)
            if data:
                ws_full.cell(row=r_idx, column=4, value=data['Hindwi_Status'])
                ws_full.cell(row=r_idx, column=5, value=data['Card_Word'])
                ws_full.cell(row=r_idx, column=6, value=data['Meaning'])
                ws_full.cell(row=r_idx, column=7, value=data['Meaning_URL'])
                
        wb_full.save(OUTPUT_EXCEL)
        print(f"[SUCCESS] Final Excel saved to: '{OUTPUT_EXCEL}'")
    except Exception as e:
        print(f"Warning: Could not save final Excel: {e}. All data is safely saved in '{PROGRESS_CSV}'.")


if __name__ == '__main__':
    main()
