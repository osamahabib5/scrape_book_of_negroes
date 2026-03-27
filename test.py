import os
import re
import pandas as pd
from docx import Document

# ── Configuration ───────────────────────────────────────────────────────────
REFERENCE_EXCEL = "book_of_negroes_original.xlsx"
FOLDER_PATH = "Book_of_Negroes"
OUTPUT_FILE = "Black_Loyalist_Directory_Consolidated.xlsx"

# ── Cleaning Helpers ──────────────────────────────────────────────────────────
def _clean_name(val):
    """Removes brackets, extra spaces, and trailing punctuation from names."""
    if not isinstance(val, str): return ""
    val = re.sub(r"[\[\]\{\}\(\)]", "", val)
    return re.sub(r"\s+", " ", val).strip().strip(",.;")

def clean_val(val):
    """Clean specific noise words and whitespace from strings."""
    if not isinstance(val, str) or val in ["-", "N/A", "", "None"]: return "-"
    val = re.split(r"[\(\)]|Born\s+free|belonging", val, flags=re.I)[0]
    val = re.sub(r"\s+", " ", val).strip().rstrip(",.;")
    return val if val else "-"

# ── Extraction Logic ──────────────────────────────────────────────────────────
def extract_header_info(line):
    """
    Extracts Ship Name, City, and Commander.
    Logic: If line starts with Ship/Sloop/etc, Ship Name is text between keyword and 'bound'.
    If not, Ship Name is all text before 'bound'.
    """
    if "bound for" in line.lower():
        # Split into [Prefix/Ship Name] and [Destination + Commander]
        parts = re.split(r"\s+bound\s+for\s+", line, flags=re.I)
        ship_part = parts[0].strip()
        rest = parts[1].strip()
        
        # Rule: Handle prefixes vs. no prefixes for Ship Name
        prefix_pattern = r"^(Ship|Brig|Sloop|Schooner|Brigantine|Snow)\s+(.*)$"
        m_prefix = re.match(prefix_pattern, ship_part, re.I)
        if m_prefix:
            ship_name = m_prefix.group(2).strip()
        else:
            ship_name = ship_part
            
        # Refined City and Commander isolation
        # List sorted by length descending to match full city names first
        cities = [
            "River St. John's", "St. John's River", "St. John's", 
            "Port Roseway", "Halifax", "Annapolis Royal", 
            "Spithead & Germany", "Shelburne", "River St. Johns"
        ]
        
        found_city, commander = "-", "-"
        for city in sorted(cities, key=len, reverse=True):
            if re.search(re.escape(city), rest, re.I):
                found_city = city
                # The text following the city is usually the commander
                cmd_raw = re.split(re.escape(city), rest, flags=re.I)[-1].strip()
                if cmd_raw:
                    # Strip ", Master" suffix if present
                    cmd_raw = re.sub(r",?\s*Master$", "", cmd_raw, flags=re.I).strip()
                    commander = cmd_raw
                break
        
        if found_city == "-": found_city = rest # Fallback to raw rest if no city matches

        return {
            "ship": clean_val(ship_name), 
            "city": found_city, 
            "cmd": clean_val(commander)
        }
    
    # Check for "On Board the Ship... Master" pattern without "bound for"
    m_master = re.search(r"(?:On\s+Board\s+the\s+)?(?:Ship|Brig|Sloop|Schooner)\s+(.*?)\s+([A-Z][A-Za-z\.]+(?:\s+[A-Z][A-Za-z\.]+)*),\s+Master", line, re.I)
    if m_master:
        return {"ship": clean_val(m_master.group(1)), "cmd": clean_val(m_master.group(2)), "city": "-"}

    return None

def extract_enslaver(line):
    """
    Extracts the enslaver name from the record line.
    """
    # 1. Parentheses check (ignoring noise like "born free" or physical descriptions)
    pm = re.search(r"\(([^)]+)\)", line)
    if pm:
        raw = pm.group(1).strip()
        noise = r"born\s+free|own\s+bottom|claims\s+to\s+be|aged|years|wench|fellow|stout|healthy"
        if not re.search(noise, raw, re.I):
            return _clean_name(raw)
    
    # 2. Key phrase check (Property of, Slave to, etc.)
    sm = re.search(r"(?:property\s+of|slave\s+to|lived\s+with)\s+([A-Z][A-Za-z\.]+(?:\s+[A-Z][A-Za-z\.]+)*)", line, re.I)
    if sm:
        return _clean_name(sm.group(1))
        
    return "-"

# ── Reference Loader ──────────────────────────────────────────────────────────
def load_reference(path):
    if not os.path.exists(path): 
        print(f"Warning: {path} not found.")
        return []
    try:
        df = pd.read_excel(path, dtype=str).fillna("")
        df.columns = [c.strip() for c in df.columns]
        lookup_list = []
        for _, row in df.iterrows():
            lookup_list.append({
                "ship_norm": clean_val(row.get("Ship_Name", row.get("Ship", ""))).lower(),
                "name_norm": _clean_name(row.get("Name", "")).lower(),
                "Ref_Page": row.get("Ref Page", row.get("Ref_Page", row.get("Page", "-"))),
                "Primary_Source_2": row.get("Primary_Source 2", "-")
            })
        return lookup_list
    except Exception as e:
        print(f"Error loading Excel: {e}"); return []

def lookup_excel(ref_list, ship_name, first_name, surname):
    s_norm = clean_val(ship_name).lower()
    full_n_norm = _clean_name(f"{first_name} {surname}").lower()
    for item in ref_list:
        if item["ship_norm"] == s_norm and item["name_norm"] == full_n_norm:
            return item
    return None

# ── Main ──────────────────────────────────────────────────────────────────────
def process_word_docs():
    files = [
        ("Book One Part One", "Book_One_Part_One_of_the_Book_of_Negroes.docx"),
        ("Book One Part Two", "Book_One_Part_Two_of_the_Book_of_Negroes.docx"),
        ("Book Two", "Book_Two.docx"), 
        ("Book Three", "Book_Three.docx")
    ]

    ref_list = load_reference(REFERENCE_EXCEL)
    all_records, global_id = [], 1

    # Persist ship info across lines and documents
    current_ship, current_commander, current_city = "-", "-", "-"
    
    for book_label, filename in files:
        file_path = os.path.join(FOLDER_PATH, filename)
        if not os.path.exists(file_path): continue
            
        doc = Document(file_path)
        print(f"Processing {book_label}...")

        for para in doc.paragraphs:
            line = para.text.strip()
            if not line: continue
            
            # --- Header Detection ---
            is_ship_line = any(kw in line.lower() for kw in ["bound for", "boarding", "on board", "master"])
            if is_ship_line and not re.search(r"^\d+", line):
                header = extract_header_info(line)
                if header:
                    current_ship = header["ship"]
                    current_commander = header["cmd"]
                    current_city = header["city"]
                continue

            # --- Individual Record Detection ---
            if "," in line and re.search(r"\d+", line) and not line.startswith(("[", "In pursuance", "Inspected")):
                raw_name_part = line.split(",")[0].strip()
                name_parts = raw_name_part.split(None, 1)
                
                first_name = name_parts[0] if name_parts else "-"
                surname = name_parts[1] if len(name_parts) > 1 else "-"
                
                age_m = re.search(r"(?:aged\s+)?(\d{1,3})", line, re.I)
                age_val = int(age_m.group(1)) if age_m else "-"
                gender = "Female" if any(w in line.lower() for w in ["woman", "wench", "girl"]) else "Male"
                
                xl_match = lookup_excel(ref_list, current_ship, first_name, surname)

                all_records.append({
                    "ID": global_id, 
                    "Ref_Page": xl_match["Ref_Page"] if xl_match else "-",
                    "Book": book_label, 
                    "Ship_Name": current_ship, 
                    "Commander": current_commander,
                    "Enslaver": extract_enslaver(line), 
                    "First_Name": _clean_name(first_name), 
                    "Surname": _clean_name(surname),
                    "Gender": gender, 
                    "Age": age_val, 
                    "Arrival_Port_City": current_city,
                    "Notes": line,
                    "Primary_Source_2": xl_match["Primary_Source_2"] if xl_match else "-"
                })
                global_id += 1

    df = pd.DataFrame(all_records)
    df = df.fillna("-").replace(["", "nan", "None"], "-")
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"SUCCESS: {len(df)} records generated.")

if __name__ == "__main__":
    process_word_docs()