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
    if not isinstance(val, str): return ""
    val = re.sub(r"[\[\]\{\}\(\)]", "", val)
    return re.sub(r"\s+", " ", val).strip().strip(",.;")

def clean_val(val):
    if not isinstance(val, str) or val in ["-", "N/A", "", "None", "nan"]: return "-"
    val = re.split(r"[\(\)]|Born\s+free|belonging", val, flags=re.I)[0]
    val = re.sub(r"\s+", " ", val).strip().rstrip(",.;")
    return val if val else "-"

# ── Logic: Race, Ethnicity, Description ──────────────────────────────────────
def extract_race_details(line):
    line_l = line.lower()
    race, ethnicity, description = "Black", "African American", "Black"
    
    # Keywords for Mixed Race
    mixed_keywords = ["mulatto", "indian", "span.", "spanish", "half indian", "between"]
    
    if any(k in line_l for k in mixed_keywords):
        ethnicity = "Mixed Race"
        if "mulatto" in line_l:
            description = "Mulatto"
            race = "Mulatto"
        if "indian" in line_l and "span" in line_l:
            race = "Indian/Spanish"
            description = "Mulatto"
        elif "half indian" in line_l:
            race = "Half Indian"
            description = "Mulatto"
    
    return race, ethnicity, description

# ── Logic: Geography Extraction ──────────────────────────────────────────────
def extract_geo_from_text(line):
    # Common 18th Century States/Colonies
    states = ["Virginia", "Maryland", "New Jersey", "New York", "Georgia", "South Carolina", "North Carolina", "Pennsylvania"]
    found_state = "-"
    found_port = "-"

    for state in states:
        if state in line:
            found_state = state
            # Try to grab the word before the state as the port/county
            match = re.search(r"([^,.;]+)\s*,\s*" + re.escape(state), line)
            if match:
                found_port = match.group(1).strip()
            break
            
    if found_port == "-" and "Jamaica South" in line:
        found_port = "Jamaica South"

    return found_port, found_state

# ── Logic: Advanced Gender ──────────────────────────────────────────────────
def determine_gender(line, age_val):
    line_l = line.lower()
    is_child = False
    
    # Determine if child by age
    if isinstance(age_val, (int, float)) and age_val < 18:
        is_child = True
    # Determine if child by keywords
    elif any(k in line_l for k in ["boy", "girl", "child"]):
        is_child = True
        
    if any(k in line_l for k in ["woman", "wench", "girl", "negress"]):
        return "Child Female" if is_child else "Female"
    else:
        return "Child Male" if is_child else "Male"

# ── Extraction Logic: Headers ────────────────────────────────────────────────
def extract_header_info(line):
    if "bound for" in line.lower():
        parts = re.split(r"\s+bound\s+for\s+", line, flags=re.I)
        ship_part = parts[0].strip()
        rest = parts[1].strip()
        
        prefix_pattern = r"^(Ship|Brig|Sloop|Schooner|Brigantine|Snow)\s+(.*)$"
        m_prefix = re.match(prefix_pattern, ship_part, re.I)
        ship_name = m_prefix.group(2).strip() if m_prefix else ship_part
            
        cities = ["River St. John's", "St. John's River", "St. John's", "Port Roseway", "Halifax", "Annapolis Royal", "Shelburne"]
        found_city, commander = "-", "-"
        for city in sorted(cities, key=len, reverse=True):
            if re.search(re.escape(city), rest, re.I):
                found_city = city
                cmd_raw = re.split(re.escape(city), rest, flags=re.I)[-1].strip()
                commander = re.sub(r",?\s*Master$", "", cmd_raw, flags=re.I).strip() if cmd_raw else "-"
                break
        
        return {"ship": clean_val(ship_name), "city": found_city, "cmd": clean_val(commander)}
    return None

# ── Main Process ─────────────────────────────────────────────────────────────
def process_word_docs():
    files = [
        ("Book One Part One", "Book_One_Part_One_of_the_Book_of_Negroes.docx"),
        ("Book One Part Two", "Book_One_Part_Two_of_the_Book_of_Negroes.docx"),
        ("Book Two", "Book_Two.docx"), 
        ("Book Three", "Book_Three.docx")
    ]

    # Load Excel Reference
    print("Loading Reference Excel...")
    try:
        df_ref = pd.read_excel(REFERENCE_EXCEL, dtype=str).fillna("-")
        df_ref.columns = [c.strip() for c in df_ref.columns]
    except:
        df_ref = pd.DataFrame()

    all_records, global_id = [], 1
    current_ship, current_commander, current_city = "-", "-", "-"
    
    # Family Context Memory
    last_male_first, last_male_sur = "-", "-"
    last_female_first, last_female_sur = "-", "-"

    for book_label, filename in files:
        file_path = os.path.join(FOLDER_PATH, filename)
        if not os.path.exists(file_path): continue
        doc = Document(file_path)
        
        for para in doc.paragraphs:
            line = para.text.strip()
            if not line: continue
            
            # Header Detection
            if any(kw in line.lower() for kw in ["bound for", "master"]) and not re.search(r"^\d+", line):
                header = extract_header_info(line)
                if header:
                    current_ship, current_commander, current_city = header["ship"], header["cmd"], header["city"]
                continue

            # Individual Record Detection
            if "," in line and re.search(r"\d+", line) and not line.startswith(("[", "In pursuance")):
                raw_name_part = line.split(",")[0].strip()
                name_parts = raw_name_part.split(None, 1)
                f_name = name_parts[0] if name_parts else "-"
                s_name = name_parts[1] if len(name_parts) > 1 else "-"
                
                age_match = re.search(r"(\d{1,3}(?:\s?½)?)", line)
                age_str = age_match.group(1) if age_match else "-"
                try: age_val = float(age_str.replace("½", ".5"))
                except: age_val = 0

                # 1. Race/Ethnicity/Description
                race, ethnicity, desc = extract_race_details(line)

                # 2. Gender Logic
                gender_cat = determine_gender(line, age_val)

                # 3. Family Logic (Scenario 1, 2, 3)
                f_father, s_father, f_mother, s_mother = "-", "-", "-", "-"
                if "daughter" in line.lower() or "son" in line.lower() or "child" in line.lower():
                    if "their" in line.lower() or ("his" in line.lower() and "wife" in globals().get('last_line','')):
                        f_father, s_father = last_male_first, last_male_sur
                        f_mother, s_mother = last_female_first, last_female_sur
                    elif "her" in line.lower():
                        f_mother, s_mother = last_female_first, last_female_sur
                    elif "his" in line.lower():
                        f_father, s_father = last_male_first, last_male_sur

                # Update Family Memory for next lines
                if "Male" in gender_cat and "Child" not in gender_cat:
                    last_male_first, last_male_sur = f_name, s_name
                elif "Female" in gender_cat and "Child" not in gender_cat:
                    last_female_first, last_female_sur = f_name, s_name

                # 4. Excel Lookup (Origination, Source, etc)
                xl_row = df_ref[(df_ref['Ship_Name'] == current_ship) & (df_ref['Name'].str.contains(f_name, na=False))].head(1)
                
                # Geography Fallback
                word_port, word_state = extract_geo_from_text(line)
                
                all_records.append({
                    "ID": global_id,
                    "Ref_Page": xl_row['Ref Page'].values[0] if not xl_row.empty else "-",
                    "Book": book_label,
                    "Ship_Name": current_ship,
                    "Commander": current_commander,
                    "First_Name": _clean_name(f_name),
                    "Surname": _clean_name(s_name),
                    "Father_FirstName": f_father,
                    "Father_Surname": s_father,
                    "Mother_FirstName": f_mother,
                    "Mother_Surname": s_mother,
                    "Gender": gender_cat,
                    "Age": age_str,
                    "Race": race,
                    "Ethnicity": ethnicity,
                    "Description": desc,
                    "Origination_Port": xl_row['Origination Port'].values[0] if not xl_row.empty and xl_row['Origination Port'].values[0] != "-" else word_port,
                    "Origination_State": word_state,
                    # "Country": xl_row['Country'].values[0] if not xl_row.empty else "-",
                    "Departure_Port": xl_row['Departure_Port'].values[0] if not xl_row.empty else "-",
                    "Departure_Date": xl_row['Departure_Date'].values[0] if not xl_row.empty else "-",
                    "Arrival_Port_City": current_city,
                    "Primary_Source_1": xl_row['Primary_Source 1'].values[0] if not xl_row.empty else "-",
                    "Primary_Source_2": xl_row['Primary_Source 2'].values[0] if not xl_row.empty else "-",
                    "Notes": line
                })
                global_id += 1

    pd.DataFrame(all_records).to_excel(OUTPUT_FILE, index=False)
    print(f"Done. {len(all_records)} records saved.")

if __name__ == "__main__":
    process_word_docs()