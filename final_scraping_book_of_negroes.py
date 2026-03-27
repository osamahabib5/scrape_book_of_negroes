import os
import re
import pandas as pd

# ── Configuration ───────────────────────────────────────────────────────────
ORIGINAL_FILE = "Black_Loyalist_Directory_Consolidated.xlsx"
MISSING_FILE = "Validated_Missing_Records.xlsx"
OUTPUT_FILE = "Black_Loyalist_Directory_Final.xlsx"

FILE_TO_BOOK_MAP = {
    "Book_One_Part_One_of_the_Book_of_Negroes.docx": "Book One Part One",
    "Book_One_Part_Two_of_the_Book_of_Negroes.docx": "Book One Part Two",
    "Book_Two.docx": "Book Two",
    "Book_Three.docx": "Book Three"
}

# ── Cleaning & Filter Helpers ────────────────────────────────────────────────
def _clean_text(val):
    if not isinstance(val, str): return ""
    val = re.sub(r"[\[\]\{\}\(\)]", "", val)
    return re.sub(r"\s+", " ", val).strip().strip(",.;")

def should_ignore(line):
    """
    Implements your ignore logic:
    1. Dates/Years (e.g., 10 July 1783 or just 1783)
    2. 'bound' keyword (Ship headers)
    3. '[Signed]' keyword
    """
    line_l = line.lower()
    
    # Match dates like '10 July 1783' or 'July 1783'
    date_pattern = r"\b\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}\b"
    # Match standalone years like 1783
    year_pattern = r"\b(17|18)\d{2}\b"
    
    if re.search(date_pattern, line, re.I) or re.search(year_pattern, line):
        return True
    if "bound" in line_l:
        return True
    if "[signed]" in line_l:
        return True
        
    return False

# ── Transformation Logic: Race & Ethnicity ──────────────────────────────────
def transform_race_data(line):
    line_l = line.lower()
    race, ethnicity, description = "Black", "African American", "Black"
    
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

# ── Transformation Logic: Geography ─────────────────────────────────────────
def transform_geo_data(line):
    states = ["Virginia", "Maryland", "New Jersey", "New York", "Georgia", "South Carolina", "North Carolina", "Pennsylvania"]
    found_state, found_port = "", ""
    
    for state in states:
        if state in line:
            found_state = state
            match = re.search(r"([^,.;]+)\s*,\s*" + re.escape(state), line)
            if match:
                found_port = match.group(1).strip()
            break
            
    if not found_port and "Jamaica South" in line:
        found_port = "Jamaica South"

    return found_port, found_state

# ── Transformation Logic: Gender & Age ──────────────────────────────────────
def transform_gender_age(line):
    line_l = line.lower()
    age_match = re.search(r"(\d{1,3}(?:\s?½)?)", line)
    age_str = age_match.group(1) if age_match else ""
    
    try:
        age_val = float(age_str.replace("½", ".5"))
    except:
        age_val = None

    is_child = False
    if (age_val is not None and age_val < 18) or any(k in line_l for k in ["boy", "girl", "child", "small boy"]):
        is_child = True
        
    if any(k in line_l for k in ["woman", "wench", "girl", "negress"]):
        gender = "Child Female" if is_child else "Female"
    else:
        gender = "Child Male" if is_child else "Male"
        
    return age_str, gender

# ── Main ETL Process ────────────────────────────────────────────────────────
def run_validation_merge():
    print(f"Extraction Phase Started...")
    try:
        df_orig = pd.read_excel(ORIGINAL_FILE, dtype=str).fillna("")
        df_miss = pd.read_excel(MISSING_FILE, dtype=str).fillna("")
    except Exception as e:
        print(f"Error reading files: {e}")
        return

    try:
        max_id = pd.to_numeric(df_orig["ID"], errors='coerce').max()
        curr_id = int(max_id) + 1 if not pd.isna(max_id) else 1
    except:
        curr_id = 1

    mem = {"male_f": "", "male_s": "", "female_f": "", "female_s": ""}
    new_rows = []

    print("Transformation Phase Started (Applying Filters)...")
    for _, row in df_miss.iterrows():
        notes = str(row.get("Notes", "")).strip()
        
        # ── APPLY IGNORE FILTERS ──
        if not notes or notes == "" or should_ignore(notes):
            continue

        # 1. Parsing Name
        raw_name = notes.split(",")[0].strip()
        name_parts = raw_name.split(None, 1)
        f_name = name_parts[0] if name_parts else ""
        s_name = name_parts[1] if len(name_parts) > 1 else ""

        # 2. Transformation Functions
        age_str, gender = transform_gender_age(notes)
        race, ethn, desc = transform_race_data(notes)
        port, state = transform_geo_data(notes)
        
        # 3. Family Logic
        f_father, s_father, f_mother, s_mother = "", "", "", ""
        notes_l = notes.lower()
        if any(k in notes_l for k in ["daughter", "son", "child"]):
            if "their" in notes_l:
                f_father, s_father, f_mother, s_mother = mem["male_f"], mem["male_s"], mem["female_f"], mem["female_s"]
            elif "her" in notes_l:
                f_mother, s_mother = mem["female_f"], mem["female_s"]
            elif "his" in notes_l:
                f_father, s_father = mem["male_f"], mem["male_s"]

        # Update Parent Memory if Adult
        if "Child" not in gender:
            if "Male" in gender:
                mem["male_f"], mem["male_s"] = f_name, s_name
            else:
                mem["female_f"], mem["female_s"] = f_name, s_name

        # 4. Map to final schema
        new_rows.append({
            "ID": curr_id,
            "Book": FILE_TO_BOOK_MAP.get(str(row.get("Source_Word_File", "")), ""),
            "Ship_Name": str(row.get("Ship_Name", row.get("Ship", ""))),
            "Commander": str(row.get("Commander_Name", row.get("Commander", ""))),
            "First_Name": _clean_text(f_name),
            "Surname": _clean_text(s_name),
            "Father_FirstName": f_father,
            "Father_Surname": s_father,
            "Mother_FirstName": f_mother,
            "Mother_Surname": s_mother,
            "Gender": gender,
            "Age": age_str,
            "Race": race,
            "Ethnicity": ethn,
            "Description": desc,
            "Origination_Port": port,
            "Origination_State": state,
            "Arrival_Port_City": str(row.get("Arrival_Port_City", "")),
            "Notes": notes,
            "Ref_Page": "", "Country": "", "Departure_Port": "", 
            "Departure_Date": "", "Primary_Source_1": "", "Primary_Source_2": ""
        })
        curr_id += 1

    # Load Phase
    df_new = pd.DataFrame(new_rows)
    df_final = pd.concat([df_orig, df_new], ignore_index=True).replace(["nan", "None"], "")
    
    df_final.to_excel(OUTPUT_FILE, index=False)
    print(f"Load Complete. {len(new_rows)} records added to {OUTPUT_FILE}.")

if __name__ == "__main__":
    run_validation_merge()