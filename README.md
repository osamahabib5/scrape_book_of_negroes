# Book of Negroes Data Scraper

A Python-based project for scraping, consolidating, and validating historical records from the Book of Negroes documents and related archives.

## Project Overview

This project processes Word documents containing the Book of Negroes records and consolidates them with the Black Loyalist Directory. It extracts ship manifests, passenger information, and creates validated datasets for analysis.

---

## Main Files

### **book_of_negroes_scraper_v2.py**
Current production scraper for extracting ship and passenger data from Book of Negroes Word documents. 
- **Key Functions:**
  - `extract_header_info()`: Extracts ship name, city, and commander from manifest headers
  - `clean_name()` / `clean_val()`: Cleans and normalizes extracted text
  - Processes files from the `Book_of_Negroes/` folder
  - Outputs consolidated data to `Black_Loyalist_Directory_Consolidated.xlsx`

**Status:** Primary active scraper

---

### **validate_book_of_negroes_records.py**
Validation and reconciliation tool that cross-references scraped data with reference records.
- **Key Functions:**
  - `extract_ship_name()`: Isolates ship names from text using regex patterns
  - `get_word_content_ordered()`: Reads Word documents in a specific order
  - Handles multi-part Book entries (Book One Part One, Part Two, Book Two, Book Three)
  - Identifies missing or mismatched records
  - Outputs validation results to `Validated_Missing_Records.xlsx`

**Status:** Active validation tool

---

### **test.py**
Testing and experimentation script for ship name extraction and pattern matching.
- Used for validating regex patterns and extraction logic
- Tests various ship name formats and edge cases
- Helps debug parsing issues

**Status:** Development/testing utility

---

## Additional Python Scripts

Located in `Additional_Python_Scripts/` folder - these are previous versions and alternative approaches:

### **book_of_negroes_scraper_v1.py**
Earlier version of the main scraper. 
- First implementation with regex-based extraction
- Includes state/location parsing logic
- Reference point for evolution of scraping approach

**Status:** Archived (superseded by v2)

---

### **book_of_negroes_scraper.py**
Original scraper implementation.
- Basic extraction patterns for ship names and destinations
- Foundation for current production versions
- Uses word document parsing via python-docx

**Status:** Archived (historical reference)

---

### **book_of_negroes_llm_scraper.py**
Advanced LLM-based scraper using Groq's free API (Llama 3 70B).
- Uses natural language processing for more intelligent text extraction
- Handles complex text variations that regex cannot
- Requires Groq API key configuration
- Generates structured output with validated data

**Status:** Alternative/experimental approach

---

## Data Files

### **Input Files**
- `Book_of_Negroes/` - Folder containing Word documents (.docx files)
  - `Book_One_Part_One_of_the_Book_of_Negroes.docx`
  - `Book_One_Part_Two_of_the_Book_of_Negroes.docx`
  - `Book_Two.docx`
  - `Book_Three.docx`

- `book_of_negroes_original.xlsx` - Reference spreadsheet with original records

- `Black_Loyalist_Directory_Consolidated.xlsx` - Reference data for cross-checking

### **Output Files**
- `Black_Loyalist_Directory_Consolidated.xlsx` - Output from scraper_v2: consolidated manifest data
- `Validated_Missing_Records.xlsx` - Output from validation script: records with discrepancies and missing entries

---

## Setup & Dependencies

### Requirements
```bash
pip install pandas python-docx openpyxl groq requests beautifulsoup4
```

### For LLM Scraper
Get a free Groq API key at: https://console.groq.com

Then set environment variable:
```bash
export GROQ_API_KEY="your_key_here"
```

---

## Usage

### Run the main scraper (v2)
```bash
python book_of_negroes_scraper_v2.py
```

### Run validation
```bash
python validate_book_of_negroes_records.py
```

### Run tests
```bash
python test.py
```

### Run LLM-based scraper
```bash
python Additional_Python_Scripts/book_of_negroes_llm_scraper.py
```

---

## Project Structure

```
scrape_data_book_of_negroes/
├── book_of_negroes_scraper_v2.py          # Main production scraper
├── validate_book_of_negroes_records.py    # Validation & reconciliation
├── test.py                                # Testing script
├── book_of_negroes_original.xlsx          # Reference data
├── Black_Loyalist_Directory_Consolidated.xlsx  # Reference data
├── Validated_Missing_Records.xlsx         # Validation output
├── Book_of_Negroes/                       # Input Word documents
├── Additional_Python_Scripts/             # Previous versions & alternatives
│   ├── book_of_negroes_scraper_v1.py
│   ├── book_of_negroes_scraper.py
│   └── book_of_negroes_llm_scraper.py
├── README.md                              # This file
└── .gitignore                             # Git ignore patterns
```

---

## Notes

- All scripts perform extensive regex-based pattern matching and text normalization
- The validation script processes documents in a specific order to enable backtracking for accurate record matching
- Excel files with `~$` prefix are temporary files created by Excel (can be safely ignored)
- The project handles multiple name variations, abbreviations, and location formats from historical records

---

## Version History

- **v2** - Current production version (book_of_negroes_scraper_v2.py)
- **v1** - Previous state-aware implementation
- **v0** - Original proof-of-concept scraper
- **LLM** - Alternative AI-powered approach for complex extraction

---

## License

[Add your license information here]

---

## Contact

[Add your contact information here]
