arkdown

# 🩸 Blood Transfusion Analyzer

A tool for normalizing blood transfusion data, calculating optimal inventory of blood components, and analyzing crossmatch transfusions

**Python:** 3.8+ | **License:** MIT | **Version:** 8.0

🇷🇺 [Russian](README.md) | 🇬🇧 [English](README.en.md)

---

## 📋 Table of Contents

- [📋 Table of Contents](#-table-of-contents)
- [🎯 Purpose](#-purpose)
- [📁 Project Structure](#-project-structure)
- [🚀 Quick Start](#-quick-start)
  - [For end users (without Python)](#for-end-users-without-python)
  - [For developers (with Python)](#for-developers-with-python)
- [📦 Components](#-components)
  - [1. Data Normalizer (`blood_normalizer_v_8_stable.py`)](#1-data-normalizer-blood_normalizer_v_8_stablepy)
  - [2. Crossmatch Transfusion Analyzer (`analyze_crossmatch_detailed_v4_1.py`)](#2-crossmatch-transfusion-analyzer-analyze_crossmatch_detailed_v4_1py)
- [📊 Example Output](#-example-output)
- [📚 Documentation](#-documentation)
- [📄 License](#-license)
- [📧 Contacts](#-contacts)
- [👨‍💻 Author](#-author)

---

## 🎯 Purpose

The project is designed to automate the processing of blood transfusion data in healthcare organizations.

**Main tasks:**
- Normalization of blood groups (ABO and Rh) considering A2 and A2B subgroups
- Identification of blood component types (red blood cells, plasma, platelets, cryoprecipitate)
- Analysis of crossmatch transfusions according to compatibility rules (Order No. 1134н)
- Calculation of optimal blood component inventory (in development)

---

## 📁 Project Structure

transfusion_scripts/
│
├── blood_normalizer_v_8_stable.py # Data normalization module
├── analyze_crossmatch_detailed_v4_1.py # Crossmatch transfusion analyzer
│
├── dist/ # Compiled EXE files
│ ├── BloodNormalizer.exe
│ └── BloodAnalyzer.exe
│
├── enter_data/ # Input data folder
│ └── transfusion_register_2023_2024_2025.xlsx
│
├── reports/ # Reports folder (auto-created)
│ └── crossmatch_analysis_v4_YYYYMMDD_HHMMSS.xlsx
│
├── archive_version/ # Archive of old versions
│ ├── preprocess_blood_data_v1_0.py
│ └── README.txt
│
├── requirements.txt # Python dependencies
├── run_normalizer.bat # Normalizer launcher
├── run_analyzer.bat # Analyzer launcher
└── README.en.md # This file
text


---

## 🚀 Quick Start

### For end users (without Python)

1. **Download** the project folder to your PC
2. **Copy** the source Excel file to the `enter_data/` folder
3. **Run** the normalizer:
   - Double-click `run_normalizer.bat` or `dist/BloodNormalizer.exe`
   - Select the Excel file in the dialog window
   - Wait for processing to complete
4. **Run** the analyzer:
   - Double-click `run_analyzer.bat` or `dist/BloodAnalyzer.exe`
   - The analyzer will automatically find the normalized file
   - The report will be saved to the `reports/` folder

### For developers (with Python)

**1. Clone the repository:**
```bash
git clone https://github.com/skokovso/BloodNormalizer.git
cd BloodNormalizer

2. Install dependencies:
bash

pip install -r requirements.txt

3. Run the normalizer:
bash

python blood_normalizer_v_8_stable.py

4. Run the analyzer:
bash

python analyze_crossmatch_detailed_v4_1.py

5. Compile to EXE (optional):
bash

pyinstaller --onefile --console --name "BloodNormalizer" --add-data "enter_data;enter_data" --add-data "reports;reports" blood_normalizer_v_8_stable.py

pyinstaller --onefile --console --name "BloodAnalyzer" --add-data "reports;reports" analyze_crossmatch_detailed_v4_1.py

📦 Components
1. Data Normalizer (blood_normalizer_v_8_stable.py)

Purpose: Preparation of raw data for analysis.

Features:

    Normalizes blood groups by ABO system (O, A, B, AB, A2, A2B)

    Normalizes Rh factor (Rh+, Rh-)

    Identifies component types (red blood cells, plasma, cryoprecipitate, platelets)

    Processes transfusion time (columns "from" and "to")

    Extracts short phenotype (C/c and E/e patterns)

    Creates Normalization_Issues sheet for quality control

    Automatically detects file name and adds _by_filename suffix

Input: Excel file with transfusions (any structure, automatic header detection)

Output: Excel file with added normalized columns

Added Columns:
Column	Description
Blood_Group_Patient_Norm	Normalized patient blood group
Rh_Patient_Norm	Normalized patient Rh factor
Blood_Group_Patient_Full	Full patient group (A+, A2-, A2B+)
Blood_Group_Source	Source of group ('patient', 'environment', 'check')
Blood_Group_Env_Norm	Normalized donor blood group
Rh_Env_Norm	Normalized donor Rh factor
Blood_Group_Env_Full	Full donor group
Component_Type	Component type
Patient_Issue	Patient issue description
Donor_Issue	Donor issue description
Start_Time	Transfusion start time (HH:MM:SS)
End_Time	Transfusion end time (HH:MM:SS)
Duration_Minutes	Transfusion duration in minutes
Short_Phenotype	Phenotype in CcEe format
2. Crossmatch Transfusion Analyzer (analyze_crossmatch_detailed_v4_1.py)

Purpose: Analysis of crossmatch transfusions according to compatibility rules (Order No. 1134н).

Features:

    Classifies transfusions as: same-group, compatible crossmatch, incompatible

    Uses compatibility rules for each component type

    Analyzes dynamics by year (2023-2025)

    Outputs statistics by site and doctor

    Creates detailed Excel report

Compatibility Rules (Order No. 1134н):
Component	Compatibility Rules
Red Blood Cells (RBC)	ABO + Rh (considering A2, A2B)
Plasma	ABO only (0→any, A→A/AB, B→B/AB, AB→AB)
Cryoprecipitate	Any groups compatible
Platelets	Any groups compatible

Input: Normalized Excel file (auto-detected)

Output: Excel report with sheets:

    Summary — table by year and component

    Incompatible — list of incompatible transfusions

    Compatible_Crossmatch — list of allowed crossmatch transfusions

    Sites — dynamics by site

📊 Example Output

Console output (analyzer):
text

🩸 CROSSMATCH TRANSFUSION ANALYSIS v4.0
============================================================

📂 Loading: transfusion_register_2023_2025_by_blood_normalizer_v_8_stable.xlsx
   ✅ Loaded 9710 records

============================================================
📊 ANALYSIS BY COMPONENT TYPES
============================================================

   📊 Red Blood Cells:
   ─────────────────────────────────────────────────────────────────
   Year    Total    Same-Group   Compatible   Incompatible
   ─────────────────────────────────────────────────────────────────
   2023    1772     1633         128          10
   2024    2273     2052         209          11
   2025    2756     2432         312          12
   ─────────────────────────────────────────────────────────────────
   TOTAL   6801     6117         649          33

📈 DYNAMICS 2023 → 2025:
   Share growth: 7.2% → 11.3% (+4.1%)
   Quantity growth: 128 → 312 (+184 transfusions)

📚 Documentation
Versions
Version	Changes
v8.0	Extended normalization algorithm, automatic file suffix
v7.0	Digital garbage check, separate issue sheet
v6.0	Time and phenotype processing
v5.0	Improved validation, fallback to donor group
v4.0	Full crossmatch analysis with site and doctor breakdown
v2.7	Digital garbage check (>3 digits)
v2.6	Added time and phenotype processing
v2.5	Fixed SZP definition
Dependencies

Specified in requirements.txt:
text

pandas>=2.0.0
openpyxl>=3.1.0
numpy>=1.24.0
tkinter (included with Python)

📄 License

Distributed under the MIT license. See LICENSE file for details.
📧 Contacts

    ✉️ Email: skokovso@ya.ru

    🖥️ GitVerse: gitverse.ru/skokovso/BloodNormalizer

    🐙 GitHub: github.com/skokovso/BloodNormalizer

    🏢 GitVerse (Mirror): hub.mos.ru/skokovso/bloodnormalizer

    🐛 Report an Issue: GitVerse Issues

👨‍💻 Author

Skokov S.O. — transfusion medicine specialist

🩸 Current version: v8.0 (July 2026)