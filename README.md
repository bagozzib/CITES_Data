# CITES Attendee Data Extraction and Analysis

## Table of Contents
- [Description](#description)
- [Installation](#installation)
- [Usage](#usage)
- [File Structure](#folder-structure)
- [Contributors](#contributors)
- [License](#license)


## Description:
The code in this repository extracts, classifies, and formats attendee records from the Convention on International Trade in Endangered Species of Wild Fauna and Flora (CITES) Conference of the Parties (CoP) rosters (CoP1–CoP19). Source documents are mixed text/scan PDFs with multi-column layouts and multilingual headings; the pipeline performs layout-aware parsing/OCR, detects delegation headers, identifies person starts, and normalizes names. The primary attendee output consists of Delegation, Honorific, Person Name, and Affiliation.
In addition, the repo includes routines to standardize person names (e.g., LAST, First → First Last), harmonize delegation names, flag multilingual delegation strings, add ISO/COW country codes, derive CoP year and host city, geolocate affiliations and CoP cities (lat/long), compute attendee–CoP distances, flag likely country-centroid geocodes, and construct a female indicator by combining honorifics with a first-name–based gender guess.

## Installation:
To set up the project environment, follow these steps:

1. **Clone the repository:**
   ```
   git clone https://github.com/bagozzib/CITES_Data.git
   ```
   
2. **Install the required packages:**
   ```
   pip install -r requirements.txt
   ```

## Usage:
   For detailed usage instructions, please refer to the project [UNFCCC Project Code Execution Steps](https://github.com/bagozzib/UNFCCC-Attendance-Data/wiki/UNFCCC-Project-Code--Execution-Steps) WIKI.

 ## Folder Structure:
   - **master_data**: This directory holds the final dataset as a definitive CSV file.
        - Files:
           -  cites.cops.csv
        
   - **python_files**: These scripts are designed for PDF text extraction and classification, facilitating the generation of CSV files.
        - Files:
           - extract_pdf_data.py: Orchestrates PDF parsing; routes pages to the appropriate extractor (text vs. scan; single/dual column), applies page-level cleanup, and writes the initial rows.

          - processing_data.py – Post-extraction cleaning and harmonization: fixes wrap/merge artifacts, drops pagination noise, normalizes whitespace, resolves multilingual Delegation strings (Belgium/Bélgica/Belgique → Belgium), and de-duplicates.

          - standardize_person_names.py – Normalizes names to First Last (LAST, First and LAST First … → First Last; supports multi-token surnames).

          - gender_guess.py – Adds a GenderGuess column (and a Female indicator) using first-name inference plus honorifics (Namsor API workflow).

          - get_lat_lang.py – Geocodes affiliations and CoP host cities to latitude/longitude (Nominatim/ArcGIS with caching/backoffs).

          - city_centroid.py – Flags likely country-centroid geocodes (is_country_centroid): reverse-geocodes the point to a country, compares to country reference points from two sources, and applies distance thresholds that scale with country size.

          - haversine_distance.py – Computes great-circle distance (km) between attendee coordinates and the corresponding CoP city coordinates.     

   - **r_code**: This directory contains R scripts designed for functions such as data cleaning and validation.
       - Files:
          - FinalDataCleaning.R – Reads CITES Extracted Data V2.xlsx, normalizes honorifics, builds Female/Male from honorifics and gender-guess scores, standardizes Status and creates Party/Observer dummies, harmonizes Delegation names (multi-language variants, historical names, typos), fixes known anomalies, adds COW country codes, and writes the final analysis file cites.cops.csv. Prints basic validation stats.
          - Descriptives.R – Reads cites.cops.csv, reports column-wise missingness, and generates: MF.pdf (Female vs. Male over time), PO.pdf (Party vs. Observer over time), ShadedMap.pdf (choropleth of attendee counts), worldpoints.pdf (all geocoded points), and delegation_wordcloud.png (observer delegations).
          - CITES Extracted Data V2.xlsx – Intermediate dataset emitted by the Python extractor; input to FinalDataCleaning.R.
          - spatialvalidation.csv – 1,000-row sample for manual lat/long checks.
          - spatialvalidation DB.csv – the manually coded results.

     
   - **requirements.txt**: This file enumerates the Python dependencies necessary for the project.
     
## Conclusion
  - We built a reproducible pipeline that reliably extracts and standardizes CITES COP attendee records from heterogeneous PDFs—combining layout-aware parsing/OCR, targeted heuristics (underlines/all-caps, honorific & dotted-initial name rules, multilingual slashes, email anchors), NLP checks, and light manual review—to produce consistent Delegation, Honorific, Person Name, and Affiliation fields ready for analysis.
## Contributors:
   - Benjamin E. Bagozzi (Corresponding author: bagozzib@udel.edu)
   - Daria Blinova
   - Rakesh Emuru
     
## License

#### This project is licensed under the Creative Commons Attribution 4.0 International License (CC-BY-4.0).
   - This license allows reusers to distribute, remix, adapt, and build upon the material in any medium or format, so long as attribution is given to the creator. The license allows for commercial use.

For more details about this license, please visit the [Creative Commons Attribution 4.0 International License webpage](https://creativecommons.org/licenses/by/4.0/).



   
