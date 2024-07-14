# Project: Excel Data Cleaning and Spell-Checking Automation

## Overview
This project aims to develop a program that automates the cleaning and spell-checking of data in Excel files. The data, initially extracted using OCR functions, often contains inconsistencies and spelling errors. This program ensures data integrity and correctness through automated processes.

## Objectives
- *Data Cleaning*: Identify and rectify inconsistencies, formatting errors, and invalid entries within the Excel file.
- *Spell Checking*: Verify and correct the spelling of text entries throughout the Excel file.
- *Changes Highlighting*: Visually highlight any modifications or discrepancies detected during the cleaning and spell-checking process.

## Features

### Data Cleaning
- Removes header columns and non-numerical values.
- Eliminates insufficient data.
- Removes empty rows and standardizes date formats.

### Spell Checking
- Integrates Groq API for spell-checking using a prompt to the API.
- Corrects spelling in alphabetic columns of the Excel sheet.

### Number Correction
- Uses custom mappings to replace strings in number columns extracted by OCR.
- Identifies and corrects patterns (e.g., Z/ as 21).

### Highlighting Changes
- Highlights differences between original and processed data for easy analysis.
- Uses random color generation to visually distinguish changes.
- Capitalizes the first letter of each word.

## Implementation

### Libraries and modules Used
- pandas
- datetime
- openpyxl
- re
- json
- groq
- dotenv
- os

### Steps
1. *Data Cleaning*:
   - Load Excel file.
   - Remove unnecessary headers and non-numerical values.
   - Standardize date formats.
   
2. *Spell Checking*:
   - Use Groq API to check and correct spelling.
   - Store responses in JSON format for further processing.
   
3. *Number Correction*:
   - Apply custom mappings to correct OCR-extracted number patterns.
   
4. *Highlighting*:
   - Compare specified columns and highlight discrepancies.
   - Use regex to clean unnecessary prompts after spell check.

### Automation
- The main script integrates all functions, iterates through all files in the input folder, processes each Excel file, and stores the output in the specified output directory.
- Allows user inputs for specific columns and number of rows to be processed.

## Future Advancements
- Integrate ChatGPT for improved spell-checking (cost consideration).
- Use local Llama model from LM Studio for offline processing (requires high space and GPU).
- Use Hugging Face Llama model for a cost-effective and space-efficient solution.

## Conclusion
This project automates the cleaning and spell-checking of OCR-processed Excel data, ensuring data integrity and correctness. Future advancements aim to enhance the spell-checking capabilities and provide more efficient solutions.
