from email import header
from scripts.spellcheck import process_excel
from scripts.num import correct_numbers_in_excel
from scripts.validation1 import clean
from scripts.highlighter import highlight
from scripts.finalcheck import highlight_cells
import os

input_folder = "data/input"
output_folder = "data/output"
Final_output_folder = "data/Fin_output"

os.makedirs(Final_output_folder, exist_ok=True)

header_row = int(input(f"Enter the row number to use as the header (1-indexed) for file: "))
columns_to_correct_str = input(f"Enter the Excel column indices to correct (comma separated) for file: ")
columns_to_correct = [col.strip() for col in columns_to_correct_str.split(',') if col.strip()]
columns_to_correct_idx = input(f"Enter the Excel column indices to correct number column (comma separated) for file: ")
columns_to_correct2 = [col.strip() for col in columns_to_correct_idx.split(',') if col.strip()]

for root, dirs, files in os.walk(input_folder):
    for filename in files:
        if filename.endswith(".xlsx"):
            rel_path = os.path.relpath(root, input_folder)
            output_folder_path = os.path.join(output_folder, rel_path)
            if not os.path.exists(output_folder_path):
                os.makedirs(output_folder_path)
            Final_output_folder_path = os.path.join(Final_output_folder, rel_path)
            if not os.path.exists(Final_output_folder_path):
                os.makedirs(Final_output_folder_path)

            input_file = os.path.join(root, filename)
            output_file = os.path.join(output_folder_path, f"spell_{filename}")
            cleaned_file = os.path.join(output_folder_path, f"cleaned_{filename}")
            output_file2 = os.path.join(output_folder_path, f"spell2_{filename}")
            output_file4 = os.path.join(output_folder_path, f"spell2_{filename}")
            output_file5 = os.path.join(Final_output_folder_path, f"outputof_{filename}")

            cleaned = clean(input_file, cleaned_file, header_row)
            op1 = process_excel(cleaned, output_file, columns_to_correct)

            op2 = correct_numbers_in_excel(op1, output_file2, columns_to_correct2)

            columns_to_correct3 = columns_to_correct + columns_to_correct2
            highlight(cleaned,op2,output_file4,columns_to_correct3)
            highlight_cells(output_file4,output_file5)
            print(f"Processed file {filename} and saved output to {output_file4}")
            