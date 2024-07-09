import pandas as pd
import numpy as np

def clean_data(input_file, cleaned_file):
    try:
        # Load the Excel file
        data = pd.read_excel(input_file, header=None)

        # Check if the header contains numbers
        if any(char.isdigit() for col in data.iloc[0] for char in str(col)):
            # Remove the first row
            data = data.iloc[1:]

        # Check if the header contains numbers and characters
        if any(char.isdigit() for col in data.iloc[0] for char in str(col)):
            # Remove the first row
            data = data.iloc[2:]

        # Set the header (use the second row as the header)
        header = data.iloc[0]  # Corrected to use the first row as the header
        data = data.iloc[1:]   # Corrected to start from the second row

        # Set the header
        data.columns = header

        # Remove empty rows (rows with all NaN values)
        data_cleaned = data.dropna(how='all')

        # Identify and remove rows that contain only one non-null value
        rows_to_remove = data_cleaned.apply(lambda x: x.count() == 1, axis=1)
        data_cleaned = data_cleaned[~rows_to_remove]

        # Save the cleaned data to an Excel file
        data_cleaned.to_excel(cleaned_file, index=False)
        print(f"Data cleaned and saved to {cleaned_file}")
        return cleaned_file

    except Exception as e:
        print(f"Error cleaning data: {e}")
        return None
