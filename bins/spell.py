import os
from groq import Groq
from dotenv import load_dotenv
import pandas as pd
import re

# Load environment variables from .env file
load_dotenv(override=True)

# Retrieve the API key from environment variables
api_key = os.environ.get("GROQ_API_KEY")
if not api_key:
    raise ValueError("GROQ_API_KEY environment variable not set")

# Function to convert Excel column index (e.g., 'A', 'B', 'C') to numeric index
def excel_to_numeric_index(excel_index):
    if isinstance(excel_index, str) and re.match(r'[A-Za-z]+', excel_index):
        excel_index = excel_index.upper()
        # Convert Excel column letter to numeric index (e.g., 'A' -> 0, 'B' -> 1, etc.)
        index = 0
        for char in excel_index:
            index = index * 26 + (ord(char) - ord('A')) + 1
        return index - 1  # Adjust for zero-based indexing in Python
    else:
        raise ValueError(f"Invalid Excel column index: {excel_index}")

# Function to create general spelling correction prompt
def create_prompt(text):
    return f"Correct the following text and give the output in maximum 3 words only,{text} if you not found any values then stop the model"

# Function to correct text using Groq API
def correct_text(text):
    try:
        client = Groq(api_key=api_key)
        chat_completion = client.chat.completions.create(
            messages=[
                {
                    "role": "user",
                    "content": create_prompt(text),
                }
            ],
            model="gemma2-9b-it",
            temperature=0.4,
            max_tokens=20,
            top_p=0.7,
            seed=10,
        )
        corrected_text = chat_completion.choices[0].message.content.strip()
        return corrected_text
    except Exception as e:
        print(f"Error: {e}")
        return text


# Function to load Excel file, apply corrections, and save updated DataFrame
def process_excel(input_file, output_file, columns_to_correct):
    try:
        # Load the Excel file
        df = pd.read_excel(input_file)

        # Apply corrections to specified columns
        for excel_index in columns_to_correct:
            col_index = excel_to_numeric_index(excel_index)
            col_name = df.columns[col_index]
            df[col_name] = df[col_name].apply(lambda x: correct_text(x) if isinstance(x, str) else x)  # Spell check only if string

        # Save the updated DataFrame to a new Excel file
        df.to_excel(output_file, index=False)
        print(f"Processed data saved to {output_file}")
        return output_file
    except Exception as e:
        print(f"Error processing Excel: {e}")
        return None
