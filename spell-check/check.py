import os
from groq import Groq
from dotenv import load_dotenv
import pandas as pd

# Load environment variables from .env file
load_dotenv(override=True)

# Retrieve the API key from environment variables
api_key = os.environ.get("GROQ_API_KEY")
if not api_key:
    raise ValueError("GROQ_API_KEY environment variable not set")

# Load the Excel file
df = pd.read_excel(r'data/validated_data.xlsx')

def correct(text):
    try:
        client = Groq(api_key=api_key)
        chat_completion = client.chat.completions.create(
            messages=[
                {
                    "role": "user",
                    "content": f"Correct the following text and give the output in maximum 3 words only: {text}",
                }
            ],
            model="llama3-70b-8192",
            temperature=0.4,
            max_tokens=100,
            top_p=0.7,
        )
        corrected_text = chat_completion.choices[0].message.content
        return corrected_text
    except Exception as e:
        print(f"Error: {e}")
        return text

def crtbrand(text):
    try:
        client = Groq(api_key=api_key)
        chat_completion = client.chat.completions.create(
            messages=[
                {
                    "role": "user",
                    "content": f"Check the following brand name and return the output in double quotes without any messages, and if any error print 'value not found': {text}",
                }
            ],
            model="llama3-8b-8192",
        )
        corrected_text = chat_completion.choices[0].message.content
        return corrected_text
    except Exception as e:
        print(f"Error: {e}")
        return "value not found"

column_to_correct = 'Equipment'
column_to_correct2 = 'Brand_Model'

# Apply the correction function to the specified columns
df[column_to_correct] = df[column_to_correct].apply(correct)
df[column_to_correct2] = df[column_to_correct2].apply(crtbrand)

# Save the updated DataFrame to a new Excel file
df.to_excel(r'data/output/Final_output.xlsx', index=False)
