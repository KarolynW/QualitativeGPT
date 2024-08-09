import os
import pandas as pd
from openai import OpenAI
from colorama import Fore, Style, init

# Initialize colorama
init(autoreset=True)

# Initialize the OpenAI client
client = OpenAI()

def process_spreadsheet(file_path, output_path):
    print(Fore.BLUE + f"Processing file: {file_path}")
    
    system_instruction = """
                    You will receive a comment extracted from a customer survey along with the result of analysis and a predefined list of themes.

                    Your task is to analyse the content of the comment to determine the presence of these themes.

                    For each theme in the list, you must return a JSON array where you replace 'ThemeX' with the actual theme name. Assign a value of 1 if the theme is relevant to the comment, and a 0 if it is absent.

                    Ensure the analysis is accurate and the JSON format is correctly structured.
                    
                    Do not preceede the JSON output with "```json" or "```" after as it will cause an error.

                    Example of expected JSON output using placeholders:

                    { 'Theme1': 1, 'Theme2': 0, 'Theme3': 1, 'Theme4': 0 }

                    In your actual output:

                        Replace 'Theme1', 'Theme2', etc., with the actual theme names provided.

                        The numbers represent whether each theme is present (1) or not present (0) in the comment.
                    
                    """
    
    try:
        # Load the spreadsheet
        data = pd.read_excel(file_path)
        
        # Extract themes assuming they start from the third column
        themes = data.columns[3:]
        print(Fore.GREEN + "Themes identified: " + ", ".join(themes))
        
        # Iterate over each row in the DataFrame
        for index, row in data.iterrows():
            # Construct the user content for the API, safely encoding the comment
            user_content = f"""
            comment: {repr(row['Combined Messages'])}
            themes: {', '.join(themes)}
            """.strip()
            print(Fore.CYAN + f"Processing row {index + 1} with content: {user_content}")
            
            # API call to OpenAI
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": system_instruction},
                    {"role": "user", "content": user_content}
                ],
                temperature=0,
                top_p=1,
                frequency_penalty=0,
                presence_penalty=0
            )
            
            print(response.choices[0].message.content)
            
            # Parse the response and update the DataFrame
            response_data = eval(response.choices[0].message.content)
            for theme in themes:
                data.at[index, theme] = response_data.get(theme, 0)  # Use .get to handle missing themes gracefully

        # Save the updated DataFrame
        data.to_excel(output_path, index=False)
        print(Fore.GREEN + f"File saved successfully: {output_path}")
    
    except Exception as e:
        print(Fore.RED + f"Error processing file {file_path}: {e}")

def process_all_spreadsheets(input_dir, output_dir):
    for filename in os.listdir(input_dir):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(input_dir, filename)
            output_path = os.path.join(output_dir, filename)
            
            # Check if the output file already exists
            if not os.path.exists(output_path):
                process_spreadsheet(file_path, output_path)
            else:
                print(Fore.YELLOW + f"Output file already exists, skipping: {output_path}")
        else:
            print(Fore.YELLOW + f"Skipped non-Excel file: {filename}")

# Example usage
input_directory = 'thematic analysis'
output_directory = 'thematic analysis/output'
process_all_spreadsheets(input_directory, output_directory)


