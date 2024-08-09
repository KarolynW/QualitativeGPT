import pandas as pd
import re
from colorama import Fore, Style

# Define a comprehensive list of keywords related to mental health
keywords = [
    'mental health', 'suicide', 'suicidal', 'depression', 'depressed', 'anxiety', 'anxius', 'stress', 'stressed'
]

# Compile regular expressions for each keyword to match whole words
keyword_patterns = [re.compile(r'\b' + re.escape(keyword) + r'\b', re.IGNORECASE) for keyword in keywords]

def is_relevant(message):
    if not message or len(message.strip()) == 0:
        return False
    # Check for exact keyword match using regular expressions
    if any(pattern.search(message) for pattern in keyword_patterns):
        return True
    return False

# Load the Excel file
file_path = 'FF Data/FF 240324 to 230724.xlsx'

try:
    full_data = pd.read_excel(file_path, sheet_name='Query result')
    print(Fore.GREEN + "File loaded successfully." + Style.RESET_ALL)
except Exception as e:
    print(Fore.RED + f"Error loading file: {e}" + Style.RESET_ALL)

# Check the total number of rows
total_rows = len(full_data)
print(Fore.BLUE + f"Total rows in dataset: {total_rows}" + Style.RESET_ALL)

# Apply relevance filter
relevant_data_list = []
try:
    for index, row in full_data.iterrows():
        if index % 100 == 0:
            print(Fore.YELLOW + f"Processing row {index} of {total_rows}" + Style.RESET_ALL)
        if is_relevant(row['Message']):
            relevant_data_list.append(row)

    relevant_data = pd.DataFrame(relevant_data_list)
    print(Fore.GREEN + f"Data filtered by relevance. Total relevant records: {len(relevant_data)}." + Style.RESET_ALL)
except Exception as e:
    print(Fore.RED + f"Error filtering relevant data: {e}" + Style.RESET_ALL)

# Save the filtered relevant data to a new Excel file
output_file_path = 'filtered_relevant_data.xlsx'
try:
    relevant_data.to_excel(output_file_path, index=False)
    print(Fore.GREEN + f"Relevance filtering complete. The relevant data is saved as '{output_file_path}'." + Style.RESET_ALL)
except Exception as e:
    print(Fore.RED + f"Error saving file: {e}" + Style.RESET_ALL)






