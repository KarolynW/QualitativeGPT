import pandas as pd
from colorama import Fore, Style

# Load the data from the two spreadsheets
file_path_1 = 'DataSet.xlsx'
file_path_2 = 'filtered prescreening data/filtered_relevant_data_with_ai_responses.xlsx'
output_file_path = 'updated_dataset_with_section.xlsx'

try:
    dataset_1 = pd.read_excel(file_path_1)
    dataset_2 = pd.read_excel(file_path_2)
    print(Fore.GREEN + "Files loaded successfully." + Style.RESET_ALL)
except Exception as e:
    print(Fore.RED + f"Error loading files: {e}" + Style.RESET_ALL)

# Create a dictionary to map thread titles to sections from the second dataset
thread_to_section = dict(zip(dataset_2['Thread Title'], dataset_2['Section Title']))

# Add a new column 'Section' to the first dataset based on the thread titles
dataset_1['Section'] = dataset_1['Thread Title'].map(thread_to_section)

# Save the updated first dataset to a new Excel file
try:
    dataset_1.to_excel(output_file_path, index=False)
    print(Fore.GREEN + f"Updated dataset saved successfully to '{output_file_path}'." + Style.RESET_ALL)
except Exception as e:
    print(Fore.RED + f"Error saving file: {e}" + Style.RESET_ALL)
