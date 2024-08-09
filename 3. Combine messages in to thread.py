import pandas as pd
from colorama import Fore, Style

# Load the data
file_path = 'filtered prescreening data/filtered_relevant_data_with_ai_responses.xlsx'
output_file_path = 'combined_thread_messages.xlsx'

try:
    data = pd.read_excel(file_path)
    print(Fore.GREEN + "File loaded successfully." + Style.RESET_ALL)
except Exception as e:
    print(Fore.RED + f"Error loading file: {e}" + Style.RESET_ALL)

# Combine messages per thread
combined_data = data.groupby('Thread ID').apply(
    lambda x: pd.Series({
        'Thread Title': x['Thread Title'].iloc[0],
        'Combined Messages': ' '.join([f"[SPEAKER: {row['Username']}] {row['Message']}" for index, row in x.iterrows()])
    })
).reset_index()

# Save the combined data to a new Excel file
try:
    combined_data.to_excel(output_file_path, index=False)
    print(Fore.GREEN + f"Combined data saved successfully to '{output_file_path}'." + Style.RESET_ALL)
except Exception as e:
    print(Fore.RED + f"Error saving file: {e}" + Style.RESET_ALL)

