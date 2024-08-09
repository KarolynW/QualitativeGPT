import pandas as pd
from openai import OpenAI
from colorama import Fore, Style
import json

# Initialize OpenAI API client
client = OpenAI()

def check_message(message):
    try:
        completion = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are a filtering quality control agent.\n\n"
                        "You will be presented with the contents of a forum discussion.\n\n"
                        "You must determine if this thread is worth retaining for further analysis. \n\n"
                        "The project's current aim is to investigate issues surrounding mental health in farmers by analysing forum data.\n\n"
                        "You must return your response in JSON format like this:\n\n"
                        "{\n  \"retain\": [true or false],\n  \"reason\": [give the reason]\n}"
                    )
                },
                {
                    "role": "user",
                    "content": message
                }
            ],
            temperature=0,
            max_tokens=1024,
            top_p=1,
            frequency_penalty=0,
            presence_penalty=0
        )
        print(Fore.GREEN + f"Message processed: {completion.choices[0].message.content}" + Style.RESET_ALL)
        return completion.choices[0].message.content
    except Exception as e:
        print(Fore.RED + f"Error processing message: {e}" + Style.RESET_ALL)
        return '{"retain": false, "reason": "Error processing message"}'

# Load the filtered data
file_path = 'filtered_relevant_data.xlsx'
output_file_path = 'filtered_relevant_data_with_ai_responses.xlsx'

try:
    filtered_data = pd.read_excel(file_path)
    print(Fore.GREEN + "File loaded successfully." + Style.RESET_ALL)
except Exception as e:
    print(Fore.RED + f"Error loading file: {e}" + Style.RESET_ALL)

# Check each message and store the AI's response
total_rows = len(filtered_data)
print(Fore.BLUE + f"Total rows to process: {total_rows}" + Style.RESET_ALL)

# Initialize columns for AI responses if not present
if 'Retain' not in filtered_data.columns:
    filtered_data['Retain'] = None
if 'Reason' not in filtered_data.columns:
    filtered_data['Reason'] = None

for index, row in filtered_data.iterrows():
    if pd.isna(row['Retain']):
        if index % 10 == 0:
            print(Fore.YELLOW + f"Processing row {index + 1} of {total_rows}" + Style.RESET_ALL)
        ai_response = check_message(row['Message'])
        try:
            response_json = json.loads(ai_response)
            filtered_data.at[index, 'Retain'] = response_json.get('retain')
            filtered_data.at[index, 'Reason'] = response_json.get('reason')
        except json.JSONDecodeError as e:
            print(Fore.RED + f"Error parsing JSON response: {e}" + Style.RESET_ALL)
            filtered_data.at[index, 'Retain'] = False
            filtered_data.at[index, 'Reason'] = "Error parsing JSON response"

        # Save the updated data to a new Excel file after every row
        try:
            filtered_data.to_excel(output_file_path, index=False)
            print(Fore.GREEN + f"Row {index + 1} processed and saved." + Style.RESET_ALL)
        except Exception as e:
            print(Fore.RED + f"Error saving file: {e}" + Style.RESET_ALL)

print(Fore.GREEN + "AI filtering complete. The updated data is saved." + Style.RESET_ALL)



