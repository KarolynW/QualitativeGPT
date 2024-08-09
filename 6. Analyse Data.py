import pandas as pd
from openai import OpenAI
from colorama import Fore, Style
import json
import time

# Initialize OpenAI API client
client = OpenAI()

def analyze_message(combined_message, retries=3):
    wait_time = 5
    for attempt in range(retries):
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {
                        "role": "system",
                        "content": """
Conduct the following analysis on the forum conversation:

- Thematic Analysis
- Conversational Analysis
- Discourse Analysis

The aim of this research is:
The aim of this study is to explore the mental health issues experienced by farmers as expressed in online forum discussions. The study seeks to identify key mental health concerns, understand the context in which they arise.

Return the results in JSON using this structure without "```json" at the beginning or end:

{
  "Summary": {
    "Summary": "[SHORT SUMMARY OF THREAD]"
  },
  "ThematicAnalysis": {
    "Theme 1": "[OBSERVATION]",
    "Theme 1 Explanation": "[OBSERVATION]",
    "Theme 1 Occurrence Count": [NUMBER],
    "Theme 2": "[OBSERVATION]",
    "Theme 2 Explanation": "[OBSERVATION]",
    "Theme 2 Occurrence Count": [NUMBER],
    "Theme 3": "[OBSERVATION]",
    "Theme 3 Explanation": "[OBSERVATION]",
    "Theme 3 Occurrence Count": [NUMBER],
    "Theme 4": "[OBSERVATION]",
    "Theme 4 Explanation": "[OBSERVATION]",
    "Theme 4 Occurrence Count": [NUMBER],
    "Theme 5": "[OBSERVATION]",
    "Theme 5 Explanation": "[OBSERVATION]",
    "Theme 5 Occurrence Count": [NUMBER]
  },
  "ConversationalAnalysis": {
    "TurnTaking": {
      "Observation": "[OBSERVATION]",
      "Significance": "[OBSERVATION]"
    },
    "Repairs": {
      "Observation": "[OBSERVATION]",
      "Significance": "[OBSERVATION]"
    },
    "AdjacencyPairs": {
      "Observation": "[OBSERVATION]",
      "Significance": "[OBSERVATION]"
    }
  },
  "DiscourseAnalysis": {
    "LanguageUse": {
      "Observation": "[OBSERVATION]",
      "Significance": "[OBSERVATION]"
    },
    "NarrativeStructure": {
      "Observation": "[OBSERVATION]",
      "Significance": "[OBSERVATION]"
    },
    "Identity and Roles": {
      "Observation": "[OBSERVATION]",
      "Significance": "[OBSERVATION]"
    },
    "Intertextuality": {
      "Observation": "[OBSERVATION]",
      "Significance": "[OBSERVATION]"
    }
  }
}
"""
                    },
                    {
                        "role": "user",
                        "content": combined_message
                    }
                ],
                temperature=0,
                max_tokens=4095,
                top_p=1,
                frequency_penalty=0,
                presence_penalty=0
            )
            message = response.choices[0].message.content
            return message
        except Exception as e:
            print(Fore.RED + f"Error processing message (attempt {attempt + 1}): {e}" + Style.RESET_ALL)
            time.sleep(wait_time)
            wait_time += 5
    return '{}'

# Load the updated dataset with section
file_path = 'updated_dataset_with_section.xlsx'
output_file_path = 'analyzed_dataset_with_section.xlsx'

try:
    data = pd.read_excel(file_path)
    print(Fore.GREEN + "File loaded successfully." + Style.RESET_ALL)
except Exception as e:
    print(Fore.RED + f"Error loading file: {e}" + Style.RESET_ALL)

# Initialize columns for JSON keys if not present
json_keys = [
    "Summary",
    "Theme 1", "Theme 1 Explanation", "Theme 1 Occurrence Count",
    "Theme 2", "Theme 2 Explanation", "Theme 2 Occurrence Count",
    "Theme 3", "Theme 3 Explanation", "Theme 3 Occurrence Count",
    "Theme 4", "Theme 4 Explanation", "Theme 4 Occurrence Count",
    "Theme 5", "Theme 5 Explanation", "Theme 5 Occurrence Count",
    "TurnTaking Observation", "TurnTaking Significance",
    "Repairs Observation", "Repairs Significance",
    "AdjacencyPairs Observation", "AdjacencyPairs Significance",
    "LanguageUse Observation", "LanguageUse Significance",
    "NarrativeStructure Observation", "NarrativeStructure Significance",
    "Identity and Roles Observation", "Identity and Roles Significance",
    "Intertextuality Observation", "Intertextuality Significance"
]

for key in json_keys:
    if key not in data.columns:
        data[key] = None

# Analyze each combined message and store the JSON response in the appropriate columns
total_rows = len(data)
print(Fore.BLUE + f"Total rows to process: {total_rows}" + Style.RESET_ALL)

for index, row in data.iterrows():
    if pd.isna(row['Summary']):
        if index % 10 == 0:
            print(Fore.YELLOW + f"Processing row {index + 1} of {total_rows}" + Style.RESET_ALL)
        json_response = analyze_message(row['Combined Messages'])
        print(Fore.CYAN + f"Row {index + 1} JSON Response: {json_response}" + Style.RESET_ALL)
        try:
            response_dict = json.loads(json_response)
            data.at[index, 'Summary'] = response_dict.get('Summary', {}).get('Summary', None)
            for i in range(1, 6):
                data.at[index, f'Theme {i}'] = response_dict.get('ThematicAnalysis', {}).get(f'Theme {i}', None)
                data.at[index, f'Theme {i} Explanation'] = response_dict.get('ThematicAnalysis', {}).get(f'Theme {i} Explanation', None)
                data.at[index, f'Theme {i} Occurrence Count'] = response_dict.get('ThematicAnalysis', {}).get(f'Theme {i} Occurrence Count', None)
            data.at[index, 'TurnTaking Observation'] = response_dict.get('ConversationalAnalysis', {}).get('TurnTaking', {}).get('Observation', None)
            data.at[index, 'TurnTaking Significance'] = response_dict.get('ConversationalAnalysis', {}).get('TurnTaking', {}).get('Significance', None)
            data.at[index, 'Repairs Observation'] = response_dict.get('ConversationalAnalysis', {}).get('Repairs', {}).get('Observation', None)
            data.at[index, 'Repairs Significance'] = response_dict.get('ConversationalAnalysis', {}).get('Repairs', {}).get('Significance', None)
            data.at[index, 'AdjacencyPairs Observation'] = response_dict.get('ConversationalAnalysis', {}).get('AdjacencyPairs', {}).get('Observation', None)
            data.at[index, 'AdjacencyPairs Significance'] = response_dict.get('ConversationalAnalysis', {}).get('AdjacencyPairs', {}).get('Significance', None)
            data.at[index, 'LanguageUse Observation'] = response_dict.get('DiscourseAnalysis', {}).get('LanguageUse', {}).get('Observation', None)
            data.at[index, 'LanguageUse Significance'] = response_dict.get('DiscourseAnalysis', {}).get('LanguageUse', {}).get('Significance', None)
            data.at[index, 'NarrativeStructure Observation'] = response_dict.get('DiscourseAnalysis', {}).get('NarrativeStructure', {}).get('Observation', None)
            data.at[index, 'NarrativeStructure Significance'] = response_dict.get('DiscourseAnalysis', {}).get('NarrativeStructure', {}).get('Significance', None)
            data.at[index, 'Identity and Roles Observation'] = response_dict.get('DiscourseAnalysis', {}).get('Identity and Roles', {}).get('Observation', None)
            data.at[index, 'Identity and Roles Significance'] = response_dict.get('DiscourseAnalysis', {}).get('Identity and Roles', {}).get('Significance', None)
            data.at[index, 'Intertextuality Observation'] = response_dict.get('DiscourseAnalysis', {}).get('Intertextuality', {}).get('Observation', None)
            data.at[index, 'Intertextuality Significance'] = response_dict.get('DiscourseAnalysis', {}).get('Intertextuality', {}).get('Significance', None)
        except json.JSONDecodeError as e:
            print(Fore.RED + f"Error parsing JSON response: {e}" + Style.RESET_ALL)

        # Save the updated data to a new Excel file after every row
        try:
            data.to_excel(output_file_path, index=False)
            print(Fore.GREEN + f"Row {index + 1} processed and saved." + Style.RESET_ALL)
        except Exception as e:
            print(Fore.RED + f"Error saving file: {e}" + Style.RESET_ALL)

print(Fore.GREEN + "AI analysis complete. The updated data is saved." + Style.RESET_ALL)







