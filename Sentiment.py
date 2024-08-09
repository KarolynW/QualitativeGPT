import os
import pandas as pd
import time
from azure.core.credentials import AzureKeyCredential
from azure.ai.textanalytics import TextAnalyticsClient
from azure.core.exceptions import ServiceRequestError, ServiceResponseError
from colorama import init, Fore, Style
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Initialize colorama
init(autoreset=True)

# Authenticate the client using your key and endpoint
def authenticate_client():
    key = os.getenv("AZURE_LANGUAGE_KEY")
    endpoint = os.getenv("AZURE_LANGUAGE_ENDPOINT")
    
    if not endpoint or not key:
        print(Fore.RED + "Azure endpoint and key must be set as environment variables.")
        exit(1)
    
    ta_credential = AzureKeyCredential(key)
    client = TextAnalyticsClient(endpoint=endpoint, credential=ta_credential)
    return client

client = authenticate_client()

# Function to perform detailed sentiment analysis with retries
def analyze_sentiment(client, document):
    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = client.analyze_sentiment([document], show_opinion_mining=True)[0]
            if not response.is_error:
                return response
            else:
                return None
        except (ServiceRequestError, ServiceResponseError) as e:
            print(Fore.RED + f"Error occurred: {e}. Retrying {attempt + 1}/{max_retries}...")
            time.sleep(2 ** attempt)  # Exponential backoff
    return None

# Create directory to save individual question Excel files
output_dir = "Individual_Question_Excel"
os.makedirs(output_dir, exist_ok=True)

# Load the base Excel file
base_excel_path = 'Base Sheet.xlsx'
df = pd.read_excel(base_excel_path)

# Function to flatten the analysis results into a dictionary
def flatten_analysis(result):
    data = {
        'Overall Sentiment': result.sentiment,
        'Positive Score': result.confidence_scores.positive,
        'Neutral Score': result.confidence_scores.neutral,
        'Negative Score': result.confidence_scores.negative
    }
    for i, sentence in enumerate(result.sentences):
        sentence_key = f'Sentence {i+1}'
        data[f'{sentence_key} Text'] = sentence.text
        data[f'{sentence_key} Sentiment'] = sentence.sentiment
        data[f'{sentence_key} Positive Score'] = sentence.confidence_scores.positive
        data[f'{sentence_key} Neutral Score'] = sentence.confidence_scores.neutral
        data[f'{sentence_key} Negative Score'] = sentence.confidence_scores.negative
        for j, opinion in enumerate(sentence.mined_opinions):
            opinion_key = f'{sentence_key} Opinion {j+1}'
            data[f'{opinion_key} Target Text'] = opinion.target.text
            data[f'{opinion_key} Target Sentiment'] = opinion.target.sentiment
            data[f'{opinion_key} Target Positive Score'] = opinion.target.confidence_scores.positive
            data[f'{opinion_key} Target Negative Score'] = opinion.target.confidence_scores.negative
            for k, assessment in enumerate(opinion.assessments):
                assessment_key = f'{opinion_key} Assessment {k+1}'
                data[f'{assessment_key} Text'] = assessment.text
                data[f'{assessment_key} Sentiment'] = assessment.sentiment
                data[f'{assessment_key} Positive Score'] = assessment.confidence_scores.positive
                data[f'{assessment_key} Negative Score'] = assessment.confidence_scores.negative
    return data

# Process each question and save to individual Excel files
for column in df.columns:
    try:
        print(Fore.GREEN + f"Processing column: {column}")
        question_df = df[[column]].dropna().copy()
        question_df = question_df.rename(columns={column: 'Comment'})
        
        # Perform sentiment analysis on each comment
        results = []
        for comment in question_df['Comment']:
            analysis = analyze_sentiment(client, comment)
            if analysis:
                flattened_result = flatten_analysis(analysis)
                flattened_result['Comment'] = comment
                results.append(flattened_result)
            else:
                results.append({'Comment': comment, 'Overall Sentiment': None})

        # Convert results to DataFrame
        expanded_df = pd.DataFrame(results)
        
        # Ensure 'Comment' is the first column
        columns = ['Comment'] + [col for col in expanded_df.columns if col != 'Comment']
        expanded_df = expanded_df[columns]
        
        # Save to individual Excel file
        sanitized_column = column[:4]  # Use the first four characters of the column header
        question_file_path = os.path.join(output_dir, f"{sanitized_column}.xlsx")
        expanded_df.to_excel(question_file_path, index=False)

        print(Fore.BLUE + f"Processed and saved: {question_file_path}")

    except Exception as e:
        print(Fore.RED + f"Failed to process column '{column}': {e}")
        continue

print(Fore.GREEN + "All columns processed and saved.")





