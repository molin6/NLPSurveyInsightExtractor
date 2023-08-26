import pandas as pd
import openai
import os

# Load the data
# Add file name & sheet name
df = pd.read_excel('file.xlsx', sheet_name='sheetname')
print(df.head())
# Define the range of rows to process

start_row = 701  # Change this to the index of the first row you want to process
end_row = 875  # Change this to the index of the last row you want to process

# Slice the dataframe to only include the rows you want to process
df = df.iloc[start_row:end_row]

# Initialize OpenAI API
# Add key
openai.api_key = ''

# Function to extract keywords
def extract_keywords(feedback):
    prompt = f"This is a piece of feedback from a student: \"{feedback}\". Please identify the main suggestion or theme in this 
    feedback and express it as a single word or short phrase."
    response = openai.Completion.create(
      engine="text-davinci-003",
      prompt=prompt,
      temperature=0.3,
      max_tokens=60
    )
    return response.choices[0].text.strip()

# Process each row and store the results in memory
df['Start_Keywords'] = df['start'].apply(extract_keywords)
df['Stop_Keywords'] = df['Stop'].apply(extract_keywords)
df['Continue_Keywords'] = df['Continue'].apply(extract_keywords)

# Create dataframes for each column's buckets
start_df = df['Start_Keywords'].value_counts().reset_index()
start_df.columns = ['Theme', 'Count']

stop_df = df['Stop_Keywords'].value_counts().reset_index()
stop_df.columns = ['Theme', 'Count']

continue_df = df['Continue_Keywords'].value_counts().reset_index()
continue_df.columns = ['Theme', 'Count']

# Save the dataframes to an Excel file
with pd.ExcelWriter('survey_keywordV1.xlsx') as writer:
    df.to_excel(writer, sheet_name='Original + Keywords')
    start_df.to_excel(writer, sheet_name='Start Buckets')
    stop_df.to_excel(writer, sheet_name='Stop Buckets')
    continue_df.to_excel(writer, sheet_name='Continue Buckets')
