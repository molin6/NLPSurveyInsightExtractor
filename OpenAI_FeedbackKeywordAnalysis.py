import openai
import pandas as pd
from collections import Counter

#openai key here
openai.api_key = ''

def extract_keywords(feedback):
    prompt = f"This is a piece of feedback from a student: \"{feedback}\". Please identify the main suggestion or theme in this feedback."
    response = openai.Completion.create(
      engine="text-davinci-003",
      prompt=prompt,
      temperature=0.3,
      max_tokens=60
    )
    return response.choices[0].text.strip()

# Load the data
# add path
file_path = ""
#add sheet name
sheet_name = ""
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Apply the function to the 'Start', 'Stop', and 'Continue' columns
df['Start_Keywords'] = df.iloc[:, 1].apply(extract_keywords)
df['Stop_Keywords'] = df.iloc[:, 2].apply(extract_keywords)
df['Continue_Keywords'] = df.iloc[:, 0].apply(extract_keywords)

# Count the occurrences of each keyword
start_counts = Counter(df['Start_Keywords'])
stop_counts = Counter(df['Stop_Keywords'])
continue_counts = Counter(df['Continue_Keywords'])

# Convert the count dictionaries to DataFrames
start_df = pd.DataFrame.from_dict(start_counts, orient='index', columns=['Count'])
stop_df = pd.DataFrame.from_dict(stop_counts, orient='index', columns=['Count'])
continue_df = pd.DataFrame.from_dict(continue_counts, orient='index', columns=['Count'])

# Save the results
with pd.ExcelWriter('keyword_analysis.xlsx') as writer:
    df.to_excel(writer, sheet_name='Original + Keywords')
    start_df.to_excel(writer, sheet_name='Start Buckets')
    stop_df.to_excel(writer, sheet_name='Stop Buckets')
    continue_df.to_excel(writer, sheet_name='Continue Buckets')
