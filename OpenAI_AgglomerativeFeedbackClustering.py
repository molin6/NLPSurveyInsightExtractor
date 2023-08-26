import pandas as pd
from sentence_transformers import SentenceTransformer
from sklearn.cluster import AgglomerativeClustering
import numpy as np

# Load the data
# Add file path and sheet name
df = pd.read_excel('file.xlsx', sheet_name='sheetname')
print(df.head())

# Define the range of rows to process
start_row = 0  # Change this to the index of the first row you want to process
end_row = 75  # Change this to the index of the last row you want to process

# Slice the dataframe to only include the rows you want to process
df = df.iloc[start_row:end_row]

# Initialize SentenceTransformer model
model = SentenceTransformer('paraphrase-MiniLM-L6-v2')

# Encode the feedback sentences into embeddings
feedback_sentences = df['Start'].tolist()  # Replace 'Start' with the column containing feedback
feedback_embeddings = model.encode(feedback_sentences)

# Perform clustering using AgglomerativeClustering
num_clusters = 5  # Change this to the number of clusters you want
clustering_model = AgglomerativeClustering(n_clusters=num_clusters)
clusters = clustering_model.fit_predict(feedback_embeddings)

# Add cluster information to the DataFrame
df['Cluster'] = clusters

# Save the DataFrame to an Excel file
# Add desired file name and sheet names
with pd.ExcelWriter('clustered_feedback.xlsx') as writer:
    df.to_excel(writer, sheet_name='Clustered Feedback')
