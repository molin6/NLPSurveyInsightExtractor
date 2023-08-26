import pandas as pd
from sentence_transformers import SentenceTransformer
from sklearn.cluster import KMeans
import numpy as np

# Replace "path_to_your_excel_file.xlsx" with the actual file path of your Excel file
survey_file_path = "file path"
df = pd.read_excel(survey_file_path)

# Specify the column name that contains the survey responses
response_column = ""

# Load the pre-trained BERT model for sentence embeddings
model_name = "bert-base-uncased"
model = SentenceTransformer(model_name)

# Generate sentence embeddings for each response
sentence_embeddings = model.encode(df[response_column].values.tolist(), show_progress_bar=True)

# Number of clusters (buckets) to create
num_clusters = 5  # You can adjust this number based on your needs

# Apply K-means clustering to sentence embeddings
kmeans = KMeans(n_clusters=num_clusters, random_state=42)
clusters = kmeans.fit_predict(sentence_embeddings)

# Step 5: Create bucket summaries with representative examples
# Create a new DataFrame to store bucket summaries
bucket_summary = pd.DataFrame(columns=["Bucket", "Count", "Examples"])

# Assign each response to its respective cluster (bucket)
df["Bucket"] = clusters

# Calculate the count and representative examples for each bucket
for i in range(num_clusters):
    bucket_responses = df[df["Bucket"] == i][response_column].tolist()
    count = len(bucket_responses)
    examples = ", ".join(bucket_responses[:min(5, count)])
    bucket_summary = bucket_summary.append({"Bucket": i, "Count": count, "Examples": examples}, ignore_index=True)

# Print the bucket summaries
print(bucket_summary)
