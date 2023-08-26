import pandas as pd
from gensim.models import LdaModel
from gensim.corpora import Dictionary

# Load the Excel file with the extracted keywords
file_path = ''  # Replace with the actual file path
df = pd.read_excel(file_path)

# Define the number of topics for LDA
num_topics = 20  # Adjust this number based on the desired number of topics

# Tokenize the keywords and create a dictionary
keywords_list = [keywords.split() for keywords in df['Start_Keywords']]
dictionary = Dictionary(keywords_list)

# Convert the keywords to bag-of-words format
corpus = [dictionary.doc2bow(keywords) for keywords in keywords_list]

# Train the LDA model
lda_model = LdaModel(corpus=corpus, num_topics=num_topics, id2word=dictionary)

# Get the topics and their keywords
topics = lda_model.show_topics(num_topics=num_topics, num_words=5)  # Adjust num_words as needed

# Print the topics and their keywords
for topic in topics:
    print(topic)

# Create buckets based on the identified topics
buckets = {}
for i, keywords in enumerate(keywords_list):
    topic_id, _ = max(lda_model[corpus[i]], key=lambda x: x[1])  # Get the most dominant topic for each feedback
    topic_keywords = topics[topic_id][1].split('"')[1::2]  # Extract the keywords from the topic string
    for keyword in topic_keywords:
        if keyword not in buckets:
            buckets[keyword] = 1
        else:
            buckets[keyword] += 1

# Print the buckets with their counts
for keyword, count in buckets.items():
    print(f"{keyword} -- {count}")
