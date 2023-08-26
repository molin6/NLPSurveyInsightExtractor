import nltk
nltk.download('punkt')
nltk.download('stopwords')
nltk.download('wordnet')

import pandas as pd
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer
from nltk.tokenize import word_tokenize
import string
from gensim.corpora import Dictionary
from gensim.models import LdaModel

#Load the data
#add path
file_path = "put path here"
sheet_name = "fall feed back"
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Initialize a lemmatizer
lemmatizer = WordNetLemmatizer()

# Define a function to clean the text
def clean_text(text):
    # Check if the input is a string
    if isinstance(text, str):
        text = text.lower()
        text = ''.join([word for word in text if word not in string.punctuation])
        words = word_tokenize(text)
        words = [lemmatizer.lemmatize(word) for word in words if word not in stopwords.words('english')]
        return words
    else:
        # If the input is not a string, return an empty list
        return []


# Apply the function to the 'Start', 'Stop', and 'Continue' columns
df['Start'] = df.iloc[:, 1].apply(clean_text)
df['Stop'] = df.iloc[:, 2].apply(clean_text)
df['Continue'] = df.iloc[:, 0].apply(clean_text)

# Create a list of all the words in the feedback
words = [word for sublist in df['Start'].tolist() for word in sublist]

# Create a Dictionary from the words
dictionary = Dictionary([words])

# Create a Corpus from the words
corpus = [dictionary.doc2bow(text) for text in [words]]

# Train the LDA model
lda_model = LdaModel(corpus=corpus, id2word=dictionary, num_topics=10)  # Choose an appropriate number of topics

# Print the most significant words for each topic
topics = lda_model.print_topics()
for topic in topics:
    print(topic)
