import pandas as pd

# Load the Excel file
file_path = 'P:\\MRO Items\\CORPORATE RELIABILITY ENGINEERING\\JonathanMa\\Documents\\Prodac_AZ-Tape_Downtime_Report.xlsx'
df = pd.read_excel(file_path)

# Display the first few rows of the dataframe to understand its structure
# print(df.head())

from collections import Counter
import re

# Extract the comments column
comments = df['Comment'].dropna()

# Function to clean and tokenize the comments
def tokenize_comment(comment):
    # Remove punctuation and convert to lowercase
    comment = re.sub(r'[^\w\s]', '', comment).lower()
    # Split into words
    words = comment.split()
    return words

# Tokenize all comments
all_words = []
for comment in comments:
    all_words.extend(tokenize_comment(comment))

# Count the frequency of each word
word_freq = Counter(all_words)

# Display the most common words
print(word_freq.most_common(200))
