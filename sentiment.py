import pandas as pd
import requests
from bs4 import BeautifulSoup
import os

# Read the Excel file
file_path = 'Input.xlsx'
df = pd.read_excel(file_path)

# Ensure the output directory exists
output_dir = 'extracted_articles'
os.makedirs(output_dir, exist_ok=True)

def extract_article_content(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Extract the title
            title_tag = soup.find('h1', class_='entry-title')
            title = title_tag.get_text() if title_tag else 'No Title Found'
            
            # Extract the article content
            article_body = soup.find('div', class_='td-post-content')
            if article_body:
                paragraphs = article_body.find_all('p')
                article_text = '\n'.join([para.get_text() for para in paragraphs])
                return title, article_text
            else:
                return None, None
        return None, None
    except Exception as e:
        print(f"Error fetching URL {url}: {e}")
        return None, None

# Process each URL
for index, row in df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    
    title, article_text = extract_article_content(url)
    if title and article_text:
        file_name = f"{output_dir}/{url_id}.txt"
        with open(file_name, 'w', encoding='utf-8') as file:
            file.write(f"{title}\n\n{article_text}")
        print(f"Saved: {file_name}")
    else:
        print(f"Failed to extract content for URL_ID: {url_id}, URL: {url}")





import os
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize
import re
import textstat
import string

# Ensure NLTK resources are downloaded
nltk.download('punkt')
nltk.download('stopwords')

# Load stop words from NLTK
nltk_stop_words = set(stopwords.words('english'))

# Load custom stop words and dictionaries
def load_stop_words(directory):
    stop_words = set()
    for file_name in os.listdir(directory):
        file_path = os.path.join(directory, file_name)
        with open(file_path, 'r') as file:
            stop_words.update(line.strip().lower() for line in file)
    return stop_words

def load_dictionary(dictionary_file):
    with open(dictionary_file, 'r') as file:
        return set(file.read().splitlines())

stop_words_directory = 'StopWords'
stop_words_list = load_stop_words(stop_words_directory).union(nltk_stop_words)
positive_dict = load_dictionary('MasterDictionary/positive-words.txt')
negative_dict = load_dictionary('MasterDictionary/negative-words.txt')

def clean_text(text):
    tokens = word_tokenize(text.lower())
    cleaned_tokens = [word for word in tokens if word.isalpha() and word not in stop_words_list]
    return cleaned_tokens

def calculate_scores(text):
    tokens = clean_text(text)
    positive_score = sum(1 for word in tokens if word in positive_dict)
    negative_score = sum(-1 for word in tokens if word in negative_dict)
    
    # Polarity Score Calculation
    total_score = positive_score + abs(negative_score)
    polarity_score = (positive_score - abs(negative_score)) / (total_score + 0.000001)
    
    # Subjectivity Score Calculation
    subjectivity_score = total_score / (len(tokens) + 0.000001)
    
    return positive_score, abs(negative_score), polarity_score, subjectivity_score
    
def average_sentence_length(text):
    sentences = len(sent_tokenize(text))
    words = len(word_tokenize(text))
    return words / sentences if sentences > 0 else 0

def percentage_complex_words(text):
    words = word_tokenize(text)
    complex_words = [word for word in words if textstat.syllable_count(word) > 2]
    return 100* (len(complex_words) / len(words)) if len(words) > 0 else 0

def fog_index(text):
    avg_sentence_length = average_sentence_length(text)
    percent_complex_words = percentage_complex_words(text)
    return 0.4 * (avg_sentence_length + percent_complex_words)

def average_number_of_words_per_sentence(text):
    sentences = len(sent_tokenize(text))  # Tokenize the text into sentences
    words = len(word_tokenize(text))      # Tokenize the text into words
    return words / sentences if sentences > 0 else 0


def complex_word_count(text):
    words = word_tokenize(text)
    return sum(1 for word in words if textstat.syllable_count(word) > 2)


def word_count(text):
    tokens = word_tokenize(text.lower())
    filtered_tokens = [word for word in tokens if word.isalpha() and word not in stop_words_list]
    return len(filtered_tokens)

def syllable_count(word):
    word = word.lower()
    
    # Handle exceptions
    if word.endswith('es') or word.endswith('ed'):
        word = word[:-2]
    
    # Count vowels (a, e, i, o, u) and handle common vowel patterns
    syllables = len(re.findall(r'[aeiouy]+', word))
    
    # Adjust for common syllable counting rules
    if word.endswith('e'):
        syllables -= 1
    if syllables == 0:
        syllables = 1  # Ensure at least one syllable for non-empty words
    
    return syllables

def syllable_count_per_word(text):
    words = [word for word in word_tokenize(text) if word.isalpha()]
    syllable_counts = [syllable_count(word) for word in words]
    if len(words) > 0:
        return sum(syllable_counts) / len(words)
    else:
        return 0




def extract_personal_pronouns(text):
    # Regular expression to find personal pronouns
    pronouns = re.findall(r'\b(I|we|my|ours|us)\b', text, re.IGNORECASE)
    
    # Remove occurrences of "US" that are likely country names
    filtered_pronouns = [p for p in pronouns if p.lower() != 'us']
    
    return len(filtered_pronouns)



def average_word_length(text):
    words = word_tokenize(text)
    total_characters = sum(len(word) for word in words)
    return total_characters / len(words) if len(words) > 0 else 0


# Function to read text from a .txt file
def read_text_from_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read()





# Directory containing the .txt files
text_files_directory = 'extracted_articles'
# Load the existing Excel file
output_file_path = 'output.xlsx'
wb = load_workbook(output_file_path)
ws = wb.active

# Iterate through all .txt files in the directory
for file_name in os.listdir(text_files_directory):
    if file_name.endswith('.txt'):
        text_file_path = os.path.join(text_files_directory, file_name)

        # Read the text from the file
        text = read_text_from_file(text_file_path)

        # Sentiment Analysis
        positive_score, negative_score, polarity_score, subjectivity_score = calculate_scores(text)

        # Readability Analysis
        avg_sentence_length = average_sentence_length(text)
        percentage_complex = percentage_complex_words(text)
        fog_idx = fog_index(text)
        avg_words_per_sentence = average_number_of_words_per_sentence(text)
        complex_word_cnt = complex_word_count(text)
        word_cnt = word_count(text)
        syllables_per_word = syllable_count_per_word(text)
        avg_word_length = average_word_length(text)

        # Personal Pronouns
        pronouns_count = extract_personal_pronouns(text)

        # Assuming 'url_id' is the filename without extension
        url_id = os.path.splitext(file_name)[0]

        # Find the row index for the given url_id
        row_index = None
        for row in ws.iter_rows(min_row=2):
            if row[0].value == url_id:
                row_index = row[0].row
                break

        # If the row is found, update the corresponding columns
        if row_index:
            ws[f"C{row_index}"] = positive_score
            ws[f"D{row_index}"] = negative_score
            ws[f"E{row_index}"] = polarity_score
            ws[f"F{row_index}"] = subjectivity_score
            ws[f"G{row_index}"] = avg_sentence_length
            ws[f"H{row_index}"] = percentage_complex
            ws[f"I{row_index}"] = fog_idx
            ws[f"J{row_index}"] = avg_words_per_sentence
            ws[f"K{row_index}"] = complex_word_cnt
            ws[f"L{row_index}"] = word_cnt
            ws[f"M{row_index}"] = syllables_per_word
            ws[f"N{row_index}"] = pronouns_count
            ws[f"O{row_index}"] = avg_word_length

# Save the updated Excel file
wb.save(output_file_path)
print("Results have been saved to output.xlsx")