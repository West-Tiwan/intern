import pandas as pd
import requests
from bs4 import BeautifulSoup
from textblob import TextBlob
import nltk
from nltk.tokenize import sent_tokenize, word_tokenize
import re
import os

nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')

dataframe1 = pd.read_excel('Input.xlsx')

columns = ['URL_ID', 'URL', 'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE', 'SUBJECTIVE SCORE',
           'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS', 'FOG INDEX', 'AVG NUMBER OF WORDS/SENTENCE',
           'COMPLEX WORD COUNT', 'WORD COUNT', 'SYLLABLES PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH']

file_paths = ['StopWords_DatesandNumbers.txt', 'StopWords_Currencies.txt', 'StopWords_Auditor.txt',
              'StopWords_Generic.txt', 'StopWords_GenericLong.txt',
              'StopWords_Geographic.txt', 'StopWords_Names.txt']


def get_title_body_from_url(url):
    req = requests.get(url)
    parser = BeautifulSoup(req.content, "html.parser")
    title = parser.find('h1', class_='entry-title').text
    body = parser.find('div', class_='td-post-content').text
    return [title, body]


def count_syllables(word):
    word = word.lower()
    syllable_count = len(re.findall(r'[aeiouy]+', word)) - len(re.findall(r'(es|ed)$', word))
    return max(1, syllable_count)


def is_complex_word(word):
    return count_syllables(word) > 2


def analyze_text(paragraph):
    sentences = sent_tokenize(paragraph)
    words = word_tokenize(paragraph)
    word_count = len(words)
    sentence_count = len(sentences)

    positive_score = sum(1 for word in words if word.lower() in positive_words)
    negative_score = sum(1 for word in words if word.lower() in negative_words)

    blob = TextBlob(paragraph)
    polarity_score = blob.sentiment.polarity
    subjectivity_score = blob.sentiment.subjectivity

    avg_sentence_length = word_count / sentence_count
    complex_words = [word for word in words if is_complex_word(word)]
    complex_word_count = len(complex_words)
    percentage_complex_words = complex_word_count / word_count * 100
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    avg_words_per_sentence = avg_sentence_length

    syllable_count = sum(count_syllables(word) for word in words)
    syllables_per_word = syllable_count / word_count

    personal_pronouns = len(re.findall(r'\b(I|we|my|ours|us)\b', paragraph, re.IGNORECASE))
    avg_word_length = sum(len(word) for word in words) / word_count

    return {
        'POSITIVE SCORE': positive_score,
        'NEGATIVE SCORE': negative_score,
        'POLARITY SCORE': polarity_score,
        'SUBJECTIVE SCORE': subjectivity_score,
        'AVG SENTENCE LENGTH': avg_sentence_length,
        'PERCENTAGE OF COMPLEX WORDS': percentage_complex_words,
        'FOG INDEX': fog_index,
        'AVG NUMBER OF WORDS/SENTENCE': avg_words_per_sentence,
        'COMPLEX WORD COUNT': complex_word_count,
        'WORD COUNT': word_count,
        'SYLLABLES PER WORD': syllables_per_word,
        'PERSONAL PRONOUNS': personal_pronouns,
        'AVG WORD LENGTH': avg_word_length
    }


def read_words_from_files(file_paths):
    words_to_remove = set()
    for file_path in file_paths:
        with open(file_path, 'r') as file:
            for line in file:
                word = line.strip()
                if word:
                    words_to_remove.add(word)
    return words_to_remove


def remove_words(paragraph, words_to_remove):
    pattern = r'\b(?:' + '|'.join(re.escape(word) for word in words_to_remove) + r')\b'
    cleaned_paragraph = re.sub(pattern, '', paragraph)
    cleaned_paragraph = re.sub(r'\s+', ' ', cleaned_paragraph).strip()
    return cleaned_paragraph


def check_counter(curr_idx):
    counter = open("counter.txt", "r")
    if int(counter.read()) > curr_idx:
        counter.close()
        return False
    counter.close()

    counter = open("counter.txt", "w")
    counter.write(f"{curr_idx}")
    counter.close()
    return True


def read_words_from_file(file_path):
    words = set()
    with open(file_path, 'r') as file:
        for line in file:
            word = line.strip()
            if word:
                words.add(word.lower())
    return words


positive_words_file = 'positive-words.txt'  # Replace with your actual file path
negative_words_file = 'negative-words.txt'  # Replace with your actual file path

positive_words = read_words_from_file(positive_words_file)
negative_words = read_words_from_file(negative_words_file)

results = []

for i in range(len(dataframe1)):
    if not check_counter(i):
        print(f'Skipping {i}, already exist')
        continue
    print(i)

    [title, body] = get_title_body_from_url(dataframe1.URL[i])
    total = title + body
    blob = TextBlob(total)
    words_to_remove = read_words_from_files(file_paths)
    cleaned_paragraph = remove_words(str(blob), words_to_remove)
    print(cleaned_paragraph)
    analysis_results = analyze_text(cleaned_paragraph)
    result_row = {
        'URL_ID': dataframe1.URL_ID[i],
        'URL': dataframe1.URL[i],
    }
    result_row.update(analysis_results)

    results.append(result_row)
    print(results)

# Create a DataFrame from the results
results_df = pd.DataFrame(results, columns=columns)

# Write the results to an Excel file
output_file_path = 'Output Data Structure.xlsx'  # Specify the path for the output file
results_df.to_excel(output_file_path, index=False)
