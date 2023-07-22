from nltk.corpus import stopwords
import string
import pandas as pd
import requests
from bs4 import BeautifulSoup
from urllib.request import urlopen
import nltk
import pyphen
import re
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


# read the Excel file into a pandas dataframe
df = pd.read_excel('Output Data Structure.xlsx')

# access the column containing the URLs
urls = df['URL']

POSITIVE_SCORE = []
NEGATIVE_SCORE = []
POLARITY_SCORE = []
SUBJECTIVITY_SCORE = []
AVG_SENTENCE_LENGTH = []
PERCENTAGE_OF_COMPLEX_WORDS = []
FOG_INDEX = []
AVG_NUMBER_OF_WORDS_PER_SENTENCE = []
COMPLEX_WORD_COUNT = []
WORD_COUNT = []
AVG_SYLLABLE_PER_WORD = []
PERSONAL_PRONOUNS = []
AVG_WORD_LENGTH = []
URL = []
URL_ID = []

# iterate over the URLs and implement the program for each URL
url_id = 37
for url in urls:
    response = requests.get(url)

    if response.status_code == 200:

        try:

            html = urlopen(url).read()

            soup = BeautifulSoup(html, features="html.parser")

            # kill all script and style elements
            for script in soup(["script", "style"]):
                script.extract()    # rip it out

            # get text
            text = soup.get_text()

            # break into lines and remove leading and trailing space on each
            lines = (line.strip() for line in text.splitlines())
            # break multi-headlines into a line each
            chunks = (phrase.strip()
                      for line in lines for phrase in line.split("  "))
            # drop blank lines
            text = '\n'.join(chunk for chunk in chunks if chunk)

            positive_score = 0
            positive_score_file = open("positive-words.txt", "r")
            for words in positive_score_file:
                if words in text:
                    # print(words)
                    positive_score = positive_score+1

            negative_score = 0
            negative_score_file = open("negative-words.txt", "r")
            for words in negative_score_file:
                if words in text:
                    negative_score = negative_score+1

            polarity_score = 0
            Polarity_Score = (positive_score - negative_score) / \
                ((positive_score + negative_score) + 0.000001)

            number_of_words = len(text.split())

            number_of_sentences = 0
            for words in text:
                if words == ".":
                    number_of_sentences = number_of_sentences+1
            number_of_sentences = number_of_sentences+1

            average_Sentence_length = number_of_words / number_of_sentences

            def is_complex_word(word):
                # Load the Pyphen module and create a dictionary for the English language
                dictionary = pyphen.Pyphen(lang='en')

                # Use the Pyphen module to get the number of syllables in the word
                syllables = len(dictionary.positions(word)) + 1

                # Return True if the word has more than two syllables, False otherwise
                return syllables > 2

            number_of_complex_words = 0
            for words in text.split():
                if is_complex_word(words):
                    number_of_complex_words = number_of_complex_words+1

            percentage_of_complex_words = number_of_complex_words / number_of_words

            fog_index = 0.4 * (average_Sentence_length +
                               percentage_of_complex_words)

            average_number_of_words_per_Sentence = number_of_words / number_of_sentences

            def count_cleaned_words(text):
                # Load the stop words from the nltk corpus
                stop_words = set(stopwords.words('english'))

                # Remove any punctuation from the text
                cleaned_text = text.translate(
                    str.maketrans('', '', string.punctuation))

                # Split the cleaned text into words
                words = cleaned_text.split()

                # Remove the stop words from the list of words
                filtered_words = [
                    word for word in words if word.lower() not in stop_words]

                # Return the total count of filtered words
                return len(filtered_words)

            total_cleaned_words_present_in_the_text = count_cleaned_words(text)

            Subjectivity_Score = 0
            Subjectivity_Score = (positive_score + negative_score) / \
                ((total_cleaned_words_present_in_the_text) + 0.000001)

            def count_personal_pronouns(text):
                # compile regular expression to match personal pronouns
                pattern = re.compile(r"\b(I|we|my|ours|us)\b", re.IGNORECASE)

                # match personal pronouns and count their occurrences
                count = len(re.findall(pattern, text))

                return count

            personal_pronoun_count = count_personal_pronouns(text)

            def count_total_characters(text):
                # split text into words
                words = text.split()

                # count total number of characters in each word
                total_characters = sum(len(word) for word in words)

                return total_characters

            sum_of_total_number_of_charecters = count_total_characters(text)
            average_word_length = sum_of_total_number_of_charecters/number_of_words

            dic = pyphen.Pyphen(lang='en')

            # define the input sentence

            # split the sentence into words
            words = text.split()

            # count the number of syllables for each word
            syllables_per_word = [len(dic.positions(word)) for word in words]

            # calculate the average number of syllables per word
            average_syllables_per_word = sum(
                syllables_per_word) / len(syllables_per_word)

            POSITIVE_SCORE.append(positive_score)
            NEGATIVE_SCORE.append(negative_score)
            POLARITY_SCORE.append(Polarity_Score)
            SUBJECTIVITY_SCORE.append(Subjectivity_Score)
            AVG_SENTENCE_LENGTH.append(average_Sentence_length)
            PERCENTAGE_OF_COMPLEX_WORDS.append(percentage_of_complex_words)
            FOG_INDEX.append(fog_index)
            AVG_NUMBER_OF_WORDS_PER_SENTENCE.append(
                average_number_of_words_per_Sentence)
            COMPLEX_WORD_COUNT.append(number_of_complex_words)
            WORD_COUNT.append(total_cleaned_words_present_in_the_text)
            AVG_SYLLABLE_PER_WORD.append(average_syllables_per_word)
            PERSONAL_PRONOUNS.append(personal_pronoun_count)
            AVG_WORD_LENGTH.append(average_word_length)
            URL.append(url)
            URL_ID.append(url_id)

        except KeyError:
            print("keyerror")
        except:
            print(url_id, "no. page not found")
        finally:
            print("url", url_id, "done")
            url_id = url_id+1

    else:
        print("url", url_id, "does not exist")
        POSITIVE_SCORE.append("NULL")
        NEGATIVE_SCORE.append("NULL")
        POLARITY_SCORE.append("NULL")
        SUBJECTIVITY_SCORE.append("NULL")
        AVG_SENTENCE_LENGTH.append("NULL")
        PERCENTAGE_OF_COMPLEX_WORDS.append("NULL")
        FOG_INDEX.append("NULL")
        AVG_NUMBER_OF_WORDS_PER_SENTENCE.append("NULL")
        COMPLEX_WORD_COUNT.append("NULL")
        WORD_COUNT.append("NULL")
        AVG_SYLLABLE_PER_WORD.append("NULL")
        PERSONAL_PRONOUNS.append("NULL")
        AVG_WORD_LENGTH.append("NULL")
        URL.append(url)
        URL_ID.append(url_id)

        url_id = url_id+1


data_dict = pd.DataFrame({
    "URL_ID": URL_ID,
    "URL": URL,
    "POSITIVE SCORE": POSITIVE_SCORE,
    "NEGATIVE SCORE": NEGATIVE_SCORE,
    "POLARITY SCORE": POLARITY_SCORE,
    "SUBJECTIVITY SCORE": SUBJECTIVITY_SCORE,
    "AVG SENTENCE LENGTH": AVG_SENTENCE_LENGTH,
    "PERCENTAGE OF COMPLEX WORDS": PERCENTAGE_OF_COMPLEX_WORDS,
    "FOG INDEX": FOG_INDEX,
    "AVG NUMBER OF WORDS PER SENTENCE": AVG_NUMBER_OF_WORDS_PER_SENTENCE,
    "COMPLEX WORD COUNT": COMPLEX_WORD_COUNT,
    "WORD COUNT": WORD_COUNT,
    "AVG SYLLABLE PER WORD": AVG_SYLLABLE_PER_WORD,
    "PERSONAL PRONOUNS": PERSONAL_PRONOUNS,
    "AVG WORD LENGTH": AVG_WORD_LENGTH

})


df = pd.DataFrame(data_dict)

# Save the DataFrame to an Excel file
df.to_excel('final_output.xlsx', index=False)


# print(df.to_string(index=False))

print("\n YOUR OUTPUT DATA HAS BEEN SUCCESSFULLY SAVED IN 'final_output.xlsx' EXCEL FILE IN THIS SAME FOLDER")
