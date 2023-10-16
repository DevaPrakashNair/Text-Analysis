import os
import string
import nltk
from nltk.corpus import stopwords
import re
import syllables
import openpyxl



def all_stopwords():
    stop_words = ""
    stopword_folder = r"C:\Users\Deva\Study_material\Blackcoffer\20211030 Test Assignment\StopWords"
    for file in os.listdir(stopword_folder):
        f = os.path.join(stopword_folder, file)
        if os.path.isfile(f):
            word = open(f, "r")
            stop_words += word.read()
    return stop_words.split("\n")


def text_analysis():
    stop_words = all_stopwords()
    directory = r"C:\Users\Deva\Study_material\Blackcoffer\20211030 Test Assignment\Py\text"
    for filename in os.listdir(directory):
        f = os.path.join(directory, filename)
        print(f)
        if os.path.isfile(f):
            file = open(f, "r", encoding="utf-8").read()
            # Sentiment analysis
            file = file.replace("\n", " ")
            file = file.replace('“', "")
            file = file.replace('”', "")
            file = " ".join([i for i in file.split(r" ") if i not in stop_words])
            # print(nltk.sent_tokenize(file))

            positive = open(
                r"C:\Users\Deva\Study_material\Blackcoffer\20211030 Test Assignment\MasterDictionary\positive-words.txt",
                "r").read().replace("\n", " ").split(" ")
            negative = open(
                r"C:\Users\Deva\Study_material\Blackcoffer\20211030 Test Assignment\MasterDictionary\negative-words.txt",
                "r").read().replace("\n", " ").split(" ")
            pos = 0  # positive score
            neg = 0  # negative score
            for i in nltk.word_tokenize(file):
                if i in positive:
                    pos += 1
                elif i in negative:
                    neg += -1

            neg *= -1

            polarity = (pos - neg) / ((pos + neg) + 0.000001)  # polarity

            subjective_score = (pos + neg) / (file.count(" ") + 1 + 0.000001)  # subjective score

            # Analysis of readability

            avg_sentence_length = (file.count(" ") + 1) / (file.count(".") + 1)

            complex_word_count = len(re.compile(r"\w+([aeiouy]+\w+)+\w+").findall(file))

            per_complex_words = complex_word_count / (file.count(" ") + 1)

            fog_index = 0.4 * (avg_sentence_length + per_complex_words)

            # word count after stop word and punctuation removal

            words_token = nltk.word_tokenize(file)
            stopword = stopwords.words("english")

            filtered = []
            for i in words_token:
                if i not in stopword and i not in string.punctuation:
                    filtered.append(i)

            word_count = len(filtered)

            # count syllables

            no_of_syllables = 0
            for i in words_token:
                no_of_syllables += syllables.estimate(i)

            # personal pronouns

            personal_pronouns_regex = re.compile(r'(?<!\w)(I|we|my|ours|us)(?!\w)(?!US)')
            count_personal_pronous = len(personal_pronouns_regex.findall(file))

            # Average word length

            avg_word_length = (len(file) - file.count(" ")) / (file.count(" ") + 1)

            # Saving output

            output = openpyxl.load_workbook(r"C:\Users\Deva\Study_material\Blackcoffer\20211030 Test "
                                            r"Assignment\Output Data Structure.xlsx")

            sh = output.active

            for row in sh.iter_rows(min_row = 2, min_col = 1):
                if str(row[0].value) == filename[:-4]:
                    row[2].value = pos
                    row[3].value = neg
                    row[4].value = polarity
                    row[5].value = subjective_score
                    row[6].value = avg_sentence_length
                    row[7].value = per_complex_words
                    row[8].value = fog_index
                    row[9].value = avg_sentence_length
                    row[10].value = complex_word_count
                    row[11].value = word_count
                    row[12].value = no_of_syllables
                    row[13].value = count_personal_pronous
                    row[14].value = avg_word_length

                    output.save(r"C:\Users\Deva\Study_material\Blackcoffer\20211030 Test Assignment\Output Data Structure.xlsx")
                    print(filename)
        # break
