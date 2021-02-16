from textblob import TextBlob
import xlrd
import spacy
import docx2txt
from stanza.server import CoreNLPClient
import re
import nltk
import stanza
from pycorenlp import *

nlp = StanfordCoreNLP("http://localhost:9000/")

# excel file of factcheck done by IBR researchers
xlfile = r"C:\Users\liamh\Downloads\Test_1.xlsx"
wb = xlrd.open_workbook(xlfile)

# reading the excel file
sheet = wb.sheet_by_index(0)
row_count = sheet.nrows
humanfacts = []

# creating a list to store the sentences selected as facts by IBR researchers
for curr_row in range(1,row_count):
    fact = sheet.cell_value(curr_row, 1)
    fact2 = re.sub(r"\W","",fact)
    humanfacts.append(fact2)

# opening an example of an IBR article to parse through looking for facts
text = docx2txt.process("11. Modern Health - Draft 4.docx")
nlpfacts = []

processed = TextBlob(text)

# parsing through the IBR article to look for facts
for line in processed.sentences:
    if line != '':
        sent = re.sub(r"(\.|,|\?|\(|\)|\[|\])", " ", line.string)
        if sent != '' and sent != ' 'and len(sent)!=1:
            
            # major decision to tell if a sentence contains a fact or not, arbitrarily chose 0.5 on a scale of 0:1 to test
            if sent.subjectivity <0.5:
                line2 = re.sub(r"\W","" sent.string)
                nlpfacts.append(line2)


numbercorrect = 0
totalfacts = len(humanfacts)
both = {}

# testing how many of the facts picked up through the nlp were also selected by IBR researchers
for autof in nlpfacts:
    for humanf in humanfacts:
        if autof in humanf:
            numbercorrect += 1
            humanfacts.remove(humanf)
            both[autof] = humanf
            break


if totalfacts != 0:
    accuracy = numbercorrect/totalfacts
else:
    accuracy = 0

print(humanfacts)
print(len(humanfacts))
print(len(nlpfacts))
print(accuracy)
print(both)
