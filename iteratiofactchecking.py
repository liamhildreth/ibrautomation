# this one isn't done, waiting for testing data from all the first iteration factchecks
# aiming to automate the repetitive process of finding facts that do not change between iterations and pasting them into the next factcheck

from textblob import TextBlob
import xlrd
import docx2txt
import pandas as pd
import re

# input word doc from your C:\Users drive
article = input("Article name: ")
iteration = input("Iteration #: ")
text = docx2txt.process(str(article) +"_"+str(iteration)+ ".docx")
nlpfacts = []

# applying nlp library to the article
processed = TextBlob(text)

# parsing through the article to pull all facts
for line in processed.sentences:
    if line != '':
        sent = re.sub(r"(\.|,|\?|\(|\)|\[|\])", " ", line.string)
        if sent != '' and sent != ' 'and len(sent)!=1:
            sent = sent.replace('\n',' ')
            sent = TextBlob(sent)
            if sent.subjectivity <=0.8:
                nlpfacts.append(sent)


# creating pandas dataframe to store the sentences to be uploaded to factchecking spreadsheet                
df = pd.DataFrame()
df['Facts']= nlpfacts

# exporting sentences to pre-existing spreadsheet for each article
df.to_excel(str(article) +'_Factcheck1.xlsx', index=False)
