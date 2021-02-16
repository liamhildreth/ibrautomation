from textblob import TextBlob
import xlrd
import spacy
import docx2txt
import pandas as pd
from stanza.server import CoreNLPClient
import re
import nltk
import stanza
from pycorenlp import *

nlp = StanfordCoreNLP("http://localhost:9000/")

# get verb phrases

def get_verb_phrases(t):
    verb_phrases = []
    num_children = len(t)
    num_VP = sum(1 if t[i].label() == "VP" else 0 for i in range(0, num_children))

    if t.label() != "VP":
        for i in range(0, num_children):
            if t[i].height() > 2:
                verb_phrases.extend(get_verb_phrases(t[i]))
    elif t.label() == "VP" and num_VP > 1:
        for i in range(0, num_children):
            if t[i].label() == "VP":
                if t[i].height() > 2:
                    verb_phrases.extend(get_verb_phrases(t[i]))
    else:
        verb_phrases.append(' '.join(t.leaves()))

    return verb_phrases


# get position of first node "VP" while traversing from top to bottom

def get_pos(t):
    vp_pos = []
    sub_conj_pos = []
    num_children = len(t)
    children = [t[i].label() for i in range(0,num_children)]

    flag = re.search(r"(S|SBAR|SBARQ|SINV|SQ)", ' '.join(children))

    if "VP" in children and not flag:
        for i in range(0, num_children):
            if t[i].label() == "VP":
                vp_pos.append(t[i].treeposition())
    elif not "VP" in children and not flag:
        for i in range(0, num_children):
            if t[i].height() > 2:
                temp1,temp2 = get_pos(t[i])
                vp_pos.extend(temp1)
                sub_conj_pos.extend(temp2)
    # comment this "else" part, if want to include subordinating conjunctions
    else:
        for i in range(0, num_children):
            if t[i].label() in ["S","SBAR","SBARQ","SINV","SQ"]:
                temp1, temp2 = get_pos(t[i])
                vp_pos.extend(temp1)
                sub_conj_pos.extend(temp2)
            else:
                sub_conj_pos.append(t[i].treeposition())

    return (vp_pos,sub_conj_pos)


# get all clauses
def get_clause_list(sent):
    parser = nlp.annotate(sent, properties={"annotators":"parse","outputFormat": "json"})
    sent_tree = nltk.tree.ParentedTree.fromstring(parser["sentences"][0]["parse"])
    clause_level_list = ["S","SBAR","SBARQ","SINV","SQ"]
    clause_list = []
    sub_trees = []
    # sent_tree.pretty_print()

    # break the tree into subtrees of clauses using
    # clause levels "S","SBAR","SBARQ","SINV","SQ"
    for sub_tree in reversed(list(sent_tree.subtrees())):
        if sub_tree.label() in clause_level_list:
            if sub_tree.parent().label() in clause_level_list:
                continue

            if (len(sub_tree) == 1 and sub_tree.label() == "S" and sub_tree[0].label() == "VP"
                and not sub_tree.parent().label() in clause_level_list):
                continue

            sub_trees.append(sub_tree)
            del sent_tree[sub_tree.treeposition()]

    # for each clause level subtree, extract relevant simple sentence
    for t in sub_trees:
        # get verb phrases from the new modified tree
        verb_phrases = get_verb_phrases(t)

        # get tree without verb phrases (mainly subject)
        # remove subordinating conjunctions
        vp_pos,sub_conj_pos = get_pos(t)
        for i in vp_pos:
            del t[i]
        for i in sub_conj_pos:
            del t[i]

        subject_phrase = ' '.join(t.leaves())

        # update the clause_list
        for i in verb_phrases:
            clause_list.append(subject_phrase + " " + i)

    return clause_list

# input word doc from your C:\Users drive
article = input("Word document name: ")
text = docx2txt.process(str(article) + ".docx")
nlpfacts = []

# applying nlp library to the article
processed = TextBlob(text)

# parsing through the article to pull all sentences
for line in processed.sentences:
    if line != '':
        sent = re.sub(r"(\.|,|\?|\(|\)|\[|\])", " ", line.string)
        if sent != '' and sent != ' 'and len(sent)!=1:
            try:
                clauses = get_clause_list(sent)
                for clause in clauses:
            # setting threshold to 1 in order to pull every single sentence
                    clause = TextBlob(clause)
                    if clause.subjectivity <= 0.5:
                        nlpfacts.append(clause)
            except:
                sent = TextBlob(sent)
                if sent.subjectivity <=0.8:
                    nlpfacts.append(clause)


# creating pandas dataframe to store the sentences to be uploaded to factchecking spreadsheet                
df = pd.DataFrame()
df['Facts']= nlpfacts

# exporting sentences to pre-existing spreadsheet for each article
df.to_excel(str(article) +'_Factcheck1.xlsx', index=False)
