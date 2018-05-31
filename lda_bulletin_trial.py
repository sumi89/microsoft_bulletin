#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu May 31 00:28:04 2018

@author: sumi
"""
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed May 30 23:36:03 2018

@author: sumi
"""

from sklearn.feature_extraction.text import TfidfVectorizer, CountVectorizer
from sklearn.decomposition import LatentDirichletAllocation
import numpy as np
import xlrd
from xlrd import open_workbook
from itertools import chain

sheet = []
list_book = []
for i in range (20):
    #book = xlrd.open_workbook('/Users/sumi/Documents/CPS_5310_grades.xlsx')
    book = xlrd.open_workbook('/Users/sumi/Documents/summer_18/MicrosoftSecurity/Implementation/data.xlsx')
    sheet = book.sheet_by_index(i)
    
    list_sheet = []
    
    for k in range(1,sheet.nrows):
        list_sheet.append(str(sheet.row_values(k)[3]))
    list_book.append(list_sheet)

document = list(chain.from_iterable(list_book))


def display_topics(H, W, feature_names, documents, no_top_words):
    for topic_idx, topic in enumerate(H):
        print ("Topic %d:" % (topic_idx))
        print (" ".join([feature_names[i]
                        for i in topic.argsort()[:-no_top_words - 1:-1]]))
#        top_doc_indices = np.argsort( W[:,topic_idx] )[::-1][0:no_top_documents]
#        for doc_index in top_doc_indices:
#            print (documents[doc_index])

# Single line documents from http://web.eecs.utk.edu/~berry/order/node4.html#SECTION00022000000000000000
#documents = [
#            "Human machine interface for Lab ABC computer applications",
#            "A survey of user opinion of computer system response time",
#            "The EPS user interface management system",
#            "System and human system engineering testing of EPS",
#            "Relation of user-perceived response time to error measurement",
#            "The generation of random, binary, unordered trees",
#            "The intersection graph of paths in trees",
#            "Graph minors IV: Widths of trees and quasi-ordering",
#            "Graph minors: A survey"
#            ]



# LDA can only use raw term counts for LDA because it is a probabilistic graphical model
tf_vectorizer = CountVectorizer(max_df=0.95, min_df=2, stop_words='english')
tf = tf_vectorizer.fit_transform(document)
tf_feature_names = tf_vectorizer.get_feature_names()

no_topics = 5


# Run LDA
lda_model = LatentDirichletAllocation(n_topics=no_topics, max_iter=5, learning_method='online', learning_offset=50.,random_state=0).fit(tf)
lda_W = lda_model.transform(tf)
lda_H = lda_model.components_

no_top_words = 3

display_topics(lda_H, lda_W, tf_feature_names, document, no_top_words)












### CORRECT CODE ########
#book = xlrd.open_workbook('/Users/sumi/Documents/CPS_5310_grades.xlsx')
#sheet = book.sheet_by_index(0)
#
#list_j = []
#
#for k in range(1,sheet.nrows):
#    list_j.append(str(sheet.row_values(k)[1]))
#### END (CORRECT CODE) #########

#################
#
#book = open_workbook('/Users/sumi/Documents/CPS_5310_grades.xlsx')
#sheet = book.sheet_by_index(0) #If your data is on sheet 1
#
#column1 = []
#column2 = []
##...
#
#for row in range(1, sheet.nrows): #start from 1, to leave out row 0
#    column1.append(sheet.cell(row, 0)) #extract from first col
#    column2.append(sheet.cell(row, 1))
#    #...