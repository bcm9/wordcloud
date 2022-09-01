#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Returns n most frequent words in a word document and plots word cloud

n = n most frequent words in docfilename in folder

Created on Sun May 10 13:58:13 2020

@author: bcm9
"""
# import packages
import os
from collections import Counter 
import readDocx
import re
import matplotlib.pyplot as plt
from wordcloud import WordCloud, STOPWORDS
    
# returns most common words in word document
def mostfrequentwordcloud(folder,docfilename,n):    
    # pre-processing: set directory, confirm, set filename
    os.chdir(folder)
    os.getcwd() # which working directory
    # doc = docx.Document(docfilename) # load in doc as document
    
    # import doc as string using readDocx
    fulltextstr=(readDocx.getText(docfilename))
      
    # split() returns list of each word within the string
    split_it = fulltextstr.split() 
    # Pass the split_it list to instance of Counter class. 
    wordlist = Counter(split_it) 
      
    # most_common() returns n frequent words 
    most_freq = wordlist.most_common(n) 
    text = re.sub(r'==.*?==+', '', fulltextstr)
    text = text.replace('\n', '')
    
    # word cloud function
    def plot_wordcloud(wordcloud):
        # Set figure size
        plt.figure(figsize=(40, 30))
        # Display image
        plt.imshow(wordcloud) 
        # No axis details
        plt.axis("off");
        
    # generate word cloud
    wordcloud = WordCloud(width= 3000, height = 2000, random_state=1, background_color='white', colormap='Paired', collocations=False, stopwords = STOPWORDS).generate(text)
    # plot
    plot_wordcloud(wordcloud)
    return(most_freq)