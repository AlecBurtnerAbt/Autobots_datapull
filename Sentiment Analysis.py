# -*- coding: utf-8 -*-
"""
Created on Thu Jul 26 13:22:02 2018

@author: C252059
"""

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
import time
import os
import win32com.client
import pandas as pd
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from bs4 import BeautifulSoup
import gzip
import shutil
import zipfile
import pandas as pd
import itertools    
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys


email = 'burtner_abt_alec@lilly.com'
pw = '1021547a'
downs = 1000
def getTwits(email,pw,downs):
    #Fire it up and login
    driver = webdriver.Chrome()
    driver.get('https://twitter.com/login/error?redirect_after_login=%2F')
    wait = WebDriverWait(driver,10)
    user_name = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="page-container"]/div/div[1]/form/fieldset/div[1]/input')))
    user_name.send_keys(email)
    pass_word = driver.find_element_by_xpath('//*[@id="page-container"]/div/div[1]/form/fieldset/div[2]/input')
    pass_word.send_keys(pw)
    login = driver.find_element_by_xpath('//*[@id="page-container"]/div/div[1]/form/div[2]/button')
    login.click()
    
    
    #search for lilly texts
    search = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="search-query"]')))
    search.send_keys('@LillyPad')    
    search.send_keys(Keys.ENTER)
    body = lambda: driver.find_element_by_tag_name('body')
    while downs >0:
        body().send_keys(Keys.PAGE_DOWN)
        downs -= 1
        
    raw = driver.find_elements_by_class_name('TweetTextSize')
    tweets = [tweet.text for tweet in raw]    
    return tweets
    #get the tweets from the first page
tweets = getTwits(email,pw,1000)
def analyzeTweets(tweets):
    from nltk.sentiment.vader import SentimentIntensityAnalyzer
    sid = SentimentIntensityAnalyzer()
    scores = pd.DataFrame(columns=['Sentence','Neg','Neu','Pos','Compound'])
    for i,sentence in enumerate(tweets):    
        score = sid.polarity_scores(sentence)
        neg = score['neg']
        pos = score['pos']
        neu = score['neu']
        compound = score['compound']
        scores = scores.append({'Sentence':sentence,'Neg':neg,'Neu':neu,'Pos':pos,'Compound':compound},ignore_index=True)
        negatives = scores[(scores['Compound']<0)].sort_values('Compound')
        positive = scores[scores['Compound']>0]
        negatives['Sentence'] =  negatives['Sentence'].str.lower()
        negatives = negatives[~negatives['Sentence'].str.contains('cancer')]
    import wordcloud
    import matplotlib.pyplot as plt
    from PIL import Image
    import imageio
    import numpy as np
    stopwords = set(wordcloud.STOPWORDS)
    stopwords.add('LillyPad')
    stopwords.add('https')
    words = []
    for sentence in negatives['Sentence']:
        sentence = sentence.split(' ')
        for word in sentence:
            words.append(word)
    words = ' '.join(words)
    plt.figure()
    skull_mask = np.array(Image.open(r'C:\Users\c252059\Desktop\skull.jpg'))
    wc = wordcloud.WordCloud(background_color='white',max_words=100,
                         mask = skull_mask,stopwords= stopwords,
                         contour_width=3,contour_color='black')
    wc.generate(words)
    plt.imshow(wc, interpolation='bilinear')
    plt.axis("off")
    plt.show()
    
    plt.figure()
    n, bins, patches = plt.hist(scores['Compound'],bins=100)
    cm = plt.cm.get_cmap('RdYlGn')
    bin_centers = .5*(bins[:-1]+bins[1:])       
    col = bin_centers-min(bin_centers)
    col/=max(col)
    for c,p in zip(col,patches):
        plt.setp(p,'facecolor',cm(c))
    plt.ylim(0,20)
    plt.xlabel('Polarity Score of Tweets')
    plt.ylabel('Count of Tweets')
    plt.title('Sentiment Analysis of Tweets About Lilly on 7/26')
