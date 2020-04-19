import win32com.client as wincl
speak = wincl.Dispatch("SAPI.SpVoice")

import speech_recognition as sr
import pyaudio


def input():
    r = sr.Recognizer()                                                                                   
    
    with sr.Microphone() as source:
        print("Speak:")                                                                                   
       
    audio = r.listen(source)   
       
    try:
        query = r.recognize_google(audio, language='en-in')
    
        print("User: " + query + "\n")
    
    except sr.UnknownValueError:
        speak.Speak("Could not understand audio")
    
    except sr.RequestError as e:
        speak.Speak("Could not request results; {0}".format(e))
    
    return query


import io
import random
import string 
import warnings
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import warnings
warnings.filterwarnings('ignore')


import nltk

f=open('Huey2.0.txt','r',errors = 'ignore')
raw=f.read()
raw=raw.lower()# converts to lowercase
nltk.download('punkt') # first-time use only
nltk.download('wordnet') # first-time use only
sent_tokens = nltk.sent_tokenize(raw)# converts to list of sentences 
word_tokens = nltk.word_tokenize(raw)# converts to list of words

lemmer = nltk.stem.WordNetLemmatizer()
#WordNet is a semantically-oriented dictionary of English included in NLTK.
def LemTokens(tokens):
    return [lemmer.lemmatize(token) for token in tokens]
remove_punct_dict = dict((ord(punct), None) for punct in string.punctuation)
def LemNormalize(text):
    return LemTokens(nltk.word_tokenize(text.lower().translate(remove_punct_dict)))

GREETING_INPUTS = ("hello", "hi", "greetings", "sup", "what's up","hey",)
GREETING_RESPONSES = ["hi", "hey", "*nods*", "hi there", "hello", "I am glad! You are talking to me"]
def greeting(sentence):
 
    for word in sentence.split():
        if word.lower() in GREETING_INPUTS:
            speak.Speak(random.choice(GREETING_RESPONSES))
        
        
def response(user_response):
    Huey_response=''
    sent_tokens.append(user_response)
    TfidfVec = TfidfVectorizer(tokenizer=LemNormalize, stop_words='english')
    tfidf = TfidfVec.fit_transform(sent_tokens)
    vals = cosine_similarity(tfidf[-1], tfidf)
    idx=vals.argsort()[0][-2]
    flat = vals.flatten()
    flat.sort()
    req_tfidf = flat[-2]
    if(req_tfidf==0):
        Huey_response=speak.Speak("I am sorry! I don't understand you")
        return Huey_response
    else:
        Huey_response = user_response + sent_tokens[idx]
        return Huey_response        
        
        
        
flag=True
while(flag==True):
    user_response = query
    user_response=user_response.lower()
    if(user_response!='bye'):
        if(user_response=='thanks' or user_response=='thank you' ):
            flag=False
            speak.Speak("You are welcome..")
       
    else:
        flag=False
        print("ROBO: Bye! take care..")    
                 
        
        
        
        
        
        
        
        