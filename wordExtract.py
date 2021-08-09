''' wordExtract.py - Extract words from texts in bin and return a dictionary
with keys of words and values of the times the word appears.

p.s. Compatible with documents ends with .txt, .docx


Typical usage example:

wordList = wordExtract.extract()

'''

import os
import glob
import re
import docx
import time
import logging

logging.basicConfig(level = logging.DEBUG, format = ' %(asctime)s - %(levelname)s - %(message)s')
logging.disable(logging.CRITICAL)

def cutUp(text):
    text = text.lower()
    # not an English alphabet
    myRegex = re.compile(r'[^a-zA-Z]+')# this should be plus but not star
    text = myRegex.sub(' ', text)
    
    #print(myRegex.findall(text))
    #print('\n\n\n\n' + text)
    
    return text

def extract():
    currentDirectory = os.getcwd()
    logging.debug(currentDirectory)
    words = {}


    # get all the input files the user has provided(txt, docx)

    res_txt = glob.glob(os.path.join('.\\source', '*.txt')) # return a list
    res_docx = glob.glob(os.path.join('.\\source', '*.docx')) # return a list

    logging.debug('txt: ' + str(res_txt))
    logging.debug('docx: ' + str(res_docx))


    if res_txt == [] and res_docx == []:
        raise Exception('Text file no found!')
    
    # for each text, collect the words and the corresponding appearing times

    for dr in res_txt:
        try:
            myFile = open(dr)
        except Exception as exc:
            print(exc + '\n cannot open directory.')
        cont = myFile.read()
        cont = cutUp(cont)

        wordList = cont.split()

        for wd in wordList:
            words.setdefault(wd, 0)
            words[wd] += 1

    for dr in res_docx:
        try:
            myFile = docx.Document(dr)
        except Exception as exc:
            print(exc + '\n cannot open directory.')
        cont = ''
        for paragraph in myFile.paragraphs:
            cont += paragraph.text
        cont = cutUp(cont)

        wordList = cont.split()

        for wd in wordList:
            words.setdefault(wd, 0)
            words[wd] += 1

            
    words = sorted(words.items())
    print('词语提取成功')
    time.sleep(0.5)
    #print(words)
    return words   #    return a dictionary of words   

if __name__=='__main__':
    extract()








