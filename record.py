# -*- coding: utf-8 -*-
# @Time : 2020/12/16 15:00
# @Author : Gordon
# @Editor : ChrisZuo

'''
    命令行参数sample：
    python record.py "./wordListSample.xlsx"    "./"      2               1000
                      sourceFilePath          binPath    repetition   interval(ms)
'''

from pydub import AudioSegment
from pyquery import PyQuery as pq
import wget
import os
import time
import requests
import sys

sourceFilePath = sys.argv[1]
binPath = sys.argv[2]
repetition = int(sys.argv[3])
interval = int(sys.argv[4])
fileName = os.path.basename(sourceFilePath).split('.')[0]

def main():
    if not os.path.exists(binPath):
        os.mkdir(binPath)
    wordList = get_wordList()
    look_up_words(wordList)

    
def get_wordList():
    '''
    获取词表
    
    '''
    wordList=[] #一个list, 存放单词
    if sourceFilePath.endswith(".txt"): # 文本文件，每行一个词
        with open(r'.\wordList.txt', encoding='utf-8') as f:
            lines = f.readlines()
            for line in lines:
                word = line.strip() # 去除首尾空格
                wordList.append(word)
    if sourceFilePath.endswith(".xlsx") or sourceFilePath.endswith(".xls"): # Excel 表格
        wb = xlrd.open_workbook(fname)
        ws = wb.sheet_by_index(0)
        # get the available row and column numbers
        rowNum = ws.nrows
        colNum = ws.ncols
        wordCol = -1 # the English row
        # find where the English words are
        for i in range(colNum):
            for j in range(1,rowNum):
                val = str(ws.col_values(i)[j])
                val = ''.join(val.split())
                wordRegex = re.compile(r'^[a-zA-Z]+$')  # English words
                if wordRegex.search(val):
                    wordCol = i
                    break
            if wordCol != -1:
                break
        if wordCol == -1:
            raise Exception('Word no find.')
        for i in range(colNum):
            word = ''.join(str(ws.col_values(i)[j]).split())
            wordList.append(word) 
    return wordList


def look_up_words(wordList):
    """
    爬取有道词典上的单词音标、词义、发音mp3
    输出为wav
    """
    song = pause(1000) # 最后输出的音频
    
    for word in wordList: #遍历单词列表中的每个单词
        if not os.path.exists(r".\sounds"):
            os.mkdir(r".\sounds")
        target_name = os.path.join(r".\sounds", word + ".mp3")
        if not os.path.exists(target_name): # 之前没用过，需要从网上爬取
            try:
                url = f'http://dict.youdao.com/w/eng/{word}/#keyfrom=dict2.index' # 获取有道词典中要查询单词所在的网页地址
                url_mp3 = f'http://dict.youdao.com/dictvoice?audio={word}'#获取有道词典中要查询单词
                res = requests.get(url).text  # 爬取单词网页文本
                doc = pq(res) # 使用PyQuery解析网页文本
                if doc(".keyword").text() == '': # 如果无法获取到keyword标签，证明单词没有查到，提醒单词不存在。
                    print(f"{word}单词不存在！")
                    continue
                else:
                    wget.download(url_mp3, out=target_name)
            except Exception as exc:
                print(exc)
        songclip = pause(interval) + AudioSegment.from_file(target_name)
        song = song + songclip * repetition
    song.export(binPath + '\\' + fileName + '.wav', format='wav')

def pause(time):
    return AudioSegment.silent(duration=time)

if __name__ == '__main__':
    main()


