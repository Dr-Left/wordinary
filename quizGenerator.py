''' quizGenerator.py - To generate quizes according to the vocabulary table

The programme first get the vocabulary table from a xlsx document,
and then generate the correspoding dictation quizes.

Typical using examples:

quizGenerator.work()

'''
import xlrd
import re
import os
import glob
import docx
import random
import currentTime
import time

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE

currentDate = time.strftime("%Y-%m-%d", time.localtime())

# regex family
wordRegex = re.compile(r'^[a-zA-Z]+$')  # English words
chineseRegex = re.compile(r'[\u4e00-\u9fa5]+')  # Chinese characters
rangeRegex = re.compile(r'[0-9]+\w*[0-9]+') # a range e.g. 1-100
numRegex = re.compile(r'[0-9]+')    # pure integers



sourceName = ''
binPath = ['.\\',
           '.\\生成的文件\\',
           '.\\source\\']



# the main funciton(find the xlsx source)
def work():
    global sourceName
    # get the vocabulary table
    files = []
    for path in binPath:
        if not os.path.exists(path):
            continue
        files += glob.glob(os.path.join(path, '*.xlsx')) # return a list
    
    atLeastOnce = False #   check if there really exists a xlsx file
    
    for file in files:
        atLeastOnce = True
        #   give the user an idea of which file is being worked on, and a chance to quit
        bsName = os.path.basename(file)
        print('正在对\n\t' + bsName + '\n生成默写...')
        print('是否需要此表？(Y/N)')
        char = input()
        if char != 'Y':
            continue


        mo = chineseRegex.findall(bsName)
        if mo != []:
            sourceName = mo[0]
            #print(sourceName)
        try:
            gen(file)
            print()
        except Exception as exc:
            print('This workbook has failed: ' + file + '\n\n' + str(exc))
    if atLeastOnce == False:
        print('未找到xlsx文件！')
    else:
        print('所有文件均已执行完毕！')



def gen(fname):
    '''get the words from the xlsx document which is directed by fname
    and then goto quizPrint() to print
    gen() ask a lot of personalizing questions

    '''
    wb = xlrd.open_workbook(fname)
    ws = wb.sheet_by_index(0)
    
    # get the available row and column numbers
    rowNum = ws.nrows
    colNum = ws.ncols
    wordCol = -1
    
    # find where the English words are
    for i in range(colNum):
        for j in range(1,rowNum):
            val = str(ws.col_values(i)[j])
            val = ''.join(val.split())
            if wordRegex.search(val):
                wordCol = i
                break
        if wordCol != -1:
            break
    if wordCol == -1:
        raise Exception('Word no find.')

    # find where the translations are
    transCol = -1
    for i in range(colNum):
        if i == wordCol:
            continue
        for j in range(1,rowNum):
            val = str(ws.col_values(i)[j])
            val = ''.join(val.split())
            if chineseRegex.search(val):
                transCol = i
                #print('DEBUG:' + val)
                break
        if transCol != -1:
            break
    if transCol == -1:
        raise Exception('Translation no find.')

    
    # generate the quiz form

    # ask for the number of students
    print('请输入不同的版本数目（统一版本请输入1）：')
    while True:
        try:
            stuNum = int(input())
            break
        except:
            print('无效的输入！')

    # ask for the range of words， and put them to head and tail
    print('请输入本次默写的单词范围（例如：100-199， 最大' + str(rowNum - 1) + '道题目）：')
    while True:
        inp = input()
        mo = rangeRegex.search(inp)
        if mo == None:
            print('无效的输入！请参考例子的格式。')
        else:
            numlist = numRegex.findall(inp)
            try:
                head = int(numlist[0])  # The start point
                tail = int(numlist[1])  # The end point
                if head > tail    or    tail > (rowNum - 1) :
                    raise Exception('末尾越界。或者首大于尾')
            except Exception as exc:
                print('无效的输入！数字错误：  %s' %(exc))
                continue
            break

    # ask if the user need eng to chn or chn to eng
    print('请问需要英翻中（1）还是中翻英（2）？（请输入选择后面的字母）：')
    while True:
        try:
            opt = int(input())
            if opt not in range(1,3):
                raise Exception('Bad input')
            break
        except:
            print('无效的输入！')

    # ask how many words should appear on every quiz
    print('请问需要每张默写纸上设置几道小题？（最大' + str(tail - head + 1) + '道题目）：')
    while True:
        try:
            questionNum = int(input())
            if questionNum > tail - head + 1:
                
                raise Exception('问题数量大于范围数量！')
            break
        except Exception as exc:
            print('无效的输入！%s'  % (esc))
     
    # read in all the words and its translations
    wordList = {}
    sampleList = random.sample(range(head, tail + 1), questionNum)
    for i in sampleList:   # loop all the rows, search for the word
        key = str(ws.col_values(wordCol)[i])
        key = ' '.join(key.split())
        val = str(ws.col_values(transCol)[i])
        val = ' '.join(val.split())
        wordList[key] = val  # use a dictionary to storage: key for the word, value for the translation
        
    # do a loop, generate a different quiz for every student
    for i in range(stuNum):
        printQuiz(i, wordList, opt, head, tail, answerOn=False)
        printQuiz(i, wordList, opt, head, tail, answerOn=True)

  
def printQuiz(stu, wordList, opt, head, tail, answerOn=False):
    '''print the quiz beautifully into docx files
    '''
    # create a document
    doc = docx.Document()
    
    # change the style
    doc.styles['Normal'].font.name = 'Cambria'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    doc.styles['Normal'].font.size = Pt(15)
    doc.styles['Normal'].font.color.rgb = RGBColor(0,0,0)
    
    # create a style
    style_song = doc.styles.add_style('Song', WD_STYLE_TYPE.CHARACTER)
    style_song.font.name = '微软雅黑'
    doc.styles['Song']._element.rPr.rFonts.set(qn('w:eastAsia'), u'微软雅黑')
    
    # add the headings
    h1 = doc.add_heading('', level=0)
    h1.add_run('默写纸', style='Song')
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    h2 = doc.add_heading('', level=4)
    h2.add_run('日期：' + currentDate, style='Song')
    h2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    
    # divide into two sections
    section = doc.sections[0]
    sectPr = section._sectPr
    cols = sectPr.xpath('./w:cols')[0]
    cols.set(qn('w:num'),'2')

    # output
    
    #   设置行距和字体，后面附上下划线。
    #   TODO:如果超过一页纸要进行警告
    #   check for the opt to decide which mode will be executed
    
    for i in wordList:
        para = doc.add_paragraph(style = 'List Number')
        para.paragraph_format.line_spacing = Pt(30)

        if opt == 1:    # 英翻中
            st = i
            ans = wordList[i]
        else:   # 中翻英
            st = wordList[i]
            ans = i

        # check if this is an answer sheet
        if answerOn == False :
            ans = '\t\t\t'
        else:
            st += ':'   #   pretty print

        # add run in the paragraph    
        run1 = para.add_run()
        run1.text = st
        run1.font.size = Pt(15)
        
        run2 = para.add_run()
        run2.text = ans
        run2.font.size = Pt(15)
        run2.font.underline = True
           
    # save
    # 把xlsx的文件名写在生成的默写纸里面（默写纸_高频词汇表_1-100词_版本1_YYYY-MM-DD hh.mm.xlsx）
    
    #   show the range of words in the file name
    rangeName = str(head) + '-' + str(tail) + '词'
    #   show the version in the file name
    versionName = '版本' + str(stu + 1)
    #   get the current time
    cT = currentTime.getTime()
    #   check whether its a quiz or an answerbook
    if answerOn == False:
        '''默写纸_高频词汇表_1-100词_版本1_YYYY-MM-DD hh.mm.xlsx
        '''
        path = '_'.join(['默写纸', sourceName, rangeName, versionName, cT]) + '.docx'
    else:
        '''答案_高频词汇表_1-100词_版本1_YYYY-MM-DD hh.mm.xlsx
        '''
        path = '_'.join(['答案', sourceName, rangeName, versionName, cT]) + '.docx'
    #   make sure the directory exists
    if not os.path.exists(binPath[1]):
        os.makedirs(binPath[1])
    #   save
    path = binPath[1] + path
    doc.save(path)
    print('生成成功：' + os.path.abspath(path))
    

if __name__=='__main__':
    print('main')
    work()
    
