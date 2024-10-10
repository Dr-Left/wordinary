""" shelfer.py - Store the vocabulary information into the shelf.

    typical usage: shelfer.work(basisName)
"""

import os
import random
import re
import shelve
import sys
import time

import xlrd


def work(fname):
    """
    fname is the Excel word list basis
    """
    if fname == "":
        # use the last shelf
        assert os.path.exists(".\\data\\dict.dat")
        return 0
    else:
        # create new shelf, with the basis in fname
        if os.path.exists(".\\data\\dict.dat"):
            os.remove(".\\data\\dict.dat")
            os.remove(".\\data\\dict.bak")
            os.remove(".\\data\\dict.dir")
        with shelve.open(".\\data\dict") as shelfFile:
            wb = xlrd.open_workbook(fname)
            ws = wb.sheet_by_index(0)
            # get the available row and column numbers
            rowNum = ws.nrows
            colNum = ws.ncols
            wordCol, transCol, start_row = whereAreWordsAndTranslations(ws)
            # print(
            #     "English words are in column %d, translations are in column %d."
            #     % (wordCol, transCol)
            # )
            # print("The first row is %d." % start_row)
            word_dict = {}
            for i in range(start_row, rowNum):
                word = str(ws.cell_value(i, wordCol))
                word = "".join(word.split())
                translation = str(ws.cell_value(i, transCol))
                translation = "".join(translation.split())
                # shelfFile[word] = translation
                word_dict[word] = translation
            print("词库数据创建成功！")
            shelfFile["dict"] = word_dict
            shelfFile.close()


def whereAreWordsAndTranslations(ws):
    """
    param:
        ws: the worksheet
    return:
        wordCol: the column number of the English words
        transCol: the column number of the translations
        start_row: the row number of the first word
    """
    # find where the English words are
    rowNum = ws.nrows
    colNum = ws.ncols
    wordRegex = re.compile(r"^[a-zA-Z]+$")  # English words
    wordCol = -1
    for i in range(colNum):
        col = ws.col_values(i)
        # randomly choose 10 rows to check, enhance efficiency
        for _ in range(1, 10):
            r = random.randint(0, rowNum - 1)
            val = str(col[r])
            val = "".join(val.split())
            if wordRegex.search(val):
                wordCol = i
                break
        if wordCol != -1:
            break
    if wordCol == -1:
        raise Exception("English words no find.")
    # find where the translations are
    chineseRegex = re.compile(r"[\u4e00-\u9fa5]+")  # Chinese characters
    transCol = -1
    for i in range(colNum):
        if i == wordCol:
            continue
        col = ws.col_values(i)
        for _ in range(1, 10):
            r = random.randint(0, rowNum - 1)
            val = str(col[r])
            val = "".join(val.split())
            if chineseRegex.search(val):
                transCol = i
                break
        if transCol != -1:
            break
    if transCol == -1:
        raise Exception("Translations no find.")
    for r in range(10):
        if wordRegex.search(str(ws.cell_value(r, wordCol))) and chineseRegex.search(
            str(ws.cell_value(r, transCol))
        ):
            start_row = r
            break
    return wordCol, transCol, start_row


if __name__ == "__main__":
    work(sys.argv[1])
