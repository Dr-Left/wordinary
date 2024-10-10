""" shelfer.py - Store the vocabulary information into the shelf.

    typical usage: shelfer.work(basisName)
"""

import os
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
            wordCol, transCol = whereAreWordsAndTranslations(ws)
            for i in range(1, rowNum):
                word = str(ws.cell_value(i, wordCol))
                word = "".join(word.split())
                translation = str(ws.cell_value(i, transCol))
                translation = "".join(translation.split())
                shelfFile[word] = translation
            print("词库数据创建成功！")
            shelfFile.close()


def whereAreWordsAndTranslations(ws):
    # find where the English words are
    rowNum = ws.nrows
    colNum = ws.ncols
    wordRegex = re.compile(r"^[a-zA-Z]+$")  # English words
    wordCol = -1
    for i in range(colNum):
        for j in range(1, rowNum):
            val = str(ws.col_values(i)[j])
            val = "".join(val.split())
            if wordRegex.search(val):
                wordCol = i
                break
        if wordCol != -1:
            break
    if wordCol == -1:
        raise Exception("Words no find.")
    # find where the translations are
    chineseRegex = re.compile(r"[\u4e00-\u9fa5]+")  # Chinese characters
    transCol = -1
    for i in range(colNum):
        if i == wordCol:
            continue
        for j in range(1, rowNum):
            val = str(ws.col_values(i)[j])
            val = "".join(val.split())
            if chineseRegex.search(val):
                transCol = i
                # print('DEBUG:' + val)
                break
        if transCol != -1:
            break
    if transCol == -1:
        raise Exception("Translations no find.")
    return wordCol, transCol


if __name__ == "__main__":
    work(sys.argv[1])
