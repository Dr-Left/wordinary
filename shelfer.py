''' shelfer.py - Store the vocabulary information into the shelf.
'''
import xlrd, shelve, os, sys, time

def work():
    if not os.path.exists('.\\data'):
        os.makedirs('.\\data')
    elif os.path.exists('.\\data\\dict.dat'):
        print('词库数据已存在')
        return 0
    shelfFile = shelve.open('.\\data\dict')
    fname = '.\\source\\' + '高考考纲词汇.xlsx'
    wb = xlrd.open_workbook(fname)
    ws = wb.sheet_by_index(0)
    # get the available row and column numbers
    rowNum = ws.nrows
    colNum = ws.ncols
    for i in range(1, rowNum):
        word = ws.cell_value(i, 0)
        word = ''.join(word.split())
        meaning = ws.cell_value(i, 1)
        meaning = ''.join(meaning.split())
        shelfFile[word] = meaning

    print('词库数据创建成功！')
    time.sleep(0.5)
    shelfFile.close()


if __name__ == '__main__':
    work()
