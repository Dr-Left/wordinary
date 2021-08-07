#! python3
# main.py - the main program

import quizGenerator
import outputToExcel
import sys
from time import sleep


print('欢迎使用wordinary~~~\n')
sleep(1)
print('开发者：Dr-Left')
sleep(0.5)
print('有任何问题请联系我:https://github.com/Dr-Left/')
sleep(0.5)
print('QQ:632826792\n')
sleep(1)

print('本程序旨在协助英语老师便捷地生成背单词的表格，和对应的默写纸！')
sleep(0.5)
print('让英语老师鼠标一点~默写不愁！\n')
sleep(1)

print('Let\'s begin!\n')

while True:
    print('请问需要执行哪个功能？（生成默写纸输入1，提取高频词输入2，退出请输入0）：')
    while True:
        inp = input()
        if inp == '0':
            sys.exit()
        if inp != '1' and inp !='2':
            print('输入不合法！重新输入：')
            continue
        else:
            opt = int(inp)
            break

    if opt == 1: # quizGenerate
        try:
            quizGenerator.work()
        except Exception as e:
            print('默写纸生成失败！\n %s' %e)
    else:   # wordExtract]
        try:
            outputToExcel.output()
        except Exception as e:
            print('高频词提取失败！\n %s' %e)
    print()

