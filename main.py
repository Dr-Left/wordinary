'''main.py - The main program of wordinary.

work of Dr-Left
https://github.com/Dr-Left/wordinary
Welcome to fork, star, or PR!

Please read README.md before using the programme.

Good luck!

'''
#!/usr/bin/env python3

import quizGenerator
import outputToExcel
import sys
import os
from time import sleep

# some welcome words
print('欢迎使用wordinary~~~\n')
sleep(1)
print('开发者：Dr-Left')
sleep(0.5)
print('有任何问题请联系我:https://github.com/Dr-Left/')
sleep(0.5)
print('QQ:632826792\n')
sleep(0.5)

print('本程序旨在协助英语老师便捷地生成背单词的表格，和对应的默写纸！')
sleep(0.5)
print('让英语老师鼠标一点~默写不愁！\n')
sleep(0.5)
print('首次运行请先阅读README.md，用记事本打开即可。\n')
sleep(0.5)
print('Let\'s begin!\n')
sleep(0.5)
# main loop
while True:
    # ask for which function the user would like to use
    print('请问需要执行哪个功能？（生成默写纸输入1，提取高频词输入2，退出请输入0）：')
    inp = input()
    if inp == '0':
        sys.exit()
    elif inp == '1': # quiz generate
        os.system("cls")
        try:
            quizGenerator.work()
        except Exception as e:
            print('默写纸生成失败！\n %s' %e)
    elif inp == '2': # word extract
        os.system("cls")
        try:
            outputToExcel.output()
        except Exception as e:
            print('高频词提取失败！\n %s' %e)
    else:
        print('输入不合法！请重新输入')
        continue
    print('任务已完成！鉴于您可能还有其他任务，程序将继续执行，直至输入0退出。')
