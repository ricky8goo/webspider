# 导入所需的库
import os
import tkinter as tk
import tkinter.ttk as ttk
import feedparser
import requests
import pandas as pd
from bs4 import BeautifulSoup
import numpy as np
from datetime import datetime
import re
import random
import string
from openpyxl import Workbook

# 定义按钮事件函数
def button_event():
    # RSS 源的 URL
    rss_url = 'https://www.notebookcheck.net/RSS-Feed-Notebook-Reviews.8156.0.html'

    # 通过 feedparser 和 requests 获取 RSS 数据
    rss = feedparser.parse(rss_url)
    r2 = requests.get(rss_url)

    # 使用 BeautifulSoup 解析 HTML
    soup2 = BeautifulSoup(r2.text,"html.parser")

    # 选择所有的 pubDate 标签
    sel = soup2.select("pubDate")

    # 保存和加载 pubDate 数据
    newtime = sel
    np.savetxt("result.txt", newtime, fmt = '%s')
    file = open('result.txt')
    lines = file.readlines()

    # 获取当前的日期和时间
    dt = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # 查找包含用户选择的月份的行
    b = [i for i in lines if mycombobox.get() in i]

    # 初始化保存索引的列表
    save_index = []

    # 如果找到了包含用户选择的月份的行
    if b:
        for j in range(0, len(b)):
            # 获取行的索引
            save_index = lines.index(b[j])

            # 获取链接并解析 HTML
            r = requests.get(rss.entries[lines.index(b[j])]['link'])
            soup = BeautifulSoup(r.text,"html.parser")

            # 生成随机的文件名
            salt = ''.join(random.sample(string.ascii_letters + string.digits, 8))
            filename = salt + '.xlsx'

            # 初始化 DataFrame
            basic_INF_data = pd.DataFrame()
            R23M_data = pd.DataFrame()
            R23S_data = pd.DataFrame()
            R23MS_data_data = pd.DataFrame()
            R20M_data = pd.DataFrame()
            R20S_data = pd.DataFrame()
            R20MS_data_data = pd.DataFrame()
            Dmark_data = pd.DataFrame()

            # 创建 ExcelWriter 对象
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:

                # 查找和保存不同的数据
                # ...

    # 如果没有找到包含用户选择的月份的行
    else:
        # 在 GUI 中显示消息，并在控制台打印消息
        w = tk.Label(root, text="該月份資料不存在")
        w.pack()
        print("該月份資料不存在")

# 创建 tkinter 的 root 对象
root = tk.Tk()

# 设置窗口的标题和大小
root.title('my window')
root.geometry('300x200')

# 创建和设置下拉列表框
comboboxList = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
mycombobox = ttk.Combobox(root, state='readonly')
mycombobox['values'] = comboboxList
mycombobox.pack(pady=10)
mycombobox.current(0)

# 创建和设置按钮
buttonText =  tk.StringVar()
buttonText.set('button')
tk.Button(root, textvariable=buttonText, command=button_event).pack()

# 启动 tkinter 的主循环
root.mainloop()
