# -*- coding: utf-8 -*-
"""
Created on Wed Jul 22 14:10:42 2020

@author: MI
"""

import pandas as pd
import matplotlib.pyplot as plt
import datetime as dt

data = pd.read_excel('C:\FUDAN\实习\绘图\\test.xlsx')

df = data.iloc[1:, 0:2].dropna().rename(columns = {'Unnamed: 0':'date', '水泥出货率':'ratio'})

df['year'] = df.date.map(lambda x:x.year)
df['date'] = df.date.map(lambda x:dt.datetime.strptime(str(x.month) + '-' + str(x.day), '%m-%d'))

df = df.set_index(['year', 'date'])

df = df.sort_index()

fig, ax = plt.subplots(1, 1, figsize = (14, 7))

for i in range(2015, 2021):
    df_year = df.query('year==%i' % i).iloc[::-1].reset_index(drop = False).drop(columns = ['year']).set_index(['date']).rename(columns = {'ratio':'%i' % i})
    df_year.plot(ax = ax, linewidth = 2.5)

plt.legend(loc = 'upper center', ncol = len(range(2015, 2021)))

xdate = []
xreal = []

for i in range(12):
    xdate.append(dt.datetime.strptime('1900-%i-1' % (i + 1), '%Y-%m-%d'))
    xreal.append('%i/1' % (i + 1))
    
plt.xticks(xdate, xreal)
