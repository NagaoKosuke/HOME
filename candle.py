import math
import eikon as ek  # the Eikon Python wrapper package
ek.set_app_key('0ed6a35e0937415eab446d3375bca7cf671d6b4c')
import numpy as np  # NumPy
import pandas as pd  # pandas
import cufflinks as cf  # Cufflinks
import configparser as cp
# import scipy.optimize as sco  # optimization routines
cf.set_config_file(offline=True, theme="white")
#Font Setting
from matplotlib.font_manager import FontProperties
import sys
if sys.platform.startswith('win'):
    FontPath= 'C:\\Windows\\Fonts\\meiryo.ttc'
elif sys.platform.startswith('darwin'):
    FontPath= '/System/Library/Fonts/ヒラノギ角ゴシック W4.ttc'
elif sys.platform.startswith('linux'):
    FontPath= '/usr/share/fonts/truetype/takao-gothic/TakaoExGothic.ttc'
jpfont = FontProperties(fname = FontPath)

import matplotlib.pyplot as plt
import matplotlib as mpl
import os
mpl_dirpath = os.path.dirname(mpl.__file__)
# デフォルトの設定ファイルのパス
default_config_path = os.path.join(mpl_dirpath, 'mpl-data', 'matplotlibrc')
# カスタム設定ファイルのパス
custom_config_path = os.path.join(mpl.get_configdir(), 'matplotlibrc')

import plotly.graph_objects as go
import chart_studio.plotly as py
import plotly.figure_factory as ff
import datetime
import xlwings as xw

now = datetime.datetime.today()
delta = datetime.timedelta(weeks=52)
sdate = now-delta

print_ric = ['米国10年債(%)','S&P500($)','日本10年債(%)','日経平均(円)',
            'ドイツ10年債(%)','DAX(€)','WTI原油先物($/BBL)','ドル円']

rics = [
    'US10YT=RR',  # Apple stock
    '.SPX',
    'JP10YT=RR',  # Amazon stock
    '.N225',
    'DE10YT=RR', 
    '.DAX',
    'CLc1',  # Gold ETF
    'JPY='  # USD/JPY exchange rate
]


fields=[
    'Open',
    'High',
    'Low',
    'Close'
]

meigara =[
    ('US10YT=RR','米国10年債(%)'),  # Apple stock
    ('.SPX','S&P500($)'),
    ('JP10YT=RR','日本10年債(%)'),  # Amazon stock
    ('.N225','日経平均(円)',),
    ('DE10YT=RR','ドイツ10年債(%)'), 
    ('.DAX','DAX(€)'),
    ('CLc1','WTI原油先物($/BBL)'),  # Gold ETF
    ('JPY=','ドル円')  # USD/JPY exchange rate
]


data = ek.get_timeseries(rics,  # the RICs
                         fields=fields,  # the required fields
                         start_date=sdate.strftime('%Y-%m-%d'),  # start date
                         end_date=now.date().strftime('%Y-%m-%d'))  # end date

from plotly import tools
import plotly.offline as po
import plotly.io as pio

for v,char in meigara:
    fig = go.Figure(data=[go.Candlestick(x=data.index,
                    open=data[v]['OPEN'],
                    high=data[v]['HIGH'],
                    low=data[v]['LOW'],
                    close=data[v]['CLOSE'])],
    )
    fig.layout.update({'title':char})
    fig.update_xaxes(
        rangeslider_visible=False,
        rangeselector=dict(
            buttons=list([
                dict(count=1,label='1m',step='month',stepmode='backward'),
                dict(count=6,label='6m',step='month',stepmode='backward'),
                dict(count=1,label='YTD',step='year',stepmode='todate'),
                dict(count=1,label='1y',step='year',stepmode='backward'),
                dict(step='all')
            ])
        ),        
        rangebreaks=[
            dict(pattern='day of week',bounds=[6,1])
        ]
    )
    
    sht = xw.Book('candle.xlsx').sheets[0]
    sht.pictures.add(fig,name=v,update=True)

    # po.plot(fig ,filename=v+'.html')
    # pio.write_image(fig,v+'.pdf')