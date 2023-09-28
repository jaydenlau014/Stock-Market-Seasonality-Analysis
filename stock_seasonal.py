# -*- coding: utf-8 -*-
"""
Created on Thu Sep 04 19:12:04 2023

@author: Jayden Lau
"""

#%%
import pandas as pd
import numpy as np
import scipy
import requests
import itertools
import time

#%%  Get stock data
# APPLE, MICROSOFT, GOOGLE, TESLA, AMAZON stock
stock_symbols_list = ['AAPL','MSFT','GOOG','TSLA','AMZN']

#%% Winning Rate
# Function to change format
def stock_df_format(x, rolling=False,rolling_time=None):
    final_stock_df, stock_df = pd.DataFrame(), pd.DataFrame()
    if rolling == True:
        for i in range(len(x.index)):
        
            stock_df = pd.concat([x, x.iloc[:rolling_time]], axis=0)
            stock_df = stock_df.iloc[i:i+rolling_time]
            stock_df = stock_df.sum(axis=0).to_frame().T
            
            if i+ rolling_time <= len(x.index):
                stock_df = stock_df.rename(index={0:'{} to {}'.format(i+1, i+rolling_time)})
            else:
                stock_df = stock_df.rename(index={0:'{} to {}'.format(i+1, i+rolling_time-len(x.index))})
        
            stock_df['Mean'], stock_df['Std'] = np.mean(stock_df, axis=1), np.std(stock_df, axis=1, ddof=1)
            stock_df['Z_score'] = (0 - stock_df['Mean'])/stock_df['Std']
            stock_df['Lose'] = np.round(scipy.stats.norm.cdf(stock_df['Z_score'])*100, 2)
            stock_df['Win'] = 100 - stock_df['Lose']
            final_stock_df = pd.concat([final_stock_df, stock_df], axis=0)
            
    elif rolling == False:
        stock_df = x
        stock_df['Mean'], stock_df['Std'] = np.mean(stock_df, axis=1), np.std(stock_df, axis=1, ddof=1)
        stock_df['Z_score'] = (0 - stock_df['Mean'])/stock_df['Std']
        stock_df['Lose'] = np.round(scipy.stats.norm.cdf(stock_df['Z_score'])*100, 2)
        stock_df['Win'] = 100 - stock_df['Lose']
        final_stock_df = pd.concat([final_stock_df, stock_df], axis=0)
            
    return(final_stock_df)

        
# Function for highlighter
def top3_highlighter_green(x):
    style_lt = "background-color:"
    style_gt = "background-color: lightgreen"
    gt_mean = x >= x.nlargest(3).iloc[2]
    return [style_gt if i else style_lt for i in gt_mean]

def below3_highlighter_red(x):
    style_lt = "background-color:"
    style_gt = "background-color: red"
    gt_mean = x >= x.nlargest(3).iloc[2]
    return [style_gt if i else style_lt for i in gt_mean]

#%%
week_win_df = pd.DataFrame()
week_volatility_df = pd.DataFrame()

month_win_df = pd.DataFrame()
month_volatility_df = pd.DataFrame()

for symbol in stock_symbols_list:
    url = f' https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol={symbol}&outputsize=full&apikey=demo&outputsize=full'
    r = requests.get(url).json()
    df_0 = pd.DataFrame.from_dict(r['Time Series (Daily)']).T.astype('float').reset_index(names='timestamp')
    df_0.columns = df_0.columns.str.replace(r'\d+.\s+','', regex=True)

    df_1 = df_0.copy()
    df_1['timestamp'] = pd.to_datetime(df_1['timestamp'])
    df_1['year'] = df_1['timestamp'].dt.year
    df_1['month'] = df_1['timestamp'].dt.month
    df_1['week']= df_1['timestamp'].dt.isocalendar().week
    df_1['day_of_week'] = df_1['timestamp'].dt.dayofweek
    df_1['day_changes'] = df_1['close'].pct_change(periods=-1)*100
    df_1['volatility_changes'] = np.abs((df_1['high'] - df_1['low'])/df_1['low'])*100

    
    # Weekly Changes (default)
    
    week_df0 = df_1.pivot_table(index='week',columns='year',values='day_changes',
                               aggfunc='sum', fill_value=0)
    
    # Change week_df0 format
    week_df0 = stock_df_format(week_df0)
    # Concat all week_df0 for each fx
    week_win_df = pd.concat([week_win_df, week_df0['Win']],
                            names='{}'.format(symbol), axis=1)
    
 
    # Weekly Volatility
    week_df1 = df_1.pivot_table(index='week',columns='year',values='volatility_changes',
                               aggfunc='sum', fill_value=0)
    week_df1 = np.mean(week_df1, axis=1).to_frame().rename(columns={0:'volatility_changes'})
    week_volatility_df = pd.concat([week_volatility_df, week_df1['volatility_changes']],
                            names='{}'.format(symbol), axis=1)
    
    # Combine both week_df0 and week_df1
    week_df = pd.concat([week_df0, week_df1], axis=1)
    
    # Style
    week_df_style = week_df.style.apply(top3_highlighter_green, subset='Win')\
        .apply(below3_highlighter_red, subset='Lose').format({'Win':'{:.2f}%', 'Lose':'{:.2f}%'})

    
    
    
    # Monthly Changes (default)
    month_df0 = df_1.pivot_table(index='month',columns='year',values='day_changes', 
                              aggfunc='sum').dropna(axis=1)

    # Change monthly format
    month_df0 = stock_df_format(month_df0)

    month_win_df = pd.concat([month_win_df, month_df0['Win']],
                             names='{}'.format(symbol), axis=1)


    # Monthly Volatility
    month_df1 = df_1.pivot_table(index='month',columns='year',values='volatility_changes',
                               aggfunc='sum', fill_value=0)
    month_df1 = np.mean(month_df1, axis=1).to_frame().rename(columns={0:'volatility_changes'})
    month_volatility_df = pd.concat([month_volatility_df, month_df1['volatility_changes']],
                            names='{}'.format(symbol), axis=1)
    
    # Combine both week_df0 and week_df1
    month_df = pd.concat([month_df0, month_df1], axis=1)
    
    # Monthly Style
    month_df_style= month_df.style.apply(top3_highlighter_green, subset='Win')\
        .apply(below3_highlighter_red, subset='Lose').format({'Win':'{:.2f}%', 'Lose':'{:.2f}%'})
    
    # Write Excel file
    with pd.ExcelWriter('./{}.xlsx'.format(symbol), engine='openpyxl') as writer:
        df_0.to_excel(writer, sheet_name = 'Initial')
        df_1.to_excel(writer, sheet_name='Post_process')
        
        week_df_style.to_excel(writer, sheet_name ='Weekly')   
        
        month_df_style.to_excel(writer, sheet_name='Monthly')
        
    time.sleep(30)
#%% Winning Rate (Cont)
week_win_df.columns = stock_symbols_list
month_win_df.columns = stock_symbols_list

week_volatility_df.columns = stock_symbols_list
month_volatility_df.columns = stock_symbols_list

with pd.ExcelWriter('./stock_summary.xlsx', engine='openpyxl') as writer:
    for i in range(len(week_win_df.index)):
        # Write week full df
        each_week_win_df = week_win_df.iloc[i].sort_values(ascending=False).to_frame().T
        each_week_win_df.to_excel(writer, sheet_name='Week', startcol= 1, startrow = 3*i+1)
        
        each_week_volatility_df = week_volatility_df.iloc[i].sort_values(ascending=False).to_frame().T
        each_week_volatility_df.to_excel(writer, sheet_name='Week Volatility', startcol= 1, startrow = 3*i+1)
        
        
    for i in range(len(month_win_df.index)):
        # Write month full df
        each_month_win_df= month_win_df.iloc[i].sort_values(ascending=False).to_frame().T
        each_month_win_df.to_excel(writer, sheet_name='Month', startcol= 1, startrow= 3*i+1)
        
        each_month_volatility_df = month_volatility_df.iloc[i].sort_values(ascending=False).to_frame().T
        each_month_volatility_df.to_excel(writer, sheet_name='Month Volatility', startcol= 1, startrow = 3*i+1)
    









        
