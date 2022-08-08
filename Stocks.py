#!/usr/bin/env python
# coding: utf-8

# # Project: Creating a Stock Dashboard

# ## Analyzing Stocks with Python and xlwings

# In[334]:


import xlwings as xw
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pandas_datareader import data 
import seaborn as sns
from statsmodels.formula.api import ols
plt.style.use("seaborn")


# # Linking Excel File to python
wb = xw.Book("Stocks.xlsx")

def main():
    wb = xw.Book.caller()
    
    # Defining excel sheet names 
    db_s = wb.sheets[0]
    prices_s = wb.sheets[1]
    
    # Generating Pandas DataFrame and acquiring Portfolio and analysis details
    symbol_array = db_s.range("M2").options(pd.DataFrame, expand = "table",index=False).value
    symbol = symbol_array.Ticker_Symbols
    start = db_s.range("F7").value
    end = db_s.range("I7").value
    benchmark = db_s.range("M3").value
    freq = db_s.range("K39").value


    # Reading stock price data for the portfolio and storing in DataFrame "df" 
    df = data.DataReader(name = symbol, data_source = "yahoo", start = start, end = end).Close
   
    # Calculating total portfolio value and storing in the new column "Port"
    df['Port'] = 0
    for i in range(1,len(symbol_array.index)):
        df.iloc[:,i]=df.iloc[:,i].mul(symbol_array.iloc[i,1])
        df['Port']= df['Port']+df.iloc[:,i]

    df.rename(columns = {benchmark:benchmark.replace("^", "")}, inplace = True)
    benchmark = benchmark.replace("^", "")

    # Storing KPI for the Portfolio

    #first = df.iloc[0,len(df.columns)-1]
    #high = df.iloc[:, len(df.columns)-1].max()
    #low = df.iloc[:, len(df.columns)-1].min()
    #last = df.iloc[-1, len(df.columns)-1]

    first = df['Port'].iloc[0]
    high = df['Port'].max()
    low = df['Port'].min()
    last = df['Port'].iloc[-1]

    total_change = last / first - 1
    db_s.range("H12").options(transpose = True).value = [first, high, low, last, total_change]

    # Plotting a chart for displaying Portfolio trend

    chart = plt.figure(figsize = (20,10))
    df['Port'].plot(fontsize = 15)
    plt.title("Portfolio", fontsize = 20)
    plt.xlabel("Date", fontsize = 15)
    plt.ylabel("Stock Price", fontsize = 15)
    plt.annotate(int(first), xy = (df.index[0],first-50), fontsize = 17)
    plt.annotate(int(last), xy = (df.index[-1],last-50), fontsize = 17)
   
    # Plotting the Portfolio chart in Excel

    db_s.pictures.add(chart, name = "Chart", update = True, 
                       left = db_s.range("C8").left, 
                       top = db_s.range("C8").top,
                       scale = 2)

    # Normalizing the stock prices for comparison 

    norm = df.div(df.iloc[0]).mul(100)

    # Storing KPI for the Normalized Portfolio

    first_n = norm['Port'].iloc[0]
    high_n = norm['Port'].max()
    low_n = norm['Port'].min()
    last_n = norm['Port'].iloc[-1]
    last_b_n = norm[benchmark].iloc[-1]

    total_change_n = last_n / first_n - 1

    # Plotting a chart to compare the Portfolio with the benchmark

    chart2 = plt.figure(figsize = (16, 9))
    norm[df.columns[0]].plot(fontsize = 15)
    norm['Port'].plot(fontsize = 15)
    plt.title("Port" + " vs. " + benchmark, fontsize = 20)
    plt.xlabel("Date", fontsize = 15)
    plt.ylabel("Normalized Price (Base 100)", fontsize = 15)
    plt.annotate(int(first_n), xy = (norm.index[0],first_n-5), fontsize = 17)
    plt.annotate(int(last_n), xy = (norm.index[-1],last_n-5), fontsize = 17)
    plt.annotate(int(last_b_n), xy = (norm.index[-1],last_b_n-5), fontsize = 17)
    plt.legend(fontsize = 20)
   
    # Plotting the Normalized Portfolio chart in Excel

    db_s.pictures.add(chart2, name = "Chart2", update = True, 
                       left = db_s.range("C21").left, 
                       top = db_s.range("C21").top + 10,
                       scale = 0.2)


    ret  = df.resample(freq).last().dropna().pct_change().dropna()
    
    # Plotting Reg Plot 
    
    chart3 = plt.figure(figsize = (12.5, 10))
    sns.regplot(data = ret, x = ret[benchmark], y = ret['Port'])
    plt.title('Portfolio' + " vs. " + benchmark, fontsize = 20)
    plt.xlabel(benchmark + " Returns", fontsize = 15)
    plt.ylabel('Portfolio' + " Returns", fontsize = 15)
    
    # Plotting the Returns reg plot in Excel

    db_s.pictures.add(chart3, name = "Chart3", update = True, 
                       left = db_s.range("C40").left, 
                       top = db_s.range("C40").top,
                       scale = 0.47)

    # Computing linear reg model

    model = ols( 'Port' + "~" + benchmark, data = ret)
    results = model.fit()
    
    # Storing the regression model KPI

    obs = len(ret)
    corr_coef = ret.corr().iloc[0,1]
    beta = results.params[1]
    r_sq = results.rsquared
    t_stat = results.tvalues[1]
    p_value = results.pvalues[1]
    conf_left = results.conf_int().iloc[1,0]
    conf_right = results.conf_int().iloc[1,1]
    interc = results.params[0]

    # Writing the KPI in Excel
    regr_list = [obs, corr_coef, beta, r_sq, t_stat, p_value, conf_left, conf_right, interc]
    db_s.range("K41").options(transpose = True).value = regr_list

    # Writing Portfolio stock values in Excel

    prices_s.range("A1").expand().clear_contents()
    prices_s.range("A1").value = df