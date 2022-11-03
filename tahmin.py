#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import xlwings as xw
import random
import numpy as np
from datetime import datetime
from statsmodels.tsa.arima.model import ARIMA
from xlwings import Range, constants

wb = xw.Book("Proje_Python.xlsm")
sheet = wb.sheets["Materials"]
sheet2 = wb.sheets["Order_Inventory Plan"]
sheet3 = wb.sheets["Forecasting"]
sheet4 = wb.sheets["Forecast Results"]
sheet2.clear()
sheet4.clear()

period = len(np.array(sheet3.range("B2:XFD2").value))-len(np.array(sheet3.range("B2:XFD2").value)
                                                          [np.array(sheet3.range("B2:XFD2").value) == None])
number_of_materials_F = np.count_nonzero(sheet3.range("A2:A10000").value)
def arima(a,b,c):
    history = [x for x in train]
    predictions=[]
    for t in range(len(test)):
        model = ARIMA(history, order=(a,b,c))
        model_fit = model.fit()
        output = model_fit.forecast()
        yhat = output[0]
        predictions.append(yhat)
        obs = test[t]
        history.append(obs)
    return predictions
def arima_mape(a,b,c):
    return np.mean(100*(abs(arima(a,b,c)-test)/ test))
def min_mape():
    mape=[]
    sa=[]
    bs=[]
    cs=[]
    for a in range(0,4):
        for b in range(0,2):
            for c in range(0,2):
                sa.append(a)
                bs.append(b)
                cs.append(c)
                mape.append(arima_mape(a,b,c))
    mape_table=pd.DataFrame({"A": sa, "B":bs, "C":cs,
                        "MAPE value": mape})
    min_mape = mape_table.sort_values(by="MAPE value",ascending=True).reset_index().head(5)
    return min_mape.loc[0]["A"],min_mape.loc[0]["B"],min_mape.loc[0]["C"]

D_F = np.array(sheet3.range((2,2),(number_of_materials_F+1,period+1)).value)
D_F = D_F.astype(int)

for i in np.arange(number_of_materials_F):
    sheet4.range((1,2+i)).value = sheet3.range((2+i,1)).value
    sheet4.range((1,2+i)).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
    sheet4.range((1,2+i)).api.Borders.LineStyle = 1
    sheet4.range((1,2+i)).api.Font.Bold = True
for i in range(D_F.shape[0]):
    size = int(len(D_F[i])*(2/3))
    train, test = D_F[i][0:size], D_F[i][size:len(D_F[i])]
    a=[]
    b=[]
    c=[]
    [a,b,c]=min_mape()
    predictions=arima(a,b,c)
    for t in range(len(test)):
        sheet4.range((2+t,i+2)).value = predictions[t]
        sheet4.range((2+t,i+2)).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
        sheet4.range((2+t,i+2)).api.Borders.LineStyle = 1
        sheet4.range((2+t,i+2)).api.Interior.ColorIndex = 24
for i in np.arange(len(test)):
    sheet4.range((2+i,1)).value = i+1
    sheet4.range((2+i,1)).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
    sheet4.range((2+i,1)).api.Borders.LineStyle = 1
    sheet4.range((2+i,1)).api.Font.Bold = True

