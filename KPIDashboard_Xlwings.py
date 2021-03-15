#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import xlwings as xw
def main():
    wb=xw.Book(r"C:\users\jesko\input.xlsx")
    wbNew=xw.Book(r"C:\users\jesko\KPIDashboard_Xlwings.xlsm")
    # more specifically, we want to work with the sheet Input
    sht = wb.sheets['Input']
    # and to be even more precise, we are only using the cells which belong to the table's range starting with cell A1
    dfInput = sht.range('A1').options(pd.DataFrame, header=1, index=False, expand='table').value
    # for training purpose, we add an easy OEE calculation to our dataframe
    dfInput['OEECalc']=(dfInput['Availability']+dfInput['Performance'] + dfInput['Quality']) /3
    shtout = wbNew.sheets['InputNew']
    shtout.range('B4').options(pd.DataFrame, index=False).value = dfInput
@xw.func
def hello(name):
    return f"Hello {name}!"
if __name__ == "__main__":
    xw.Book("input.xlsx").set_mock_caller()
    main()


# In[ ]:




