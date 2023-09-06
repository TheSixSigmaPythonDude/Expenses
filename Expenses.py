import pandas as pd
import numpy as np
import glob
from datetime import datetime

def cleanExpenseFile(expenseFile):
    expenseFile = pd.read_excel(expenseFile).iloc[3:,1:].reset_index()
    expenseFile.columns = expenseFile.iloc[0,:]
    expenseFile.drop([0], axis=0, inplace=True)
    expenseFile = expenseFile[['Date','Account','Description','Category','Amount']]
    expenseFileFinal = expenseFile[expenseFile['Account'].notna()]

    return expenseFileFinal

def extractMonth(dateObj):
    try:
        # Extract the month from the datetime object and return it as a string
        month = dateObj.strftime('%B')  # 'B' represents the full month name
        return month
    except AttributeError:
        return "Invalid input. Please provide a valid datetime object."

if __name__ == '__main__':

    df = pd.DataFrame()
    expenseFiles = glob.glob(r'C:/Users/shing/OneDrive/Desktop/Expenses/*')

    for x in expenseFiles:
        df = pd.concat([df,cleanExpenseFile(x)])

    df['Month'] = df['Date'].apply(extractMonth)

    for month in range(12):
    grandAverage = df['Amount'].groupby(df['Category']).mean()
    monthlyAverage = grandAverage/len(df['Month'].value_counts())
    print(grandAverage)
