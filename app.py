import pandas as pd
import re
from functions import make_headers, convert_currencies, color_sum_headers

df = pd.read_csv('data.csv')

needsDf = pd.DataFrame(columns=['Source', 'Date', 'Amount'])
wantsDf = pd.DataFrame(columns=['Source', 'Date', 'Amount'])
savingDf = pd.DataFrame(columns=['Source', 'Date', 'Amount'])
totalDf = pd.DataFrame(columns=['Category', 'Expenses', 'Fund', 'Rest'])

cellRange = {'A1:C1':'Needs','D1:F1':'Wants','G1:I1':'Saving'}
dataFrames = [needsDf,wantsDf,savingDf]
count = 0

for key in cellRange:
    for index,row in df.iterrows():
        if cellRange[key] in re.sub(r'\([^)]*\)', '', row['Account']):
            df.loc[index, 'Account'] = cellRange[key]
            new_row = {'Source': row['Name'], 'Date': row['Created time'], 'Amount': row['Formulation']}
            dataFrames[count] = pd.concat([dataFrames[count], pd.DataFrame(new_row, index=[0])], ignore_index=True)
            if row['Formulation'] > 0:
                new_row = {'Category': cellRange[key], 'Fund': row['Formulation']}
                totalDf = pd.concat([totalDf, pd.DataFrame(new_row, index=[0])], ignore_index=True)
            else:
                new_row = {'Category': cellRange[key], 'Expenses': abs(row['Formulation'])}
                totalDf = pd.concat([totalDf, pd.DataFrame(new_row, index=[0])], ignore_index=True)
    count += 1

totalDf = totalDf.fillna(0)
totalDf['Expenses'] = totalDf['Expenses'].astype(int)
totalDf['Fund'] = totalDf['Fund'].astype(int)
sum_df = totalDf.groupby('Category').agg({'Expenses': 'sum', 'Fund': 'sum', 'Rest': 'sum'}).reset_index()
sum_df['Rest'] = sum_df['Fund'] - sum_df['Expenses']
writer = pd.ExcelWriter('output.xlsx', engine='openpyxl')

dataFrames[0].to_excel(writer, sheet_name='Sheet1', index=False, startrow=1, startcol=0)
dataFrames[1].to_excel(writer, sheet_name='Sheet1', index=False, startrow=1, startcol=3)
dataFrames[2].to_excel(writer, sheet_name='Sheet1', index=False, startrow=1, startcol=6)
sum_df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=1, startcol=10)

workbook = writer.book
worksheet = workbook['Sheet1']

color_sum_headers(sum_df,worksheet)
# convert_currencies(worksheet)
make_headers(cellRange,worksheet)

writer._save()