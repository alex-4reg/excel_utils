import pandas as pd
from excel_file_preprocesssor import preprocessor

invoice_df = pd.concat(pd.read_excel(preprocessor("./input/materials.xlsx"), sheet_name=None,
                                     names=['No.', 'name', 'unit of measure', 'quantity', 'price', 'cost of goods',
                                            'value with tax',
                                            'remainder', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10',
                                            'invoice No', 'date']))
pd.set_option('display.max_columns', None)
invoice_df.dropna(subset=['price', 'name'], inplace=True)

invoice_df = invoice_df.reindex(
    columns=['invoice No', 'date', 'No.', 'name', 'unit of measure', 'quantity', 'price', 'cost of goods',
             'value with tax', 'remainder', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10'])
invoice_df.drop(invoice_df.index[invoice_df['invoice No'] == 'invoice No'], inplace=True)
invoice_df.reset_index(inplace=True, drop=True)

invoice_df[
    ['quantity', 'price', 'cost of goods', 'value with tax', 'remainder', '1', '2', '3', '4', '5', '6', '7', '8', '9',
     '10']] = invoice_df[
    ['quantity', 'price', 'cost of goods', 'value with tax', 'remainder', '1', '2', '3', '4', '5', '6', '7', '8', '9',
     '10']].fillna(0)

invoice_df['No.'] = invoice_df['No.'].astype("int64")
invoice_df[['date']] = invoice_df[['date']].astype("datetime64")
invoice_df[['invoice No']] = invoice_df[['invoice No']].astype("str")

invoice_df[
    ['quantity', 'price', 'cost of goods', 'value with tax', 'remainder', '1', '2', '3', '4', '5', '6', '7', '8', '9',
     '10']] = \
    invoice_df[
        ['quantity', 'price', 'cost of goods', 'value with tax', 'remainder', '1', '2', '3', '4', '5', '6', '7', '8',
         '9', '10']].astype(
        "float64")

actual_invoice_No = ""
for index, row in invoice_df.iterrows():
    if row['invoice No'] != "nan":
        actual_invoice_No = str(row['invoice No'])
        # print(actual_invoice_No)
    else:
        invoice_df.at[index, 'invoice No'] = actual_invoice_No

invoice_df.insert(loc=9, column='total spent', value=0)
total_spent = invoice_df['total spent'] = invoice_df.iloc[:, 11:].sum(axis=1)
invoice_df['remainder'] = invoice_df['quantity'] - invoice_df['total spent']
invoice_df['cost of goods'] = invoice_df['quantity'] * invoice_df['price']
invoice_df['value with tax'] = invoice_df['cost of goods'] * 1.20
invoice_df.insert(loc=11, column='cost of reminder', value=0)
cost_of_reminder = invoice_df['cost of reminder'] = invoice_df['remainder'] * invoice_df['price']

agg_functions = {'invoice No': 'first', 'date': 'first', 'No.': 'sum', 'unit of measure': 'first', 'quantity': 'sum',
                 'price': 'first', 'cost of goods': 'sum', 'value with tax': 'sum', 'remainder': 'sum',
                 'cost of reminder': 'sum', '1': 'sum', '2': 'sum', '3': 'sum', '4': 'sum', '5': 'sum', '6': 'sum',
                 '7': 'sum', '8': 'sum', '9': 'sum', '10': 'sum'}

df_new = invoice_df.groupby(invoice_df['name']).aggregate(agg_functions)
print(df_new)

invoice_df.to_excel("./output/materials_all_sheets_merged.xlsx")
df_new.to_excel("./output/materials_grouped_by_name.xlsx")
