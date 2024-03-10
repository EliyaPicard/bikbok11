import pandas as pd
import calendar
from datetime import datetime

# filename = 'QuestionsReport-QVYPz.xlsx'

def hours(filename, price_per_mile, month):
    def year():
        if datetime.now().month == 1:
            return datetime.now().year - 1
        else:
            return datetime.now().year

    def price_per_hour(name):
        df_names = pd.read_excel("salary per worker.xlsx") ##
        if name in df_names["עובד"].values:
            salary = df_names.loc[df_names["עובד"] == name, "תעריף לשעה"].iloc[0]
            return salary
        else:
            return 40

    df = pd.read_excel(filename, sheet_name='Questions Report', header=4)
    tables = []
    start_row = 0
    for i, row in df.iterrows():
        if df.iloc[i, 0] == " " and i != start_row:
            tables.append(df.loc[start_row:i - 1])
            start_row = i + 3
    df = pd.merge(tables[0], tables[1], how='left',
                  left_on=['מחלקה', 'מיקום', 'התחלה', 'סיום', 'סך שעות', 'תאריך', 'שם עובד'],
                  right_on=['מחלקה', 'מיקום', 'התחלה', 'סיום', 'סך שעות', 'תאריך', 'שם עובד'],
                  suffixes=('_left_1', '_right_1'))
    df = pd.merge(df, tables[2], how='left',
                  left_on=['מחלקה', 'מיקום', 'התחלה', 'סיום', 'סך שעות', 'תאריך', 'שם עובד'],
                  right_on=['מחלקה', 'מיקום', 'התחלה', 'סיום', 'סך שעות', 'תאריך', 'שם עובד'],
                  suffixes=('_left_2', '_right_2'))
    df = pd.merge(df, tables[3], how='left',
                  left_on=['מחלקה', 'מיקום', 'התחלה', 'סיום', 'סך שעות', 'תאריך', 'שם עובד'],
                  right_on=['מחלקה', 'מיקום', 'התחלה', 'סיום', 'סך שעות', 'תאריך', 'שם עובד'],
                  suffixes=('_left_3', '_right_3'))
    print(df)
    df['ש"ע 100%'] = df['סך שעות'].apply(lambda x: x if x <= 9 else 9)
    df['ש"ע 125%'] = df['סך שעות'].apply(lambda x: 0 if x <= 9 else x - 9 if x <= 11 and x > 9 else 2)
    df['ש"ע 150%'] = df['סך שעות'].apply(lambda x: x - 11 if x >= 11 else 0)
    df = df.iloc[:, [0, 1, 2, 3, 4, 5, 6, 11, 12, 13, 7, 8, 9, 10]]
    df['יום'] = df['תאריך'].apply(lambda x: x.split()[0])
    df['יום'] = df['יום'].apply(lambda x: x.split('/')[0])
    df = df.iloc[:, [14, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]]
    df = df.iloc[:, [0, 11, 3, 1, 2, 3, 5, 6, 7, 8, 9, 10, 12, 13, 14]]
    df = df.iloc[:, [0, 3, 1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]]
    df['תאריך'] = df['תאריך'].apply(lambda x: x.split()[0])

    # rename the columns in the data frame

    # df
    df_by_workers = {}
    for i in df["שם עובד"].values:
        df1 = df[df["שם עובד"] == i]

        df1.loc[:, 'תאריך'] = pd.to_datetime(df1.loc[:, 'תאריך'], format='%d/%m/%Y')
        start_date = df1.loc[:, 'תאריך'].min().replace(day=1)
        end_date = calendar.monthrange(start_date.year, start_date.month)[1]
        end_date = start_date.replace(day=end_date)
        all_dates = pd.date_range(start=start_date, end=end_date, freq='D')
        new_df = pd.DataFrame({'תאריך': all_dates})

        df1['תאריך'] = pd.to_datetime(df1['תאריך'], format='%d/%m/%Y')
        new_df['תאריך'] = pd.to_datetime(new_df['תאריך'], format='%d/%m/%Y')
        merged_df = pd.merge(df1, new_df, on='תאריך', how='outer')

        # merge the new DataFrame with the original DataFrame
        # merged_df = pd.merge(df1, new_df, on='תאריך', how='outer')
        merged_df = merged_df.sort_values('תאריך')

        merged_df['יום'] = merged_df['תאריך'].apply(lambda x: str(x).split(' ')[0]).apply(
            lambda x: str(x).split('-'))
        merged_df['יום'] = merged_df['יום'].apply(lambda x: x[2])

        col_names = list(merged_df.columns)
        col_names[-1] = 'החזר נסיעות'
        merged_df.columns = col_names

        col_names = list(merged_df.columns)
        col_names[-2] = 'שינה בבית'
        merged_df.columns = col_names

        col_names = list(merged_df.columns)
        col_names[-3] = 'ארוחות'
        merged_df.columns = col_names

        col_names = list(merged_df.columns)
        col_names[2] = 'לקוח'
        merged_df.columns = col_names

        merged_df['ארוחות'] = merged_df['ארוחות'].apply(lambda x: 10 if x == "כן" else 0)
        merged_df['גילום ארוחות'] = merged_df.apply(
            lambda row: 0 if row['ארוחות'] == 0 else (
                100 if row['שינה בבית'] == "לא" else 40 if row['שינה בבית'] == "כן" else 0),
            axis=1
        )

        cols = ['ש"ע 100%', 'ש"ע 125%', 'ש"ע 150%', 'סך שעות', 'החזר נסיעות', 'ארוחות']

        # Calculate the sum of each column
        sums = merged_df[cols].sum(axis=0)

        # Insert a new row after the last row
        merged_df.loc[merged_df.index.max() + 1, cols] = sums
        merged_df.drop('תאריך', axis=1, inplace=True)
        merged_df.drop('שינה בבית', axis=1, inplace=True)

        total = {
            'ש"ע 150%': merged_df.iloc[:-1]['ש"ע 150%'].sum() * price_per_hour(i) * 1.5,
            'ש"ע 125%': merged_df.iloc[:-1]['ש"ע 125%'].sum() * price_per_hour(i) * 1.25,
            'ש"ע 100%': merged_df.iloc[:-1]['ש"ע 100%'].sum() * price_per_hour(i),
            'גילום ארוחות': merged_df['גילום ארוחות'].sum(),
            'ארוחות': merged_df.iloc[:-1]['ארוחות'].sum(),
            'החזר נסיעות': merged_df.iloc[:-1]['החזר נסיעות'].sum() * price_per_mile,

        }
        merged_df.loc[len(merged_df)] = total

        cell1 = None
        new_row1 = {'ש"ע 100%': cell1}
        merged_df.loc[len(merged_df)] = new_row1

        cell = None
        new_row = {'ש"ע 100%': cell}
        merged_df.loc[len(merged_df)] = new_row

        cell1 = None
        new_row1 = {'ש"ע 100%': cell1}
        merged_df.loc[len(merged_df)] = new_row1

        if i == "אברהם דוד חזן" or i == "מרדכי בנימין קפלן":
            cell2 = "משכורת קבועה של 9,000 שח ברוטו לחודש"
            new_row2 = {'סך שעות': cell2}
            merged_df.loc[len(merged_df)] = new_row2

        else:
            cell = merged_df.iloc[-4, 8] + merged_df.iloc[-4, 9] + merged_df.iloc[-4, 10] + merged_df.iloc[-4, 12]
            cell2 = cell
            new_row2 = {'ש"ע 100%': cell2}
            merged_df.loc[len(merged_df)] = new_row2

        if i == "נחמן סאפר" or i == "יואל גינסבורי":
            cell2 = 0
            new_row2 = {'ש"ע 100%': cell2}
            merged_df.loc[len(merged_df)] = new_row2
        else:
            cell2 = merged_df.iloc[-5, -3]
            new_row2 = {'ש"ע 100%': cell2}
            merged_df.loc[len(merged_df)] = new_row2

        cell2 = merged_df.iloc[-6, -1]
        new_row2 = {'ש"ע 100%': cell2}
        merged_df.loc[len(merged_df)] = new_row2

        if i == "אברהם דוד חזן" or i == "מרדכי בנימין קפלן":
            pass
        else:
            merged_df.iloc[-3, -7] = 'לתשלום'
        merged_df.iloc[-2, -7] = 'לפני הפחתה'
        merged_df.iloc[-2, -5] = 'של ארוחות'
        merged_df.iloc[-1, -7] = 'יש לגלם'
        merged_df.iloc[-1, -5] = 'עבור הארוחות'
        merged_df.iloc[-8, -8] = 'סה"כ'

        df_by_workers[i] = merged_df

    writer = pd.ExcelWriter(f'{month} {year()}.xlsx', engine='xlsxwriter')
    workbook = writer.book

    # Iterate over the keys and values of the dictionary and write each DataFrame to a separate sheet
    for sheet_name, df in df_by_workers.items():

        df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=4)
        worksheet = writer.sheets[sheet_name]

        cell_format3 = workbook.add_format({'bold': True, 'font_color': 'red', 'num_format': '₪ #.##0'})
        worksheet.write('C2', "חודש:")
        worksheet.write('D2', month)
        worksheet.write('E2', year())

        worksheet.write('C3', "שם:")
        worksheet.write('D3', sheet_name)
        worksheet.write('C4', "כתובת:")
        worksheet.write('G3', "תעריף", cell_format3)
        worksheet.write('M3', "תעריף", cell_format3)
        worksheet.write('G4', price_per_hour(sheet_name), cell_format3)
        worksheet.write('M4', price_per_mile, cell_format3)

        # Set up a conditional format to add borders to cells with values
        format = workbook.add_format({'border': 1})
        worksheet.conditional_format('A1:N45', {'type': 'formula',
                                                'criteria': 'LEN($A1)>0',
                                                'format': format})

        cell_format = workbook.add_format({'num_format': '₪ #.##0'})

        total_rows = worksheet.dim_rowmax + 1

        # Set the last six rows to the desired format
        last_six_rows = range(total_rows - 7, total_rows)
        for row in last_six_rows:
            worksheet.set_row(row, cell_format=cell_format)

        worksheet.right_to_left()
        worksheet.autofit()

    # Close the Excel writer object
    writer.close()
    return writer


# hours(filename,1,'Februray')