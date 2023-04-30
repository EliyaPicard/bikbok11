from flask import Flask, request, render_template, send_file, url_for, jsonify,flash
import pandas as pd
import numpy as np
import xlsxwriter
import calendar
from datetime import datetime
import io, os, sys

if getattr(sys, 'frozen', False):
    template_folder = os.path.join(sys._MEIPASS, 'templates')
    app = Flask(__name__, template_folder=template_folder)
else:
    app = Flask(__name__)




@app.route('/')
def home1():
    return render_template("index.html")

@app.route('/home')
def home():
    return render_template("home.html")

@app.route('/hours')
def hours():
    return render_template('hours.html')

@app.route('/query')
def home2():
    return render_template("query.html")

@app.route('/update_files')
def home3():
    return render_template("update_files.html")

@app.route('/calculator', methods=['POST'])
def calc():
    filename = request.files['file']
    price_per_mile = float(request.form['price'])
    month = request.form['month']

    def hours(filename,price_per_mile, month):
        def year():
            if datetime.now().month == 1:
                return datetime.now().year - 1
            else:
                return datetime.now().year

        def price_per_hour(name):
            df_names = pd.read_excel("salary per worker.xlsx")
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
                      right_on=['מחלקה', 'מיקום', 'התחלה', 'סיום', 'סך שעות', 'תאריך', 'שם עובד'])
        df = pd.merge(df, tables[2], how='left',
                      left_on=['מחלקה', 'מיקום', 'התחלה', 'סיום', 'סך שעות', 'תאריך', 'שם עובד'],
                      right_on=['מחלקה', 'מיקום', 'התחלה', 'סיום', 'סך שעות', 'תאריך', 'שם עובד'])
        df = pd.merge(df, tables[3], how='left',
                      left_on=['מחלקה', 'מיקום', 'התחלה', 'סיום', 'סך שעות', 'תאריך', 'שם עובד'],
                      right_on=['מחלקה', 'מיקום', 'התחלה', 'סיום', 'סך שעות', 'תאריך', 'שם עובד'])
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

            # merge the new DataFrame with the original DataFrame
            merged_df = pd.merge(df1, new_df, on='תאריך', how='outer')
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
            merged_df['גילום ארוחות'] = merged_df['שינה בבית'].apply(
                lambda x: 100 if x == "לא" else 40 if x == "כן" else 0)

            cols = ['ש"ע 100%', 'ש"ע 125%', 'ש"ע 150%', 'סך שעות', 'החזר נסיעות','ארוחות']

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
                'החזר נסיעות': merged_df.iloc[:-1]['החזר נסיעות'].sum()*price_per_mile,

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
                cell = merged_df.iloc[-4, 8] + merged_df.iloc[-4, 9] + merged_df.iloc[-4, 10]+merged_df.iloc[-4, 12]
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
        writer.save()
        return writer
    output = hours(filename,price_per_mile, month)  # modify this line
    return send_file(output)

@app.route('/result', methods=['POST'])
def result():
    filename = request.files['file']
    price = request.form.get('price')
    if price:
        price = float(price)
    else:
        price = 0.0

    magnom = request.form['magnum_price']
    if magnom:
        magnom = float(magnom)
    else:
        magnom = 0.0

    mishtachim = request.form['mishtachim_price']
    if mishtachim:
        mishtachim = float(mishtachim)
    else:
        mishtachim = 0.0

    karton = request.form['karton_price']
    if karton:
        karton = float(karton)
    else:
        karton = 0.0

    maarach = request.form['maarach_price']
    if maarach:
        maarach = float(maarach)
    else:
        maarach = price

    Yekev_name = request.form['winery']
    if not Yekev_name:
        Yekev_name = '0'

    month = request.form['month']
    if not month:
        month = '0'

    def recipt(filename, price, magnom, mishtachim, Yekev_name, karton, maarach, month):
        xlsx = pd.ExcelFile(filename)
        sheet_names = xlsx.sheet_names
        sheet_names = [sheet for sheet in sheet_names if '.' in sheet]
        df_dict = {}
        for i in sheet_names:
            try:
                df1 = pd.read_excel(filename, sheet_name=i, header=6)
                df1 = df1.dropna(axis=0, how='all')
                df1 = df1.dropna(axis=1, how='all')
                df1 = df1.iloc[:29]

                df1['מילוי\nאו\nמערך חוזר'] = df1['מילוי\nאו\nמערך חוזר'].str.strip()
                df1['סוג יין'] = df1['סוג יין'].str.strip()

                df1['סוג\nקפסולות'] = df1['סוג\nקפסולות'].fillna('ללא')
                df1['סוג\nתויות'] = df1['סוג\nתויות'].fillna('ללא')
                df1['קרטון'] = df1['קרטון'].fillna('ללא')
                df1['בקבוק'] = df1['בקבוק'].fillna('ללא')
                df1['מילוי\nאו\nמערך חוזר'] = df1['מילוי\nאו\nמערך חוזר'].fillna('מילוי')

                pivot_plats = df1.pivot_table(
                    index=['סוג יין', 'סוג\nקפסולות', 'סוג\nתויות', 'קרטון', 'בקבוק', 'מדבקת\nקרטון', "סטרץ'\nמכונה",
                           "מילוי\nאו\nמערך חוזר"],
                    values=['כמות \nבקבוקים\nבמשטח'], aggfunc=['sum', 'count'])

                df = pd.DataFrame(pivot_plats.to_records())

                # pivot_plats
                df.rename(columns={df.columns[1]: 'קפסולות', df.columns[2]: 'תויות', df.columns[3]: 'אריזה',
                                   df.columns[4]: 'סוג בקבוק', df.columns[5]: 'מדבקת קרטון',
                                   df.columns[6]: 'משטחים לחיוב', df.columns[7]: 'מילוי או מערך חוזר',
                                   df.columns[8]: 'כמות בקבוקים', df.columns[9]: 'מספר משטחים'}, inplace=True)
                df['לתשלום'] = np.where(df['סוג יין'].str.contains("מגנום"), df["כמות בקבוקים"] * magnom,
                                        df["כמות בקבוקים"] * price)
                df['תאריך'] = i
                cols = df.columns.tolist()
                cols = cols[-1:] + cols[:-1]
                df = df[cols]
                df_dict[i] = df

            except Exception:
                error_message = f"Failed to process sheet '{i}'"
                continue

        df = pd.concat(df_dict.values(), ignore_index=True)

        df.insert(1, "יקב", Yekev_name)

        ####
        rating = []
        for row, val in zip(df["מדבקת קרטון"], df["כמות בקבוקים"]):
            if row == "לחיוב":
                rating.append(val / 6)
            else:
                rating.append(0)
        df['מדבקת קרטון לחיוב'] = rating

        ###

        strech = []
        for row, val in zip(df["משטחים לחיוב"], df["מספר משטחים"]):
            if row == "לחיוב":
                strech.append(val)
            else:
                strech.append(0)
        df["משטחים סטרץ' לחיוב"] = strech
        ###

        df['לתשלום'] = np.where(
            (~df["מילוי או מערך חוזר"].str.contains("מילוי")) & (~df['סוג יין'].str.contains("מגנום")),
            df["כמות בקבוקים"] * maarach, df['לתשלום'])

        ###

        df['מילוי'] = np.where((df["מילוי או מערך חוזר"].str.contains("מילוי")), df["כמות בקבוקים"], 0)
        df['מערך חוזר'] = np.where((~df["מילוי או מערך חוזר"].str.contains("מילוי")), df["כמות בקבוקים"], 0)
        ###

        df['הערות'] = np.where((df["מילוי או מערך חוזר"].str.contains("מילוי")), "", "")
        ###
        df.loc['Total'] = df.sum(numeric_only=True, axis=0)

        df.loc['לתשלום הכללי', 'לתשלום'] = df.loc['Total', 'לתשלום']
        df.loc['לתשלום על המשטחים', 'לתשלום'] = df.loc['Total', "משטחים סטרץ' לחיוב"] * mishtachim
        df.loc['לתשלום הכללי', 'מערך חוזר'] = "ביקבוק"
        df.loc['לתשלום על המשטחים', 'מערך חוזר'] = "סטרץ' מכונה"
        df.loc['מדבקות קרטון', 'לתשלום'] = df.loc['Total', "מדבקת קרטון לחיוב"] * karton
        df.loc['מדבקות קרטון', 'מערך חוזר'] = 'מדבקות קרטון'
        df.loc['סה"כ לפני מע"מ', 'לתשלום'] = df.loc['לתשלום הכללי', 'לתשלום'] + df.loc['לתשלום על המשטחים', 'לתשלום'] + \
                                             df.loc['מדבקות קרטון', 'לתשלום']
        df.loc['סה"כ לפני מע"מ', 'מערך חוזר'] = 'סה"כ לפני מע"מ'
        df.loc['מע"מ 17%', 'מערך חוזר'] = 'מע"מ 17%'
        df.loc['מע"מ 17%', 'לתשלום'] = df.loc['סה"כ לפני מע"מ', 'לתשלום'] * 0.17
        df.loc['לתשלום עד:', 'מערך חוזר'] = "לתשלום עד:"
        df.loc['לתשלום עד:', 'לתשלום'] = df.loc['סה"כ לפני מע"מ', 'לתשלום'] + df.loc['מע"מ 17%', 'לתשלום']
        df.loc['לתשלום עד:', 'מדבקת קרטון לחיוב'] = 'סה"כ לתשלום:'

        ###
        column_to_move = df.pop("כמות בקבוקים")
        column_to_move1 = df.pop("מילוי")
        column_to_move2 = df.pop("הערות")
        column_to_move3 = df.pop("מערך חוזר")
        column_to_move4 = df.pop("לתשלום")
        column_to_move5 = df.pop("מדבקת קרטון לחיוב")
        column_to_move6 = df.pop("משטחים סטרץ' לחיוב")
        df.insert(3, "כמות בקבוקים", column_to_move)
        df.insert(4, "מילוי", column_to_move1)
        df.insert(9, "הערות", column_to_move2)
        df.insert(10, "מערך חוזר", column_to_move3)
        df.insert(11, "לתשלום", column_to_move4)
        df.insert(12, "מדבקת קרטון לחיוב", column_to_move5)
        df.insert(13, "מדבקת סטרץ' לחיוב", column_to_move6)
        ###

        df['כמות יומית'] = np.where((df["מילוי או מערך חוזר"].str.contains("מילוי")), "", "")
        df['כמות מנות'] = np.where((df["מילוי או מערך חוזר"].str.contains("מילוי")), "", "")
        df['ספק מזון'] = np.where((df["מילוי או מערך חוזר"].str.contains("מילוי")), "", "")

        ###

        df = df.drop(columns=['מדבקת קרטון', 'משטחים לחיוב', 'מילוי או מערך חוזר', 'מספר משטחים', 'סוג בקבוק'])
        new_row = pd.DataFrame().reindex_like(df).iloc[:1]

        idx_pos = -6

        df = pd.concat([df.iloc[:idx_pos], new_row, df.iloc[idx_pos:]]).reset_index(drop=True)
        ##
        df.rename(columns={"מדבקת סטרץ' לחיוב": "משטחים"}, inplace=True)
        df.rename(columns={"מדבקת קרטון לחיוב": "מדבקת קרטון"}, inplace=True)
        ###
        output = io.BytesIO()

        writer = pd.ExcelWriter(f'{Yekev_name} {month} {datetime.now().year}.xlsx', engine='xlsxwriter')
        df.to_excel(writer, sheet_name=f'{Yekev_name} {month} {datetime.now().year}', index=False, startrow=5)
        workbook = writer.book
        worksheet = writer.sheets[f'{Yekev_name} {month} {datetime.now().year}']

        # Define the formats
        cell_format1 = workbook.add_format({'num_format': '₪ #,##0'})
        cell_format2 = workbook.add_format({'num_format': '#,##0'})
        cell_format3 = workbook.add_format({'bold': True, 'font_color': 'red', 'num_format': '₪ #.##0'})
        cell_format4 = workbook.add_format({'bold': True})
        cell_format5 = workbook.add_format({'bold': True, 'font_color': 'red', 'num_format': '₪ 0.###'})

        # Set the columns width and format
        worksheet.set_column('K:K', None, cell_format1)
        worksheet.set_column('E:E', None, cell_format2)
        worksheet.set_column('D:D', None, cell_format2)
        worksheet.set_column('L:L', None, cell_format2)

        if magnom > 0:
            worksheet.write('F4', "מגנום:", cell_format4)
            worksheet.write('E4', magnom, cell_format3)
        if maarach > 0:
            worksheet.write('J5', maarach, cell_format3)
        worksheet.write('F5', "רגיל:", cell_format4)
        worksheet.write('E5', price, cell_format3)
        worksheet.write('H1', f'ביקבוק יקב {Yekev_name}', cell_format4)
        worksheet.write('I2', 'חודש:', cell_format4)
        worksheet.write('H2', f'{month}', cell_format4)
        worksheet.write('G2', datetime.now().year, cell_format4)
        worksheet.write('D4', 'דוא"ל:', cell_format4)
        worksheet.write('D3', 'ח.פ', cell_format4)
        worksheet.write('D2', 'שם העסק:', cell_format4)
        worksheet.write('L5', karton, cell_format5)
        worksheet.write('M5', mishtachim, cell_format3)

        worksheet.autofit()

        # Write the file
        writer.save()
        output.seek(0)
        return writer

    output = recipt(filename, price, magnom, mishtachim, Yekev_name, karton, maarach, month)  # modify this line
    return send_file(output)



@app.route('/query1', methods=['POST'])
def query():

    # Get the winery name from the submitted form data
    winery_name = request.form['winery_name']

    # Read the data from an Excel file 'total_2015_to_2022.xlsx' into a pandas DataFrame
    data = pd.read_excel('total_2015_to_2022.xlsx')

    # Filter the data based on the specified winery name
    data = data[data['winery_name'] == winery_name]

    # Select all columns from the filtered data
    data = data.loc[:, :]

    # Get a list of all the unique winery names from the filtered data
    wineries = data['winery_name'].unique().tolist()


    # Convert the filtered data to a list of lists for rendering in the template
    data = data.values.tolist()

    # Pass the filtered data and winery list to the 'query.html' template and render it
    return render_template('query.html', data=data, wineries=wineries)


@app.route('/process_files', methods=['POST'])
def process_files():
    # get directory path from request data
    dir_path = '2023 update/'
    list1 = []
    for filename in os.listdir(dir_path):
        # read excel file
        df = pd.read_excel(os.path.join(dir_path, filename))

        # extract relevant information from dataframe
        winery_name = [df.iloc[5, 1].strip()]
        total_bootls = [df.iloc[df.iloc[:, 3].last_valid_index(), 3]]
        price = [df.iloc[3, 4]]
        year = [str(df.iloc[0, 7]) + ' ' + str(df.iloc[0, 6])]
        max_index = df.iloc[:, 2].last_valid_index()
        val = df.iloc[5:max_index + 1, 2].to_list()
        unique_values = list(set(val))
        bottles = [unique_values]
        max_index1 = df.iloc[:, 2].last_valid_index()
        val1 = df.iloc[5:max_index1 + 1, 7].to_list()
        karton = val1
        unique_values = list(set(karton))
        unique_karton = [unique_values]

        # create new dataframe with extracted information
        df = pd.DataFrame({'winery_name': winery_name, 'year': year, 'price': price, 'total_bottles': total_bootls,
                           'bottles': bottles, 'karton': unique_karton}).set_index('winery_name')
        list1.append(df)

    # concatenate dataframes
    concatenated_df = pd.concat(list1)
    concatenated_df['bottles'] = concatenated_df['bottles'].astype(str)
    concatenated_df['karton'] = concatenated_df['karton'].astype(str)
    concatenated_df['karton'] = concatenated_df['karton'].apply(lambda x: x.replace('[', ''))
    concatenated_df['karton'] = concatenated_df['karton'].apply(lambda x: x.replace(']', ''))
    concatenated_df['bottles'] = concatenated_df['bottles'].apply(lambda x: x.replace('[', ''))
    concatenated_df['bottles'] = concatenated_df['bottles'].apply(lambda x: x.replace(']', ''))
    concatenated_df['karton'] = concatenated_df['karton'].apply(lambda x: x.replace("'", ''))
    concatenated_df['bottles'] = concatenated_df['bottles'].apply(lambda x: x.replace("'", ''))

    # save dataframe to excel
    concatenated_df.to_excel("check.xlsx")

    # read in existing data and combine with new data, dropping duplicates
    df2 = pd.read_excel("total_2015_to_2022.xlsx")
    df1 = pd.read_excel("check.xlsx")

    df = pd.concat([df1, df2], axis=0).drop_duplicates(subset=['winery_name', 'year', 'price', 'total_bottles'])

    # save combined dataframe to excel
    df.to_excel("total_2015_to_2022.xlsx", index=False)

    return jsonify({'message': 'הנתונים התעדכנו בהצלחה!', 'button_text': 'Back to home', 'button_url': '/'})


if __name__ == '__main__':
    app.run(debug=True)
