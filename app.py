from flask import Flask, request, render_template, send_file, url_for, jsonify,flash
import pandas as pd
import os, sys
from calculator import hours
from invoice import recipt

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
    output = hours(filename, price_per_mile, month)
    return send_file(output)


@app.route('/result', methods=['POST'])
def result():
    """Generate invoice based on input data."""
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

    output = recipt(filename, price, magnom, mishtachim, Yekev_name, karton, maarach, month)
    return send_file(output)


@app.route('/query1', methods=['POST'])
def query():
    """Perform query based on winery name."""
    # Get the winery name from the submitted form data
    winery_name = request.form['winery_name']

    # Check if an "Other Winery" value was provided and use that as the winery name instead
    other_winery_name = request.form.get('other_winery')
    if other_winery_name:
        winery_name = other_winery_name

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
    """Process files for updating data."""
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
