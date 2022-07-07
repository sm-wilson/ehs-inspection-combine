# import dependencies
import os
import pandas as pd
import glob
import xlwings as xw
import openpyxl
from operator import add
from flask import Flask, render_template, request
from flask_dropzone import Dropzone

basedir = os.path.abspath(os.path.dirname(__file__))

# flask app init and configuration
app = Flask(__name__)
app.config.update(
    UPLOADED_PATH=os.path.join(basedir, 'uploads'),
    DROPZONE_MAX_FILE_SIZE=1024,
    DROPZONE_TIMEOUT=60*60*1000,
    DROPZONE_ALLOWED_FILE_CUSTOM=True,
    DROPZONE_ALLOWED_FILE_TYPE='.xlsx',
    DROPZONE_INVALID_FILE_TYPE="Can't upload files of this type."
)

dropzone = Dropzone(app)

# UUID when uploading


@app.route('/', methods=['POST', 'GET'])
def upload():
    if request.method == 'POST':
        f = request.files.get('file')
        f.save(os.path.join(app.config['UPLOADED_PATH'], f.filename))
    return render_template('index.html')


@app.route("/combine/", methods=['POST'])
def combine():
    # Combine code

    # pull Excel files from Uploads folder & store in file_list
    file_list = glob.glob(os.path.join(app.config['UPLOADED_PATH'], "*.xlsx"))

    overall_counts = [0, 0, 0, 0, 0, 0, 0, 0, 0,
                      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]

    for file_ in file_list:
        # read Pareto sheet from Excel file

        # Check to make sure all Excel files have the required Pareto sheet
        # i.e. make sure all files are the correct inspection sheets
        try:
            df = pd.read_excel(file_, sheet_name="Pareto")
        except ValueError:
            message = 'Incompatible Excel file found. All files must include "Pareto" tab with an "Incident Count" column heading.\n\
                    Please double check your Excel spreadsheets and try again.'
            return render_template('index.html', message=message)

        # set the column where values are
        counts_column = df['Issue Count']

        # init local list
        local_list = []

        # add each category count to the list
        for row in counts_column:
            local_list.append(row)

        # add local list values to the running total in overall_counts
        overall_counts = list(map(add, overall_counts, local_list))

    # select output file and set Issue Counts column
    output_file = pd.read_excel(
        'Flask/drag-drop-file-flask/output/Pareto Output.xlsx')
    output_file['Issue Count'] = overall_counts

    sheet_mapping = {
        'Data': output_file
    }
    # wb = openpyxl.load_workbook(
    #     'Flask/drag-drop-file-flask/output/Pareto Output.xlsx')
    # sheet = wb['Data']
    # print(sheet)

    # open Excel in background and open template file
    with xw.App(visible=False) as xlapp:
        wb = xlapp.books.open(output_file)

    # create list of current worksheet names
    current_sheets = [sheet.name for sheet in wb.sheets]

    # Iterate over sheet/df mapping
    # If sheet already exists, overwrite current content. Else, add new sheet
    for sheet_name in sheet_mapping.keys():
        if sheet_name in current_sheets:
            wb.sheets(sheet_name).range("A1").options(
                index=False).value = sheet_mapping.get(sheet_name)
        else:
            new_sheet = wb.sheets.add(after=wb.sheets.count)
            new_sheet.range("A1").options(
                index=False).value = sheet_mapping.get(sheet_name)
            new_sheet.name = sheet_name
    wb.save()

    message = "Counts totaled, click Download to download output file."
    return render_template('index.html', message=message)


if __name__ == '__main__':
    app.run(debug=True)

# TODO
# get writing to template working (pandas?)
# make sure files are unique to each user
# - (add user ID to folder/files to avoid users overwriting each others files)
# clear uploads for user after output is downloaded
