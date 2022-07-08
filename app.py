# import dependencies
from fileinput import filename
import os
import pandas as pd
import glob
import uuid
from operator import add
from flask import Flask, render_template, request, send_from_directory
from flask_dropzone import Dropzone

basedir = os.path.abspath(os.path.dirname(__file__))

# each user gets a UUID when visiting/uploading
userid = uuid.uuid4()
user = str(userid)


# flask app init and configuration
app = Flask(__name__)
app.config.update(
    UPLOADED_PATH=os.path.join(basedir, 'uploads/' + user),
    DROPZONE_MAX_FILE_SIZE=10000,
    DROPZONE_TIMEOUT=60*60*1000,
    DROPZONE_ALLOWED_FILE_CUSTOM=True,
    DROPZONE_ALLOWED_FILE_TYPE='.xlsx',
    DROPZONE_INVALID_FILE_TYPE="Can't upload files of this type."
)

dropzone = Dropzone(app)


@app.route('/', methods=['POST', 'GET'])
def upload():
    if request.method == 'POST':
        # create user upload folder
        user_dir = 'Flask/ehs-inspection-combine/uploads/'+user
        if not os.path.exists(user_dir):
            os.makedirs(user_dir)
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

    # copy output file to user's folder and rename
    import shutil
    shutil.copy('Flask/ehs-inspection-combine/output/Pareto Output.xlsx',
                'Flask/ehs-inspection-combine/uploads/'+user+'/Pareto Output - '+user+'.xlsx')
    # assign user's output file (filepath string)
    output_file = 'Flask/ehs-inspection-combine/uploads/' + \
        user+'/Pareto Output - '+user+'.xlsx'

    # select output file and set Issue Counts column
    output_file_df = pd.read_excel(
        'Flask/ehs-inspection-combine/uploads/'+user+'/Pareto Output - '+user+'.xlsx')
    output_file_df['Issue Count'] = overall_counts

    # write dataframe to user output file
    with pd.ExcelWriter(output_file, mode='a', if_sheet_exists="replace") as writer:
        output_file_df.to_excel(writer, sheet_name='Data')

    message = "Counts totaled from " + \
        str(len(file_list)) + " files, click Download to download output file."
    return render_template('index.html', message=message)


# routing for output download
# @app.route('/download/', methods=['POST', 'GET'])
# def download():
#     message = 'Downloading output file.'
#     return send_file(
#         filename_or_fp='Flask/ehs-inspection-combine/uploads/'+user+'/Pareto Output - '+user+'.xlsx',
#         download_name='Pareto Output.xlsx',
#         as_attachment=True
#     ), render_template('index.html', message=message)


if __name__ == '__main__':
    app.run(debug=True)

# TODO
# download file on button click working
# clear uploads for user after output is downloaded
