from flask import Flask, render_template, request,render_template_string
from prettytable import PrettyTable
import pandas as pd
import xlwings as xw
import numpy as np
import openpyxl

indirectm = Flask(__name__)

@indirectm.route('/')
def index():
    return render_template('index.html')

@indirectm.route('/process_data', methods=['POST'])
def process_data():
    uploaded1 = request.files['Indirect']
    indirect_path = 'uploads/input_data.xlsx'
    uploaded1.save(indirect_path)
    indirect = pd.read_excel(indirect_path)


    #indirect
    column_name = '%'  # Replace with the actual column name

    # Read the specific column into a DataFrame
    indirect_df = pd.read_excel(indirect_path, usecols=[column_name])
    column_name = '%'  # Replace with the actual column name

    # Read the specific column into a DataFrame
    indirect1 = pd.read_excel(indirect_path, usecols=[column_name])
    transposed_indirect1 = indirect1.T

    transposed_indirect1 = transposed_indirect1.drop(transposed_indirect1.columns[0], axis=1)
    header_names = ["CO1", "CO2", "CO3",'CO4','CO5','CO6']  # Replace with actual header names
    transposed_indirect1.columns = header_names
    print(transposed_indirect1)

    indirect_html_table = transposed_indirect1.to_html(index=False)

    # Return the results to the HTML template
    return render_template('result.html', result=indirect_html_table)
    # return render_template_string('<html><body>{{ html_table }}</body></html>', html_table=iat2_html_table)

if __name__ == '__main__':
    indirectm.run(debug=True)
