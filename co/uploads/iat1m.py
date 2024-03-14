from flask import Flask, render_template, request,render_template_string
from prettytable import PrettyTable
import pandas as pd
import xlwings as xw
import numpy as np
import openpyxl

iat1m = Flask(__name__)

@iat1m.route('/')
def index():
    return render_template('index.html')

@iat1m.route('/process_data', methods=['POST'])
def process_data():
    IAT1_uploaded = request.files['IAT1']
    IAT1_path = 'uploads/input_data.xlsx'
    IAT1_uploaded.save(IAT1_path)
    IAT1 = pd.read_excel(IAT1_path)

        #IAT1
    # Specify the common column you want to process and the number of columns to include
    C_C = 'CO1'
    num_columns_to_process = 4  # You want to process "CO4" and the next 3 columns

    # Find the index of the common column in the DataFrame's columns
    common_column_index = IAT1.columns.get_loc(C_C)
    co_attainment_dict = {}

    iat1_table = PrettyTable()
    # Extract the columns you want to process based on the index

    columns_to_p = IAT1.columns[common_column_index:common_column_index + num_columns_to_process]

    iat1_table.field_names = ["Column", "Total Students Passed (> 60%)", "Total Students Attempted", "CO Attainment", "Attainment Level"]

    for column in columns_to_p:
        # Convert the column to numeric, ignoring non-numeric values
    # Check if the column exists in the DataFrame
        if column in IAT1.columns:
         IAT1[column] = pd.to_numeric(IAT1[column], errors='coerce')

        # Count cells in the column with a value greater than or equal to 3
        count_greater_than_3 = (IAT1[column] >= 3).sum()

        # Count filled cells in the column
        filled_cells_count = IAT1[column].count()

        # Calculate CO_attainment
        CO_attainment = round((count_greater_than_3 / filled_cells_count) * 100)

        # Determine attainment level
        if CO_attainment >= 50 and CO_attainment < 60:
            attainment_Level = 1
        elif CO_attainment >= 60 and CO_attainment < 70:
            attainment_Level = 2
        elif CO_attainment >= 70:
            attainment_Level = 3
        else:
            attainment_Level = 0

        co_attainment_dict[column] = CO_attainment
        iat1_table.add_row([column, count_greater_than_3, filled_cells_count, CO_attainment, attainment_Level])
        # Print or store the results as needed
        # print(f'Column: {column},   Total student passed with > 60%:  {count_greater_than_3},   Total # students attempted the QUESTION:  {filled_cells_count},   CO_attainment: {CO_attainment},    attainment_Level: {attainment_Level}')

    second_column_index = IAT1.columns.get_loc(C_C)+4
    columns_to_next = IAT1.columns[second_column_index:second_column_index + (num_columns_to_process-2) ]
    for column in columns_to_next:
        # Convert the column to numeric, ignoring non-numeric values
    # Check if the column exists in the DataFrame
        if column in IAT1.columns:
         IAT1[column] = pd.to_numeric(IAT1[column], errors='coerce')

        # Count cells in the column with a value greater than or equal to 3
        count_greater_than_3 = (IAT1[column] >= 6).sum()

        # Count filled cells in the column
        filled_cells_count = IAT1[column].count()

        # Calculate CO_attainment
        CO_attainment = round((count_greater_than_3 / filled_cells_count) * 100)

        # Determine attainment level
        if CO_attainment >= 50 and CO_attainment < 60:
            attainment_Level = 1
        elif CO_attainment >= 60 and CO_attainment < 70:
            attainment_Level = 2
        elif CO_attainment >= 70:
            attainment_Level = 3
        else:
            attainment_Level = 0


        co_attainment_dict[column] = CO_attainment
        iat1_table.add_row([column, count_greater_than_3, filled_cells_count, CO_attainment, attainment_Level])

    # Convert PrettyTable to DataFrame
    iat1_table_df = pd.read_html(iat1_table.get_html_string())[0]

    # Transpose the DataFrame
    iat1_transposed_table = iat1_table_df.T

    # Print the transposed table

    new_column_values = ['', 'Total Students Passed (> 60%) ', 'Total Students Attempted', 'CO Attainment ','Attainment Level']  # Replace this with your actual values
    iat1_transposed_table.insert(0, "NewColumn", new_column_values)
    # Set the first row as the header
    iat1_transposed_table.columns = iat1_transposed_table.iloc[0]

    # Drop the first row (which is now the header) to avoid duplicate headers
    transposed_table = iat1_transposed_table[1:]
    print(transposed_table)

    v_variables = {}
    counter = 1
    avg_attainment_list = []

    # Create a dictionary to store average CO attainment values for columns with the same initial three characters
    average_co_attainment_dict = {}

    # Iterate through the columns to calculate average CO attainment
    for column_name, attainment in co_attainment_dict.items():
        initial_3_characters = column_name[:3]
        if initial_3_characters in average_co_attainment_dict:
            average_co_attainment_dict[initial_3_characters].append(attainment)
        else:
            average_co_attainment_dict[initial_3_characters] = [attainment]

    v_variables = {}
    counter = 1
    avg_attainment_list = []
    # Calculate the average CO attainment for each group of columns with the same initial three characters
    for group, attainments in average_co_attainment_dict.items():
        avg_attainment = sum(attainments) / len(attainments)
        avg_attainment_list.append(avg_attainment)



    iat1 = {
        'CO':['CO1','CO2','CO3','CO4','CO5','CO6'],
        'avg_attainment':avg_attainment_list
    }
    while len(avg_attainment_list) < len(iat1['CO']):
        avg_attainment_list.append(0)
    iat1_data = pd.DataFrame(iat1)
    print(iat1_data)

    iat1_html_table = iat1_data.to_html(index=False)

    # Return the results to the HTML template
    return render_template('result.html', result=iat1_html_table)
    # return render_template_string('<html><body>{{ html_table }}</body></html>', html_table=iat2_html_table)

if __name__ == '__main__':
    iat1m.run(debug=True)